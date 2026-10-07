"""Native anchor topology guards, without Office or proprietary fixtures."""
from pathlib import Path
import io
import sys
import tempfile
import unittest
import zipfile
from unittest.mock import patch

SCRIPTS = Path(__file__).resolve().parents[1] / 'skills/thinkcell-edit/scripts'
sys.path[:0] = [str(SCRIPTS / 'gantt'), str(SCRIPTS)]
from dependency_adapter import anchor_edges, finish_start_model, inspect_native_anchors
import dependency_adapter


def model():
    return ('<root><CGanttSE id="1"/>'
            '<CGanttBar id="2"><m_anchorAttachedLow idref="0"/><m_anchorAttachedHigh idref="0"/></CGanttBar>'
            '<CGanttBar id="3"><m_anchorAttachedLow idref="0"/><m_anchorAttachedHigh idref="0"/></CGanttBar>'
            '<Unrelated id="4" value="preserve"/></root>')


class NativeGanttDependencyContracts(unittest.TestCase):
    def test_inspector_reads_actual_cfb_stream_instead_of_fragmented_raw_xml(self):
        stream = ('<root><version val="38790"/><CGanttSE id="1"><m_bConsistent val="1"/></CGanttSE>'
                  '<CGanttBar id="2"><m_varsrcLow><m_varval type="6"><m_datetime val="2026-01-05T00:00:00"/></m_varval><m_guid val="start"/></m_varsrcLow>'
                  '<m_varsrcHigh><m_varval type="6"><m_datetime val="2026-01-12T00:00:00"/></m_varval><m_guid val="end"/></m_varsrcHigh>'
                  '<m_pptgenline><m_pptautoshpline idref="3"/></m_pptgenline></CGanttBar>'
                  '<CPPTAutoShapeLine id="3"><m_bPlaced val="0"/><m_bstrShapeName>hidden</m_bstrShapeName></CPPTAutoShapeLine></root>')
        class Carrier:
            def __enter__(self): return self
            def __exit__(self, *_): return False
            def exists(self, name): return name == 'think-cellXML'
            def openstream(self, name): return io.BytesIO(stream.encode())
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'source.pptx'
            with zipfile.ZipFile(source, 'w') as archive:
                archive.writestr('ppt/embeddings/carrier.bin', b'<root><CGanttSE broken\x00XML</root>')
            with patch('olefile.OleFileIO', return_value=Carrier()):
                state = inspect_native_anchors(source)
            self.assertTrue(state['consistent'])
            self.assertEqual(state['schema_version'], '38790')
            self.assertEqual(state['bars'][0]['date_guids'], ['start', 'end'])
            self.assertFalse(state['bars'][0]['visible'])

    def test_reciprocal_finish_start_anchor_allocates_fresh_identity(self):
        candidate, identity = finish_start_model(model(), '2', '3')
        self.assertEqual(identity, '5')
        self.assertEqual(anchor_edges(candidate), [('2', 'High', '3', 'Low', '5')])
        self.assertIn('<Unrelated id="4" value="preserve"/>', candidate)

    def test_duplicate_identity_refused(self):
        with self.assertRaisesRegex(ValueError, 'duplicate native'):
            finish_start_model(model().replace('id="4"', 'id="3"'), '2', '3')

    def test_self_link_refused(self):
        with self.assertRaisesRegex(ValueError, 'distinct'):
            finish_start_model(model(), '2', '2')

    def test_double_attachment_refused(self):
        candidate, _ = finish_start_model(model(), '2', '3')
        candidate = candidate.replace('</root>', '<CGanttBar id="6"><m_anchorAttachedLow idref="0"/><m_anchorAttachedHigh idref="0"/></CGanttBar></root>')
        with self.assertRaisesRegex(ValueError, 'already attached'):
            finish_start_model(candidate, '6', '3')

    def test_missing_endpoint_refused(self):
        with self.assertRaisesRegex(ValueError, 'exactly once'):
            finish_start_model(model(), '2', '99')

    def test_dangling_anchor_refused(self):
        candidate, _ = finish_start_model(model(), '2', '3')
        candidate = candidate.replace('<elem idref="3"/>', '<elem idref="99"/>')
        with self.assertRaisesRegex(ValueError, 'child is missing'):
            anchor_edges(candidate)

    def test_nonreciprocal_attachment_refused(self):
        candidate, _ = finish_start_model(model(), '2', '3')
        candidate = candidate.replace('<m_anchorAttachedLow idref="5"/>', '<m_anchorAttachedLow idref="0"/>')
        with self.assertRaisesRegex(ValueError, 'reciprocal'):
            anchor_edges(candidate)

    def test_cycle_refused(self):
        candidate, _ = finish_start_model(model(), '2', '3')
        with self.assertRaisesRegex(ValueError, 'cyclic'):
            finish_start_model(candidate, '3', '2')

    def test_cli_path_alias_refused_before_preparation(self):
        with tempfile.TemporaryDirectory() as folder:
            src = Path(folder) / 'source.pptx'
            src.write_bytes(b'source')
            args = ['gantt', '--source', str(src), '--expected-sha256', 'unused',
                    '--output', str(Path(folder) / 'output.pptx'), '--report', str(src),
                    '--predecessor-bar', '2', '--successor-bar', '3']
            with patch.object(sys, 'argv', args), patch.object(dependency_adapter, 'prepare_finish_start') as prepare:
                with self.assertRaisesRegex(ValueError, 'paths must be distinct'):
                    dependency_adapter.main()
                prepare.assert_not_called()
            self.assertEqual(src.read_bytes(), b'source')

    def test_cli_stale_source_refused_before_preparation(self):
        with tempfile.TemporaryDirectory() as folder:
            src = Path(folder) / 'source.pptx'
            src.write_bytes(b'source')
            args = ['gantt', '--source', str(src), '--expected-sha256', '0' * 64,
                    '--output', str(Path(folder) / 'output.pptx'), '--report', str(Path(folder) / 'report.json'),
                    '--predecessor-bar', '2', '--successor-bar', '3']
            with patch.object(sys, 'argv', args), patch.object(dependency_adapter, 'prepare_finish_start') as prepare:
                with self.assertRaisesRegex(ValueError, 'SHA256 mismatch'):
                    dependency_adapter.main()
                prepare.assert_not_called()


if __name__ == '__main__':
    unittest.main()
