"""External-link contracts using neutral synthetic XML and workbook fixtures."""
from pathlib import Path
import base64
import json
import sys
import struct
import tempfile
import unittest
from unittest.mock import patch
import zipfile

SCRIPTS = Path(__file__).resolve().parents[1] / 'skills/thinkcell-edit/scripts'
sys.path.insert(0, str(SCRIPTS / 'external_links'))
sys.path.insert(0, str(SCRIPTS))
from extract_thinkcell_named_datasheet import wrap_biff_workbook_stream
import portable_rebind as rebind
import batch_rebind as batch


def carrier(path=r'C:\source.xlsx', padded=True):
    # Deliberately opaque prefix/suffix. No serialized moniker format is assumed.
    payload = b'prefix' + path.encode('utf-16-le') + b'\0\0' + b'suffix'
    encoded = base64.b64encode(payload)
    if not padded:
        encoded = encoded.rstrip(b'=')
    xml = (b'<root><m_advisesink idref="7"/><m_bExternalStorage val="1"/>'
           b'<m_bstrRangeName val="range"/><m_bNeedsUpdateFromSheetOnMakeTC val="0"/>'
           b'<CAdviseSink id="7"><m_guidLink val="guid"/><m_lnkid>link</m_lnkid>'
           b'<m_bAutoUpdate val="1"/><m_vecbMoniker>' + encoded +
           b'</m_vecbMoniker></CAdviseSink><untouched val="yes"/></root>')
    return xml, payload


def workbook(path, reference="'Sales'!$A$1:$B$4", name='range', local_id='0'):
    with zipfile.ZipFile(path, 'w') as package:
        package.writestr('xl/workbook.xml',
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheets><sheet name="Sales"/></sheets><definedNames>'
            f'<definedName name="{name}" localSheetId="{local_id}">{reference}</definedName>'
            '</definedNames></workbook>')


def synthetic_ole(xml):
    # Reuse the public neutral CFB fixture wrapper and name its one stream.
    raw = bytearray(wrap_biff_workbook_stream(struct.pack('<HHHH', 0x0809, 4, 0x0600, 5)))
    raw[512:4608] = xml.ljust(4096, b' ')
    directory_sector = struct.unpack_from('<I', raw, 48)[0]
    entry = 512 + directory_sector * 512 + 128
    name = 'think-cellXML\0'.encode('utf-16-le')
    raw[entry:entry + 64] = name.ljust(64, b'\0')
    struct.pack_into('<H', raw, entry + 64, len(name))
    return bytes(raw)


class MonikerTests(unittest.TestCase):
    def test_only_workbook_path_changes(self):
        xml, payload = carrier()
        changed, old = rebind.rebind_xml(xml, payload, r'C:\target.xlsx')
        self.assertEqual(old, r'C:\source.xlsx')
        self.assertEqual(changed.replace(base64.b64encode(payload.replace(
            'source'.encode('utf-16-le'), 'target'.encode('utf-16-le'))),
            base64.b64encode(payload)), xml)
        self.assertEqual(rebind._fields(changed)[0]['guid_link'], 'guid')

    def test_unpadded_base64_style_preserved(self):
        xml, payload = carrier(padded=False)
        changed, _ = rebind.rebind_xml(xml, payload, r'C:\target.xlsx')
        self.assertEqual(len(changed), len(xml))
        self.assertEqual(rebind._fields(changed)[2], payload.replace(
            'source'.encode('utf-16-le'), 'target'.encode('utf-16-le')))

    def test_utf16_byte_length_is_authority(self):
        xml, payload = carrier('C:\\😀.xlsx')
        changed, _ = rebind.rebind_xml(xml, payload, 'C:\\AB.xlsx')
        self.assertEqual(rebind._path_runs(rebind._fields(changed)[2])[0][2], 'C:\\AB.xlsx')

    def test_arbitrary_length_is_rejected(self):
        xml, payload = carrier()
        with self.assertRaisesRegex(RuntimeError, 'UTF16-byte-length'):
            rebind.rebind_xml(xml, payload, r'C:\longer-target.xlsx')

    def test_invalid_target_paths(self):
        xml, payload = carrier()
        for target in ('relative.xlsx', r'\\server\file.xlsx', 'C:\\x\0.xlsx', r'C:\target.txt'):
            with self.subTest(target=target), self.assertRaisesRegex(RuntimeError, 'absolute Windows'):
                rebind.rebind_xml(xml, payload, target)

    def test_case_insensitive_noop_is_rejected(self):
        xml, payload = carrier()
        with self.assertRaisesRegex(RuntimeError, 'no-op'):
            rebind.rebind_xml(xml, payload, r'c:\SOURCE.xlsx')

    def test_payload_must_match_sealed_xml(self):
        xml, payload = carrier()
        with self.assertRaisesRegex(RuntimeError, 'differs from XML'):
            rebind.rebind_xml(xml, payload + b'changed', r'C:\target.xlsx')

    def test_duplicate_moniker_is_rejected(self):
        xml, payload = carrier()
        xml = xml.replace(b'</root>', b'<m_vecbMoniker>AA==</m_vecbMoniker></root>')
        with self.assertRaisesRegex(RuntimeError, 'serialization'):
            rebind.rebind_xml(xml, payload, r'C:\target.xlsx')

    def test_base64_noise_is_rejected(self):
        with self.assertRaises(ValueError):
            rebind._moniker('%%%%')

    def test_unterminated_and_extension_prefix_paths_are_rejected(self):
        self.assertEqual(rebind._path_runs(r'C:\source.xlsx'.encode('utf-16-le')), [])
        payload = 'C:\\source.xlsx.backup\0'.encode('utf-16-le')
        self.assertEqual(rebind._path_runs(payload), [])

    def test_native_item_moniker_boundary_is_accepted(self):
        path = r'C:\source.xlsx'
        payload = b'prefix' + path.encode('utf-16-le') + rebind.ITEM_MONIKER_CLSID + b'opaque'
        self.assertEqual([item[2] for item in rebind._path_runs(payload)], [path])
        altered = payload.replace(rebind.ITEM_MONIKER_CLSID, b'\x05' + rebind.ITEM_MONIKER_CLSID[1:])
        self.assertEqual(rebind._path_runs(altered), [])

    def test_multiple_workbook_paths_are_rejected(self):
        xml, payload = carrier()
        payload += ('C:\\other.xlsx\0').encode('utf-16-le')
        with self.assertRaisesRegex(RuntimeError, 'exactly one absolute Windows'):
            rebind.rebind_xml(xml, payload, r'C:\target.xlsx')

    def test_advise_sink_identity_must_resolve(self):
        xml, _ = carrier()
        with self.assertRaisesRegex(ValueError, 'advise-sink identity'):
            rebind._fields(xml.replace(b'idref="7"', b'idref="8"'))


class DiscoveryTests(unittest.TestCase):
    def test_optional_guards_select_unique_carrier(self):
        xml, _ = carrier()
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'neutral.pptx'
            with zipfile.ZipFile(path, 'w') as package:
                package.writestr('ppt/embeddings/oleObject1.bin', b'one')
                package.writestr('ppt/embeddings/oleObject2.bin', b'two')
            def read(raw):
                return xml if raw == b'one' else xml.replace(b'val="guid"', b'val="other"')
            with patch.object(rebind, '_xml', side_effect=read):
                with self.assertRaisesRegex(RuntimeError, 'found 2'):
                    rebind.discover(path)
                self.assertEqual(rebind.discover(path, guid='other')['part'],
                                 'ppt/embeddings/oleObject2.bin')
                with self.assertRaisesRegex(RuntimeError, 'found 0'):
                    rebind.discover(path, guid='other', link_id='wrong')

    def test_malformed_external_carrier_is_not_silently_skipped(self):
        xml, _ = carrier()
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'neutral.pptx'
            with zipfile.ZipFile(path, 'w') as package:
                package.writestr('ppt/embeddings/oleObject1.bin', b'one')
            with patch.object(rebind, '_xml', return_value=xml.replace(b'idref="7"', b'idref="8"')):
                with self.assertRaisesRegex(RuntimeError, 'malformed linked carrier'):
                    rebind.discover(path)


class NamedRangeTests(unittest.TestCase):
    def test_contiguous_local_range(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'neutral.xlsx'
            workbook(path)
            self.assertEqual(rebind.named_range(path, 'range', 'Sales')['range'], 'A1:B4')

    def test_formula_union_external_and_wrong_sheet_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'neutral.xlsx'
            for reference in ("'Sales'!A1,B2", "'Sales'!OFFSET(A1,1,0)",
                              "'[other.xlsx]Sales'!A1", "'Other'!A1"):
                with self.subTest(reference=reference):
                    workbook(path, reference)
                    with self.assertRaises(RuntimeError):
                        rebind.named_range(path, 'range', 'Sales')


class BatchTests(unittest.TestCase):
    def setup_batch(self, folder):
        base = Path(folder)
        src, out, manifest, report = [base / name for name in ('in.pptx', 'out.pptx', 'plan.json', 'report.json')]
        with zipfile.ZipFile(src, 'w') as package:
            package.writestr('ppt/embeddings/one.bin', b'one')
            package.writestr('ppt/embeddings/two.bin', b'two')
            package.writestr('ppt/slides/slide1.xml', b'unchanged')
        source, target = base / 'source.xlsx', base / 'target.xlsx'
        source.write_bytes(b'source'); target.write_bytes(b'target')
        links = [dict(source_workbook=str(source), source_workbook_sha256=rebind.sha256(source),
                      target_workbook=str(target), target_workbook_sha256=rebind.sha256(target),
                      guid=guid) for guid in ('one', 'two')]
        manifest.write_text(json.dumps({'schema': 'thinkcell-batch-rebind-v1', 'links': links}))
        return src, out, manifest, report

    def fake_rebind(self, argv):
        args = dict(zip(argv[::2], argv[1::2]))
        part = 'ppt/embeddings/' + args['--guid'] + '.bin'
        with zipfile.ZipFile(args['--input-presentation']) as before:
            contents = {name: before.read(name) for name in before.namelist()}
        contents[part] += b'-rebound'
        with zipfile.ZipFile(args['--output-presentation'], 'w') as after:
            for name, value in contents.items():
                after.writestr(name, value)
        fields = ('source_workbook', 'source_workbook_sha256', 'target_workbook', 'target_workbook_sha256',
                  'old_moniker_workbook_path', 'new_moniker_workbook_path', 'identity_before',
                  'identity_after', 'guards')
        evidence = {key: {} for key in fields}
        evidence['selected_part'] = part
        Path(args['--report']).write_text(json.dumps(evidence))
        return 0

    def test_batch_changes_only_two_selected_parts(self):
        with tempfile.TemporaryDirectory() as folder:
            src, out, manifest, report = self.setup_batch(folder)
            digest = rebind.sha256(src)
            with patch.object(batch, 'rebind_main', side_effect=self.fake_rebind):
                result = batch.run(src, out, manifest, report, digest)
            self.assertEqual(result['changed_parts'], ['ppt/embeddings/one.bin', 'ppt/embeddings/two.bin'])
            self.assertEqual(result['output_sha256'], rebind.sha256(out))
            self.assertEqual(rebind.sha256(src), digest)
            self.assertTrue(report.exists())

    def test_second_request_failure_publishes_nothing(self):
        with tempfile.TemporaryDirectory() as folder:
            src, out, manifest, report = self.setup_batch(folder)
            def fail_second(argv):
                if argv[-1] == 'two':
                    raise RuntimeError('second request failed')
                return self.fake_rebind(argv)
            with patch.object(batch, 'rebind_main', side_effect=fail_second):
                with self.assertRaisesRegex(RuntimeError, 'second request'):
                    batch.run(src, out, manifest, report, rebind.sha256(src))
            self.assertFalse(out.exists()); self.assertFalse(report.exists())

    def test_duplicate_selected_carrier_publishes_nothing(self):
        with tempfile.TemporaryDirectory() as folder:
            src, out, manifest, report = self.setup_batch(folder)
            plan = json.loads(manifest.read_text()); plan['links'][1]['guid'] = 'one'
            manifest.write_text(json.dumps(plan))
            with patch.object(batch, 'rebind_main', side_effect=self.fake_rebind):
                with self.assertRaisesRegex(ValueError, 'distinct linked carriers'):
                    batch.run(src, out, manifest, report, rebind.sha256(src))
            self.assertFalse(out.exists()); self.assertFalse(report.exists())

    def test_manifest_requires_identity_guard(self):
        with tempfile.TemporaryDirectory() as folder:
            _, _, manifest, _ = self.setup_batch(folder)
            plan = json.loads(manifest.read_text()); del plan['links'][0]['guid']
            manifest.write_text(json.dumps(plan))
            with self.assertRaisesRegex(ValueError, 'identity selector'):
                batch.load_requests(manifest)

    def test_sealed_input_alias_and_wrong_hash_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            src, out, manifest, report = self.setup_batch(folder)
            with self.assertRaisesRegex(ValueError, 'distinct from all sealed'):
                batch.run(src, manifest, manifest, report, rebind.sha256(src))
            with self.assertRaisesRegex(ValueError, 'SHA-256 mismatch'):
                batch.run(src, out, manifest, report, '0' * 64)
            self.assertFalse(out.exists()); self.assertFalse(report.exists())


class OfflinePackageTests(unittest.TestCase):
    def test_real_ole_write_package_readback_and_source_guard(self):
        with tempfile.TemporaryDirectory() as folder:
            base = Path(folder)
            source, target = base / 'source.xlsx', base / 'target.xlsx'
            workbook(source); workbook(target)
            xml, old_payload = carrier(str(source))
            payload = old_payload + 'Sales!range\0'.encode('utf-16-le')
            xml = xml.replace(base64.b64encode(old_payload), base64.b64encode(payload))
            src, out, report = base / 'in.pptx', base / 'out.pptx', base / 'report.json'
            raw = synthetic_ole(xml)
            with zipfile.ZipFile(src, 'w') as package:
                package.writestr('ppt/embeddings/one.bin', raw)
                package.writestr('ppt/slides/slide1.xml', b'unchanged')
            argv = ['--input-presentation', str(src), '--output-presentation', str(out),
                    '--target-workbook', str(target), '--source-workbook', str(source),
                    '--input-sha256', rebind.sha256(src), '--source-workbook-sha256', rebind.sha256(source),
                    '--target-workbook-sha256', rebind.sha256(target), '--report', str(report)]
            with patch('sys.stdout', new=__import__('io').StringIO()):
                self.assertEqual(rebind.main(argv), 0)
            result = json.loads(report.read_text())
            self.assertEqual(result['changed_parts'], ['ppt/embeddings/one.bin'])
            self.assertEqual(rebind._path_runs(rebind.discover(out)['payload'])[0][2], str(target))
            self.assertTrue(all(result['preservation_after_readback'].values()))
            wrong_source = base / 'wrong.xlsx'; workbook(wrong_source)
            argv[argv.index('--source-workbook') + 1] = str(wrong_source)
            argv[argv.index('--source-workbook-sha256') + 1] = rebind.sha256(wrong_source)
            argv[argv.index('--output-presentation') + 1] = str(base / 'wrong-out.pptx')
            argv[argv.index('--report') + 1] = str(base / 'wrong-report.json')
            with self.assertRaisesRegex(RuntimeError, 'sealed link moniker'):
                rebind.main(argv)
            self.assertFalse((base / 'wrong-out.pptx').exists())


if __name__ == '__main__':
    unittest.main()
