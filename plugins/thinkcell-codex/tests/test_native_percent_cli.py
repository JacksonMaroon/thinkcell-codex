"""Public CLI preflight and exact-target checks, without Office writes."""
import argparse
import json
from pathlib import Path
import sys
import tempfile
import unittest

SCRIPTS = Path(__file__).resolve().parents[1] / 'skills/thinkcell-edit/scripts'
sys.path.insert(0, str(SCRIPTS))
import thinkcell
from prepare_thinkcell_name import inventory

LAB = Path(__file__).resolve().parents[3] / 'research/native-percent-labels'


class NativePercentCliTests(unittest.TestCase):
    def create_args(self, filename, output, request=None):
        source = LAB / filename
        return argparse.Namespace(input=source, expected_sha256=thinkcell.hash_file(source),
            slide_number=1, output_directory=output, style_file=None, execute=False,
            native_percent=True, data_json=request)

    def test_create_requires_real_native_percent_donor_before_writes(self):
        with tempfile.TemporaryDirectory() as folder:
            out = Path(folder) / 'new'
            with self.assertRaisesRegex(ValueError, 'Donor does not have verified'):
                thinkcell.create(self.create_args('00-native-absolute.pptx', out))
            self.assertFalse(out.exists())

    def test_one_shot_create_validates_changed_data_without_writes(self):
        with tempfile.TemporaryDirectory() as folder:
            out = Path(folder) / 'new'
            a = self.create_args('02-controlled-percent.pptx', out)
            _, charts, _ = inventory(a.input.read_bytes())
            data = thinkcell.baseline_request(a.input, charts[0])
            data['matrix'][3][1] = 200
            data['expected_model']['series_values'][2][0] = 200
            a.data_json = Path(folder) / 'request.json'
            a.data_json.write_text(json.dumps(data), encoding='utf-8')
            result = thinkcell.create(a)
            self.assertEqual(result['status'], 'CREATE_PREFLIGHT_PASS')
            self.assertTrue(result['data_request_validated'])
            self.assertEqual(result['native_percentage_labels']['label_count'], 9)
            self.assertFalse(out.exists())
            data['expected_model']['series_values'][2][0] = 999
            a.data_json.write_text(json.dumps(data), encoding='utf-8')
            with self.assertRaises(ValueError):
                thinkcell.create(a)
            self.assertFalse(out.exists())

    def test_changed_tag_never_falls_back_to_first_chart(self):
        source = {'frames': [{'shape_tag': 'original'}]}
        with self.assertRaisesRegex(ValueError, 'identity did not survive'):
            thinkcell.percent_target([{'frames': [{'shape_tag': 'different'}]}], source)


if __name__ == '__main__':
    unittest.main()
