"""Offline regression guards for the bounded new-break route (no Office)."""
from pathlib import Path
import contextlib, copy, hashlib, io, json, tempfile, unittest

import experimental_axis_break_insert as prepare
import run_native_axis_break_insert as native

ROOT=Path(__file__).resolve().parents[1]
SOURCE=ROOT/'assets/feature-fixtures/axis-break-ordinary-38764.pptx'
DONOR=ROOT/'assets/feature-fixtures/axis-break-seed-38764.pptx'
PLAN=json.loads((ROOT/'references/new-axis-break-evidence/same-data.json').read_text())
SHA=lambda p:hashlib.sha256(p.read_bytes()).hexdigest().upper()

class Guards(unittest.TestCase):
    def test_rejects_wrong_source_before_write(self):
        with tempfile.TemporaryDirectory() as folder:
            output=Path(folder)/'candidate.pptx'
            with self.assertRaisesRegex(RuntimeError,'hash'):
                prepare.build(SOURCE,DONOR,'0'*64,SHA(DONOR),output,Path(folder)/'report.json')
            self.assertFalse(output.exists())

    def test_rejects_nonfinite_boolean_and_inconsistent_data(self):
        for value in [True,float('inf'),float('nan')]:
            plan=copy.deepcopy(PLAN)
            plan['targets'][0]['data']['expected_model']['series_values'][0][0]=value
            with self.assertRaises(RuntimeError): native.profile(plan)
        plan=copy.deepcopy(PLAN)
        plan['targets'][0]['data']['matrix'][2][1]+=1
        with self.assertRaisesRegex(RuntimeError,'disagree'): native.profile(plan)

    def test_seed_is_not_accepted_as_regenerated_feature(self):
        with tempfile.TemporaryDirectory() as folder:
            output=Path(folder)/'candidate.pptx'
            with contextlib.redirect_stdout(io.StringIO()):
                prepare.build(SOURCE,DONOR,SHA(SOURCE),SHA(DONOR),output,Path(folder)/'report.json')
            result=native.verify(output,PLAN,2)
            self.assertEqual(result['status'],'PREPARED_SEED_NATIVE_REGENERATION_REQUIRED')
            with self.assertRaises(RuntimeError): native.verify(output,PLAN,1)

if __name__=='__main__': unittest.main()
