"""Static regression for an existing unplaced, direct-precision scalar label."""
import argparse, hashlib, importlib.util, tempfile
from pathlib import Path
HERE=Path(__file__).resolve().parent; spec=importlib.util.spec_from_file_location('labels',HERE/'prepare_existing_labels.py'); labels=importlib.util.module_from_spec(spec); spec.loader.exec_module(labels)
p=argparse.ArgumentParser(); p.add_argument('--fixture',required=True); a=p.parse_args(); fixture=Path(a.fixture); digest=hashlib.sha256(fixture.read_bytes()).hexdigest().upper()
with tempfile.TemporaryDirectory() as td:
    td=Path(td); report=labels.prepare(fixture,td/'out.pptx',digest,'KTC_STACKCOL_ABS_01','MT Total Gross Profit',['Expansion opportunity'],False,False,False,td/'report.json')
    selected=report['selected']; assert len(selected)==1 and selected[0]['direct_precision'] and selected[0]['physical_shape_absent']
    assert selected[0]['text_variable_id'] is None and report['changed_zip_entries']==['ppt/embeddings/oleObject13.bin']
    assert report['target_source_values_unchanged']=={'454':'2.96999999999999992895E+01'}
print('PASS direct-precision scalar suppression profile')
