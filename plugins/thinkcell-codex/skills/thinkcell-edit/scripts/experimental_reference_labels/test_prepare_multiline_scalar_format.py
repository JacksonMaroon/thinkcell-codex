"""Static regression test for the exact multiline scalar-format profile."""
import argparse, hashlib, importlib.util, tempfile
from pathlib import Path
HERE=Path(__file__).resolve().parent; spec=importlib.util.spec_from_file_location('fmt',HERE/'prepare_multiline_scalar_format.py'); fmt=importlib.util.module_from_spec(spec); spec.loader.exec_module(fmt)
p=argparse.ArgumentParser(); p.add_argument('--fixture',required=True); a=p.parse_args(); fixture=Path(a.fixture); digest=hashlib.sha256(fixture.read_bytes()).hexdigest().upper()
with tempfile.TemporaryDirectory() as td:
    td=Path(td); report=fmt.prepare(fixture,digest,td/'out.pptx',td/'report.json')
    assert report['whole_model_reversal_pass'] and report['only_selected_carrier_changed']
    assert report['preserved_prefix']=='$' and report['new_suffix']=='B\nExpansion\nopportunity'
    try: fmt.prepare(fixture,'0'*64,td/'bad.pptx',td/'bad.json'); raise AssertionError('wrong SHA accepted')
    except Exception as e: assert 'SHA-256 mismatch' in str(e) and not (td/'bad.pptx').exists()
print('PASS multiline scalar-format static profile')
