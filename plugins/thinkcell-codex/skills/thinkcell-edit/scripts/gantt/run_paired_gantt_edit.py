"""CLI for the portable paired Gantt date+geometry adapter."""
from __future__ import annotations
import argparse, hashlib, json
from pathlib import Path
from portable_gantt_adapter import paired_date_geometry_edit

def sha(p):
 h=hashlib.sha256();
 with Path(p).open('rb') as f:
  for b in iter(lambda:f.read(1024*1024),b''): h.update(b)
 return h.hexdigest()

def main():
 ap=argparse.ArgumentParser(); ap.add_argument('--source',required=True); ap.add_argument('--expected-sha256',required=True); ap.add_argument('--output',required=True); ap.add_argument('--select-start',required=True); ap.add_argument('--select-end',required=True); ap.add_argument('--start-date',required=True); ap.add_argument('--end-date',required=True); ap.add_argument('--report',required=True); a=ap.parse_args()
 src=Path(a.source); out=Path(a.output); report_path=Path(a.report); before=sha(src)
 if before.lower()!=a.expected_sha256.lower(): raise SystemExit('source SHA256 mismatch')
 if out.exists(): raise SystemExit(f'output already exists: {out}')
 if report_path.exists(): raise SystemExit(f'report already exists: {report_path}')
 edit=paired_date_geometry_edit(src,out,{'start':a.select_start,'end':a.select_end},a.start_date,a.end_date)
 result={'status':'PAIRED_GANTT_EDIT_PREPARED','source':str(src.resolve()),'output':str(out.resolve()),'source_sha256_before':before,'source_sha256_after':sha(src),'source_unchanged':before==sha(src),'edit':edit}
 report_path.write_text(json.dumps(result,indent=2)+'\n',encoding='utf8'); print(json.dumps(result,indent=2)); return 0
if __name__=='__main__': raise SystemExit(main())
