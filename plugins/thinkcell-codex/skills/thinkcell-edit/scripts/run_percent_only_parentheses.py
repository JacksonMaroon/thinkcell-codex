"""Bounded native conversion for the tested dual-label sequence topology."""
from __future__ import annotations
import argparse, hashlib, json, subprocess, sys
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
HERE=Path(__file__).resolve().parent
sys.path[:0]=[str(HERE),str(HERE/'thinkcell_no_click'/'implementation')]
from office_operation_lock import OfficeOperationLock, run_locked_subprocess
from runtime import powershell, powershell_env
from multi_chart_update import run as official_run
from percent_conversion_guards import need, ratio, saved_ratio, dual, check as checked_field, write_new

TAG='tbUpiE_yCia4NQkLBVWFCIA'; CARRIER='ppt/embeddings/oleObject13.bin'; SCALAR='191'
def sha(p): return hashlib.sha256(Path(p).read_bytes()).hexdigest().upper()
def cmd(args, operation):
 p=run_locked_subprocess(args,operation=operation,timeout_seconds=240,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,env=powershell_env(),creationflags=subprocess.CREATE_NO_WINDOW)
 if p.returncode: raise RuntimeError((p.stderr or p.stdout)[-1200:])
def trace(path,out):
 cmd([sys.executable,str(HERE/'trace_percent_semantics.py'),'--input',str(path),'--carrier',CARRIER,'--scalar',SCALAR,'--shape-tag',TAG,'--out',str(out)],'percent-trace')
 return json.loads(out.read_text(encoding='utf8'))
def main(a):
 src=Path(a.input); interim=Path(a.work_dir)/'com-parentheses.pptx'; com_report=Path(a.work_dir)/'com-report.json'; static=Path(a.work_dir)/'com-trace.json'; final=Path(a.output); report=Path(a.report); final_trace=Path(a.work_dir)/'final-trace.json'
 pretrace=Path(a.work_dir)/'preflight-trace.json'; staged=Path(a.work_dir)/'verified-generated.pptx'; internal_report=Path(a.work_dir)/'official-report.json'; paths=[interim,com_report,static,final,report,final_trace,pretrace,staged,internal_report,Path(a.evidence)];
 if not Path(a.work_dir).is_dir() or any(p.parent!=Path(a.work_dir) and p not in (final,report,Path(a.evidence)) for p in [interim,com_report,static,final_trace]): raise ValueError('work paths must stay inside existing work directory')
 if len({p.resolve() for p in paths+[src,Path(a.plan)]})!=len(paths)+2 or any(p.exists() for p in paths): raise ValueError('all source/plan/output paths must be distinct and outputs new')
 if sha(src)!=a.expected_sha256.upper(): raise ValueError('source SHA-256 mismatch')
 requested=json.loads(Path(a.plan).read_text(encoding='utf-8-sig'))
 need(len(requested['targets'])==1,'Exactly one explicit chart required')
 numerator,denominator,derived=ratio(requested['targets'][0]['data']['matrix'])
 if a.expected_display!=derived: raise ValueError(f'expected display must match tested ROUND_HALF_UP ratio {numerator}/{denominator}: {derived}')
 before=trace(src,pretrace);current=saved_ratio(src,before);absolute_text,relative_text=dual(before)
 need(Decimal(absolute_text)==current[0] and relative_text==current[2],'Original dual label disagrees with saved data')
 with OfficeOperationLock('percent-only-parentheses',metadata={'input':str(src),'tag':TAG}):
  cmd([powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(HERE/'native_delete_absolute_keep_parentheses.ps1'),'-InputFile',str(src),'-OutputFile',str(interim),'-ReportFile',str(com_report),'-ExpectedSha256',a.expected_sha256,'-AbsoluteText',absolute_text,'-RelativeText',relative_text],'percent-native-com-partial-delete')
  scope=json.loads(com_report.read_text(encoding='utf-8-sig'))
  need(scope.get('source_unchanged') and scope.get('other_presentations_unchanged'),'COM preservation check failed')
  after=trace(interim,static);checked_field(after,relative_text,'post-COM')
  need(after['visible_shapes'][0]['fields'][0]['id']==before['visible_shapes'][0]['fields'][1]['id'],'Original relative field was replaced')
  need(saved_ratio(interim,after)==current,'Conversion changed raw data')
  official=official_run(argparse.Namespace(input=interim,expected_sha256=sha(interim),plan=Path(a.plan),output=staged,report=internal_report,prepare=True,execute=True,ppttc=None))
  write_new(internal_report,json.dumps(official,indent=2).encode())
 final_state=trace(staged,final_trace);checked_field(final_state,derived,'post-native',True)
 actual=saved_ratio(staged,final_state);need(actual==(numerator,denominator,derived),'Saved final ratio differs from requested ratio')
 if sha(src)!=a.expected_sha256.upper(): raise RuntimeError('source changed during conversion')
 write_new(final,staged.read_bytes());official['delivered_output']=str(final);official['percent_feature_gate']='PASS'
 write_new(report,json.dumps(official,indent=2).encode())
 result={'status':'PERCENT_ONLY_PARENTHESES_NATIVE_PASS','input_sha256':a.expected_sha256.upper(),'output_sha256':sha(final),'shape_tag':TAG,'scalar_id':SCALAR,'actual_native_ratio':f'{actual[0]}/{actual[1]}','expected_display':derived,'com_report':str(com_report),'official_report':str(report),'final_trace':str(final_trace),'absolute_raw_retained_unbound':True,'one_native_relative_field':True,'source_unchanged':True,'output_withheld_until_feature_pass':True}
 write_new(a.evidence,json.dumps(result,indent=2).encode());print(json.dumps(result,indent=2))
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--input',required=True);p.add_argument('--expected-sha256',required=True);p.add_argument('--plan',required=True);p.add_argument('--expected-display',required=True);p.add_argument('--work-dir',required=True);p.add_argument('--output',required=True);p.add_argument('--report',required=True);p.add_argument('--evidence',required=True);main(p.parse_args())
