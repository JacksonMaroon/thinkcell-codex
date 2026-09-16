"""Size existing native think-cell text through PowerPoint, regenerate and verify.

Series/point fills are rejected: a changed-data canary reverted those API edits.
No text, number formats, labels or chart features are inserted or deleted.
"""
from pathlib import Path
from types import SimpleNamespace
import argparse,json,sys,zipfile,subprocess,math,re
from lxml import etree as E
HERE=Path(__file__).resolve().parent;sys.path[:0]=[str(HERE),str(HERE/'thinkcell_no_click/implementation')]
import multi_chart_update as multi
from thinkcell import baseline_request
from office_operation_lock import serialized_office, run_locked_subprocess
from prepare_thinkcell_name import inventory,need,sha,logical_slides
from runtime import powershell,powershell_env
NS={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart','a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main'}

def visible(path,c):
 with zipfile.ZipFile(path) as z:r=E.fromstring(z.read(c['frames'][0]['native_chart_part']))
 return {str(i):s for i,s in enumerate(r.xpath('.//c:ser',namespaces=NS),1)}

def verify_features(path,plan):
 _,cs,_=inventory(path.read_bytes());bytag={c['frames'][0]['shape_tag']:c for c in cs}
 for target in plan['targets']:
  series=visible(path,bytag[target['shape_tag']])
  need(set(series)==set(target['visible_series_names']),'Visible series changed')
  for name,color in target.get('series_fills',{}).items():
   need(series[name].xpath('./c:spPr/a:solidFill/a:srgbClr/@val',namespaces=NS)==[color.upper()],'Series fill did not survive regeneration: '+name)
  for name,colors in target.get('point_fills',{}).items():
   points={int(p.find('c:idx',NS).get('val')):p for p in series[name].findall('c:dPt',NS)}
   for index,color in enumerate(colors):
    if color is not None:need(index in points and points[index].xpath('./c:spPr/a:solidFill/a:srgbClr/@val',namespaces=NS)==[color.upper()],'Point fill did not survive regeneration')
 if plan.get('native_text_font_size'):
  with zipfile.ZipFile(path) as z:
   root=E.fromstring(z.read(logical_slides(z)[0]['part']))
   sizes=root.xpath('.//p:sp[.//p:tags]//a:rPr/@sz',namespaces=NS)
   need(sizes and all(abs(float(s)/100-plan['native_text_font_size'])<0.01 for s in sizes),'Native text font size did not survive regeneration')
 return {'series_and_point_fills':'not_supported','native_text_font_size':'pass' if plan.get('native_text_font_size') else 'not_requested'}

@serialized_office
def run(a):
 src,out,report=map(lambda p:Path(p).resolve(),(a.input,a.output,a.report));plan=json.loads(Path(a.plan).read_text(encoding='utf-8-sig'))
 need(sha(src.read_bytes())==a.expected_sha256.upper(),'Input hash mismatch')
 need(len({src,out,report,Path(a.plan).resolve()})==4 and not out.exists() and not report.exists(),'Use distinct new output/report paths')
 need(set(plan)=={'targets','native_text_font_size'} and plan.get('targets'),'Specify all chart targets and native_text_font_size')
 _,cs,_=inventory(src.read_bytes());tagmap={c['frames'][0]['shape_tag']:c for c in cs}
 size=plan.get('native_text_font_size');need(not isinstance(size,bool) and isinstance(size,(int,float)) and math.isfinite(size) and 1<=size<=72,'Font size must be a finite number from 1 to 72')
 for t in plan['targets']:
  need(set(t)=={'shape_tag'} and t['shape_tag'] in tagmap,'Only existing native text sizing is supported; fill edits revert on data updates')
  series=visible(src,tagmap[t['shape_tag']]);t['visible_series_names']=list(series)
  for key in ('series_fills','point_fills'):
   t.setdefault(key,{})
   need(set(t[key])<=set(series),'Unknown visible series')
   for name,value in t[key].items():
    values=value if key=='point_fills' else [value]
    need(isinstance(values,list) and all(v is None or (isinstance(v,str) and re.fullmatch('[0-9A-Fa-f]{6}',v)) for v in values),'Invalid fill color')
    if key=='point_fills':
     count=int(series[name].find('c:val/c:numRef/c:numCache/c:ptCount',NS).get('val'));need(len(values)==count,'Point count mismatch')
 need(len({t['shape_tag'] for t in plan['targets']})==len(plan['targets']),'Duplicate target')
 need({t['shape_tag'] for t in plan['targets']}==set(tagmap),'Text sizing requires every chart on the slide to be explicit')
 source_baseline={'targets':[{'selector':{'slide_number':1,'shape_tag':c['frames'][0]['shape_tag']},'name':c['owner_name'],'data':baseline_request(src,c)} for c in cs]}
 multi.validate_plan(src.read_bytes(),source_baseline)
 stage=out.parent/(out.stem+'_feature_work');need(not stage.exists(),'Staging already exists');stage.mkdir()
 plan.update(input=str(src),input_sha256=sha(src.read_bytes()))
 pp=stage/'native-plan.json';pp.write_text(json.dumps(plan),encoding='utf-8');prepared=stage/'styled.pptx'
 with (stage/'native.log').open('w') as log:
  code=run_locked_subprocess([powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(HERE/'native_feature_api.ps1'),'-PlanFile',str(pp),'-OutputFile',str(prepared),'-ReportFile',str(stage/'api.json')],operation='native-fonts',timeout_seconds=180,stdout=log,stderr=log,env=powershell_env(),creationflags=subprocess.CREATE_NO_WINDOW).returncode
 need(code==0 and prepared.exists(),'Native API operation failed')
 _,after,_=inventory(prepared.read_bytes());baseline={'targets':[{'selector':{'slide_number':1,'shape_tag':c['frames'][0]['shape_tag']},'name':c['owner_name'],'data':baseline_request(src,tagmap[c['frames'][0]['shape_tag']])} for c in after]}
 bp=stage/'baseline.json';bp.write_text(json.dumps(baseline),encoding='utf-8');verified=stage/'verified.pptx'
 gates=multi.run(SimpleNamespace(input=prepared,expected_sha256=sha(prepared.read_bytes()),plan=bp,output=verified,report=stage/'gates.json',prepare=True,execute=True,ppttc=None));(stage/'gates.json').write_text(json.dumps(gates,indent=2))
 retained=verify_features(verified,plan);need(sha(src.read_bytes())==a.expected_sha256.upper(),'Source changed')
 with out.open('xb') as f:f.write(verified.read_bytes())
 result={'status':'ALL_GATES_PASS','feature_retention':retained,'data_native_gates':str(stage/'gates.json'),'source_unchanged':True,'output':str(out),'render':gates['render']};report.write_text(json.dumps(result,indent=2));return result

if __name__=='__main__':
 p=argparse.ArgumentParser(description=__doc__)
 for n in ('input','plan','output','report'):p.add_argument('--'+n,type=Path,required=True)
 p.add_argument('--expected-sha256',required=True);a=p.parse_args()
 print(json.dumps(run(a)))
