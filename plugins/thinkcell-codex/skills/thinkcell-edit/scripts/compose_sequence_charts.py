"""Append a compatible native sequence-chart donor, then regenerate and verify.

Default is read-only preflight. --execute releases a distinct output only after
official JSON, exact-data and native save/reopen gates.
"""
from pathlib import Path
from types import SimpleNamespace
import argparse, contextlib, hashlib, importlib.util, io, json, sys, zipfile
HERE=Path(__file__).resolve().parent
sys.path[:0]=[str(HERE),str(HERE/'thinkcell_no_click/implementation')]
from office_operation_lock import serialized_office, run_locked_subprocess
from prepare_thinkcell_name import inventory, logical_slides, link_contract, need
from chart_semantics import kind
from thinkcell import baseline_request
import multi_chart_update
spec=importlib.util.spec_from_file_location('tc_graph_merge_core',HERE/'experimental_composition/merge_sequence_graphs.py')
core=importlib.util.module_from_spec(spec);spec.loader.exec_module(core)
def digest(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest().upper()
def preflight(receiver,donor):
    rb,db=receiver.read_bytes(),donor.read_bytes()
    rd,rc,_=inventory(rb);dd,dc,_=inventory(db)
    for raw in (rb,db):
        with zipfile.ZipFile(io.BytesIO(raw)) as z:need(len(logical_slides(z))==1,'Use one-slide input copies')
    need(len(rd)==len(dd)==1 and rc and len(dc)==1,'Require one receiver carrier and one donor chart')
    need(rd[0]['root'].get('reqver')==dd[0]['root'].get('reqver'),'Regenerate inputs to matching model versions')
    for c in rc+dc:
        need(c['exact'] and c['owner'].tag=='CSequenceChartSE' and kind(c) in {'CSequenceChartSE','waterfall'},'This route supports ordinary sequences and one fixed-topology waterfall')
        core.active(c,c is dc[0]);link_contract(c)
    need(sum(kind(c)=='waterfall' for c in rc+dc)<=1,'At most one fixed-topology waterfall is supported')
    return {'receiver_chart_count':len(rc),'donor_chart_count':len(dc)}
@serialized_office
def run(a):
    receiver,donor,out,report=(Path(x).resolve() for x in (a.receiver,a.donor,a.output,a.report))
    need(receiver.is_file() and donor.is_file(),'Inputs must be saved native files')
    need(out.suffix.lower()=='.pptx' and report.suffix.lower()=='.json','Use .pptx output and .json report')
    need(out not in {receiver,donor} and report not in {receiver,donor,out},'Input/output/report paths must differ')
    need(not out.exists() and not report.exists(),'Choose new output and report paths')
    need(out.parent.is_dir() and report.parent.is_dir(),'Output folders must already exist')
    before=(digest(receiver),digest(donor))
    need(before==(a.receiver_sha256.upper(),a.donor_sha256.upper()),'Source hash changed')
    facts=preflight(receiver,donor)
    result={'status':'DRY_RUN_PASS','source_unchanged':True,'document_writes':False,**facts}
    if not a.execute:return result
    stage=out.parent/(out.stem+'_composition_work');need(not stage.exists(),'Choose a fresh staging path')
    stage.mkdir();raw=stage/'raw-candidate.pptx'
    with contextlib.redirect_stdout(io.StringIO()) as captured:core.merge(receiver,donor,raw)
    (stage/'composition-log.json').write_text(captured.getvalue(),encoding='utf-8')
    _,charts,_=inventory(raw.read_bytes());plan={'targets':[]}
    for c in charts:
        plan['targets'].append({'selector':{'slide_number':1,'shape_tag':c['frames'][0]['shape_tag']},'name':c['owner_name'],'data':baseline_request(raw,c)})
    plan_path=stage/'all-chart-plan.json';plan_path.write_text(json.dumps(plan,indent=2),encoding='utf-8')
    verified=stage/'verified-native.pptx';gates_path=stage/'all-chart-gates.json'
    gates=multi_chart_update.run(SimpleNamespace(input=raw,expected_sha256=digest(raw),plan=plan_path,output=verified,report=gates_path,prepare=True,execute=True,ppttc=a.ppttc))
    gates_path.write_text(json.dumps(gates,indent=2),encoding='utf-8')
    need(gates.get('status')=='ALL_GATES_PASS' and verified.is_file(),'Native composition gates did not pass')
    need(before==(digest(receiver),digest(donor)),'A source changed during composition')
    with out.open('xb') as f:f.write(verified.read_bytes())
    result.update(status='ALL_GATES_PASS',output=str(out),output_sha256=digest(out),document_writes=True,staging_directory=str(stage),all_chart_gate_report=str(gates_path),render=gates['render'],visual_review_required=True)
    return result
def main():
    p=argparse.ArgumentParser(description=__doc__)
    for name in ('receiver','donor','output','report'):p.add_argument('--'+name,type=Path,required=True)
    p.add_argument('--receiver-sha256',required=True);p.add_argument('--donor-sha256',required=True)
    p.add_argument('--execute',action='store_true');p.add_argument('--ppttc');a=p.parse_args()
    try:
        result=run(a)
        with a.report.open('x',encoding='utf-8') as f:json.dump(result,f,indent=2)
        print(json.dumps(result));return 0
    except Exception as e:
        result={'status':'REJECTED','error':str(e),'output_withheld':not a.output.exists()}
        if a.report.resolve() not in {a.receiver.resolve(),a.donor.resolve(),a.output.resolve()} and not a.report.exists() and a.report.parent.is_dir():
            with a.report.open('x',encoding='utf-8') as f:json.dump(result,f,indent=2)
        print(json.dumps(result),file=sys.stderr);return 1
if __name__=='__main__':raise SystemExit(main())
