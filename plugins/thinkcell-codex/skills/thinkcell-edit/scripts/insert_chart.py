"""Compose a new slide copy from a complete native donor and ordinary target content."""
from pathlib import Path
import argparse
import json
import math
import subprocess
import sys
import zipfile
from lxml import etree as E

HERE=Path(__file__).resolve().parent
sys.path.insert(0,str(HERE/'thinkcell_no_click/implementation'))
from prepare_thinkcell_name import inventory, need, xml, sha
from audit_thinkcell_integrity import NS, logical_slides, relationship_map
from thinkcell import baseline_request, hash_file
from update_thinkcell_json import validate_output, snapshot
from chart_semantics import kind
from runtime import powershell, powershell_env
from placement_plan import plan as placement
import chart_geometry as geometry


def dump(path, value):
    with Path(path).open('x',encoding='utf-8') as f:
        json.dump(value,f,indent=2,allow_nan=False)


def package_info(path, target=False):
    with zipfile.ZipFile(path) as z:
        slides=logical_slides(z);need(len(slides)==1,'Extract exactly one slide before composition')
        part=slides[0]['part'];root=xml(z.read(part));rels=relationship_map(z,part)
        prs=xml(z.read('ppt/presentation.xml'));sz=prs.find('p:sldSz',NS)
        layout=next(x['resolved'] for x in rels.values() if x['type'].endswith('/slideLayout'))
        master=next(x['resolved'] for x in relationship_map(z,layout).values() if x['type'].endswith('/slideMaster'))
        theme=next(x['resolved'] for x in relationship_map(z,master).values() if x['type'].endswith('/theme'))
        if target:
            need(root.find('p:timing',NS) is None and root.find('p:transition',NS) is None,'Animated/transition target needs a preservation-capable route')
            need(root.find('p:cSld/p:bg',NS) is None,'Custom slide background unsupported by this composition route')
            need(not root.findall('.//p:oleObj',NS),'Target OLE objects require a separate preservation route')
            need(not any(x['type'].endswith(('/comments','/commentAuthors','/slide')) for x in rels.values()),'Target comments or slide links unsupported')
            for ref in root.findall('.//p:tags',NS):
                tags=xml(z.read(rels[ref.get('{'+NS['r']+'}id')]['resolved']))
                need(not any('THINKCELL' in n.get('name','').upper() for n in tags),'Target contains think-cell ownership')
            for rel in rels.values():
                if rel['type'].endswith('/notesSlide'):
                    nr=xml(z.read(rel['resolved']))
                    for shape in nr.findall('p:cSld/p:spTree/*',NS):
                        if shape.tag in {'{'+NS['p']+'}nvGrpSpPr','{'+NS['p']+'}grpSpPr'}:continue
                        need(shape.find('.//p:ph',NS) is not None,'Custom note objects require a separate preservation route')
        return {'slide_id':slides[0]['id'],'size':[int(sz.get('cx')),int(sz.get('cy'))],
                'first_slide_number':int(prs.get('firstSlideNum','1')),
                'theme_sha256':sha(z.read(theme)),
                'theme_format_sha256':sha(E.tostring(xml(z.read(theme)).find('a:themeElements',NS),method='c14n'))}


def donor_type(path):
    _,cs,_=inventory(Path(path).read_bytes());need(len(cs)==1 and cs[0]['exact'],'Use a one-chart native donor')
    c=cs[0];k=kind(c)
    if k in {'waterfall','mekko-percent','mekko-units'}:return k
    if c['owner'].tag=='CPieChartSE':
        model=snapshot(Path(path),c['owner_name'])['target_model'] if c['owner_name'] else None
        if model is None:
            from update_thinkcell_json import model_of
            model=model_of(c)
        return 'doughnut' if float(model.get('hole_percent') or 0)>0 else 'pie'
    part=c['frames'][0]['native_chart_part'];need(part,'Unrecognized native chart type')
    with zipfile.ZipFile(path) as z:r=xml(z.read(part))
    if r.find('.//c:bubbleChart',NS) is not None:return 'bubble'
    if r.find('.//c:scatterChart',NS) is not None:return 'scatter'
    bar=r.find('.//c:barChart',NS);line=r.find('.//c:lineChart',NS)
    if bar is not None:
        if line is not None:
            axes={n.get('val') for n in r.findall('.//c:axId',NS)}
            return 'column-line-dual-axis' if len(axes)>2 else 'column-line'
        direction=bar.find('c:barDir',NS).get('val')
        group=bar.find('c:grouping',NS).get('val')
        prefix='bar' if direction=='bar' else 'column'
        if group=='percentStacked':return prefix+'-100'
        if group=='stacked':return prefix+'-stacked'
        return 'bar' if prefix=='bar' else 'column-clustered'
    if line is not None:return 'line'
    if r.find('.//c:areaChart',NS) is not None:return 'area'
    raise ValueError('Unsupported donor type')


def native(args, log_dir, timeout=180):
    """A timeout leaves the running native helper intact and records its PID."""
    log_dir=Path(log_dir);log_dir.mkdir()
    with (log_dir/'stdout.txt').open('w') as out,(log_dir/'stderr.txt').open('w') as err:
        p=subprocess.Popen(args,stdout=out,stderr=err,env=powershell_env(),creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        try:code=p.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            dump(log_dir/'running.json',{'pid':p.pid,'status':'TIMEOUT_INSPECT_BEFORE_RETRY'})
            raise RuntimeError(f'Native helper still running, PID {p.pid}; inspect logs before retrying')
    need(code==0,'Native helper failed; inspect '+str(log_dir))


def update_chart(source, request, output):
    _,cs,_=inventory(source.read_bytes());c=cs[0]
    native([sys.executable,str(HERE/'thinkcell.py'),'update','--input',str(source),
        '--expected-sha256',hash_file(source),'--slide-number','1','--shape-tag',c['frames'][0]['shape_tag'],
        '--data-json',str(request),'--output',str(output),'--execute'],output.with_suffix('.logs'),300)


def outside_pixels(before, after, frame, width):
    from PIL import Image, ImageChops, ImageDraw
    a,b=Image.open(before).convert('RGB'),Image.open(after).convert('RGB')
    need(a.size==b.size,'Native render sizes differ')
    delta=ImageChops.difference(a,b);scale=a.width/width
    box=(math.floor(frame['left']*scale)-2,math.floor(frame['top']*scale)-2,
         math.ceil((frame['left']+frame['width'])*scale)+2,math.ceil((frame['top']+frame['height'])*scale)+2)
    ImageDraw.Draw(delta).rectangle(box,fill=(0,0,0))
    need(delta.getbbox() is None,'Pixels outside the chart region changed; candidate withheld')
    return True


def run(a):
    target,donor=a.target.resolve(),a.donor.resolve();out=a.output_directory.resolve()
    need(not out.exists(),'Use a new output directory')
    need(not out.is_relative_to(HERE.parent),'Outputs must be outside the installed skill')
    need(hash_file(target)==a.target_sha256.upper() and hash_file(donor)==a.donor_sha256.upper(),'Source hash changed')
    ctx=json.loads(a.context.read_text(encoding='utf-8-sig'));brief=json.loads(a.brief.read_text(encoding='utf-8-sig'))
    need(ctx.get('status')=='LAYOUT_INSPECTED_REVIEW_REQUIRED' and ctx.get('source_unchanged') and ctx.get('other_presentations_unchanged'),'Require a successful native layout inspection')
    need(ctx['source_sha256'].upper()==a.target_sha256.upper(),'Layout snapshot belongs to another source')
    need('notes_font_runs' in ctx,'Inspect the source again to capture original note formatting')
    ti,di=package_info(target,True),package_info(donor)
    need(ti['size']==di['size'],'Donor and target slide dimensions must match')
    need(all(isinstance(ctx[k],(int,float)) and not isinstance(ctx[k],bool) and math.isfinite(ctx[k]) and
             abs(ctx[k]-extent/12700)<.001 for k,extent in zip(('width','height'),ti['size'])),
         'Layout snapshot dimensions do not match the source presentation')
    need(str(ti['slide_id'])==str(ctx['slide_id']),'Layout snapshot slide identity changed')
    need(ti['theme_format_sha256']==di['theme_format_sha256'],'Choose a donor with matching theme colors, fonts and effects')
    pp=placement(ctx,brief);actual_type=donor_type(donor)
    requested=pp['chart_type']
    need(requested==actual_type,
         'Chosen chart type does not match donor: '+actual_type)
    request=json.loads(a.data_json.read_text(encoding='utf-8-sig'))
    _,cs,_=inventory(donor.read_bytes())
    from update_thinkcell_json import run as dry_update
    dry_update(donor,a.donor_sha256,out/'preflight.pptx',request,slide_number=1,shape_tag=cs[0]['frames'][0]['shape_tag'],execute=False)
    if not a.execute:return {'status':'INSERT_PREFLIGHT_PASS','writes':False,'placement':pp,'donor_type':actual_type}
    out.mkdir();dump(out/'placement.json',pp);dump(out/'data.json',request)
    result={'status':'FAILED','source_target_sha256':a.target_sha256,'source_donor_sha256':a.donor_sha256}
    try:
        named=donor
        if not cs[0]['owner_name']:
            baseline=baseline_request(donor,cs[0]);dump(out/'baseline.json',baseline)
            named=out/'named.pptx';update_chart(donor,out/'baseline.json',named)
        gp=geometry.make_plan(named,pp['frame']);dump(out/'geometry-plan.json',gp)
        intermediate=out/'geometry.pptx';dump(out/'geometry-preparation.json',geometry.prepare(named,intermediate,gp))
        positioned=out/'positioned.pptx';update_chart(intermediate,out/'data.json',positioned)
        before_geometry=geometry.verify(positioned,gp)
        composition=out/'composition'
        native([powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(HERE/'compose_slide.ps1'),
                '-Target',str(target),'-Donor',str(positioned),'-TargetSha256',hash_file(target),
                '-DonorSha256',hash_file(positioned),'-FirstSlideNumber',str(ti['first_slide_number']),
                '-OutputDirectory',str(composition)],out/'composition.logs')
        cr=json.loads((composition/'composition.json').read_text(encoding='utf-8-sig'))
        need(cr['status']=='COMPOSED_REVIEW_REQUIRED' and cr['notes_unchanged'] and cr['sources_unchanged'] and cr['other_presentations_unchanged'],'Composition preservation check failed')
        need(cr.get('notes_font_runs')==ctx['notes_font_runs'],'Original speaker-note font runs changed')
        candidate=composition/'composed.pptx'
        checks=validate_output(candidate,candidate,gp['automation_name'],request)
        old,new=snapshot(positioned,gp['automation_name']),snapshot(candidate,gp['automation_name'])
        need(old['target_semantics']==new['target_semantics'],'Chart semantics changed during composition')
        need(old['target_native_frame']['shape_tag']==new['target_native_frame']['shape_tag'],'Native chart identity changed')
        final_geometry=geometry.verify(candidate,gp)
        preview=Path(ctx['preview']);need(preview.is_file(),'Original native preview is missing')
        need(hash_file(preview)==str(ctx.get('preview_sha256','')).upper(),'Native preview hash is missing or changed; inspect the source again')
        outside_pixels(preview,composition/'after.png',pp['frame'],ctx['width'])
        need(hash_file(target)==a.target_sha256.upper() and hash_file(donor)==a.donor_sha256.upper(),'Input changed during composition')
        final=out/'slide-with-chart.pptx'
        with final.open('xb') as f:f.write(candidate.read_bytes())
        result.update(status='NATIVE_COMPOSITION_VERIFIED_VISUAL_REVIEW_REQUIRED',output=str(final),
            output_sha256=hash_file(final),render=str(composition/'after.png'),placement=pp,
            native_geometry=final_geometry,integrity_scope=checks['integrity_scope'],
            original_content_pixels_unchanged=True,original_notes_text_unchanged=True,original_notes_font_runs_unchanged=True,
            sources_unchanged=True,exact_chart_data_pass=True,native_reopen_pass=True,
            scope='New one-slide copy; complete native donor chart retained, ordinary target content copied natively',
            review='Inspect chart choice, data labels, feature meaning and internal label overlap before delivery')
    except Exception as e:
        result['error']=str(e);dump(out/'result.json',result);raise
    dump(out/'result.json',result);return result


def main():
    p=argparse.ArgumentParser(description=__doc__)
    for key in ['target','donor','context','brief','data-json','output-directory']:
        p.add_argument('--'+key,type=Path,required=True)
    for key in ['target-sha256','donor-sha256']:p.add_argument('--'+key,required=True)
    p.add_argument('--execute',action='store_true');a=p.parse_args()
    try:print(json.dumps(run(a),allow_nan=False))
    except Exception as e:print(json.dumps({'status':'REJECTED','error':str(e)}),file=sys.stderr);return 1
    return 0


if __name__=='__main__':sys.exit(main())
