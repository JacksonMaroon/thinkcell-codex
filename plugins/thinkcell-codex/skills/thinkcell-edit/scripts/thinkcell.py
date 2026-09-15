"""Compact command interface for copy-only native think-cell updates."""
from pathlib import Path
import argparse
import hashlib
import json
import os
import subprocess
import sys
import time

HERE = Path(__file__).resolve().parent
IMPL = HERE/'thinkcell_no_click/implementation'
sys.path.insert(0, str(IMPL))
from runtime import doctor, find_ppttc, powershell, powershell_env

def write_json(path, data):
    with Path(path).open('x', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False, allow_nan=False)

def hash_file(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest().upper()

def select(candidates, a):
    from prepare_thinkcell_name import choose
    return choose(candidates, a.slide_id, a.slide_number, a.shape_id, a.shape_tag)

def baseline_request(path, c):
    from prepare_thinkcell_name import link_contract, xml
    from update_thinkcell_json import model_of, canonical_sequence_contract, validate_request
    from read_named_datasheet import read_blob
    from extract_thinkcell_named_datasheet import wrap_biff_workbook_stream
    ds = link_contract(c)
    key = ds['storage']
    streams = c['doc']['streams']
    if (key,'Package') in streams:
        sheets = read_blob(streams[(key,'Package')], 'xlsb_package')
    elif (key,'Workbook') in streams:
        sheets = read_blob(wrap_biff_workbook_stream(streams[(key,'Workbook')]), 'legacy_biff_cfb')
    else:
        raise ValueError('Unsupported embedded datasheet.')
    if len(sheets) != 1 or not sheets[0]['nonempty_cells']:
        raise ValueError('Expected one nonempty embedded datasheet.')
    cells = sheets[0]['nonempty_cells']
    matrix = [[None]*max(x['column'] for x in cells) for _ in range(max(x['row'] for x in cells))]
    for x in cells:
        matrix[x['row']-1][x['column']-1] = x['value']
    model = model_of(c)
    family = c['owner'].tag
    fields = {'CPieChartSE':['categories','values'], 'CSequenceChartSE':['categories','series_names','series_values'],
              'CScatterChartSE':['point_labels','group_labels','x_values','y_values','size_values']}[family]
    expected = {k:model[k] for k in fields}
    if 'column_widths' in model:
        expected['column_widths']=model['column_widths']
    candidates = [matrix, [list(r) for r in zip(*matrix)]]
    # Source row layout is authoritative. The canonical contract chooses the correct orientation.
    if family == 'CSequenceChartSE':
        orientation = xml(streams[(key,'think-cellXML')]).find('PersistentType/m_eorient')
        if orientation is None or orientation.get('val') not in {'0','1'}:
            raise ValueError('Unknown source datasheet orientation.')
        candidates = [candidates[int(orientation.get('val'))]]
    for rows in candidates:
        request = {'matrix':rows, 'expected_model':dict(expected)}
        if family=='CSequenceChartSE' and 'column_widths' not in model and len(rows)>1 and rows[1][0] not in model['series_names'] and all(isinstance(v,(int,float)) for v in rows[1][1:]):
            from chart_semantics import kind
            if kind(c)=='CSequenceChartSE':request['expected_model']['category_extents']=model['category_extents']
        try:
            validate_request(request, family)
            target = {'slide_id':c['doc']['slide_id'], **c['frames'][0]}
            canonical_sequence_contract(path, request, target=target)
            if family == 'CPieChartSE':
                if len(rows) != len(model['categories'])+1 or any(len(row)!=2 for row in rows) or [row[0] for row in rows[1:]] != model['categories'] or [row[1] for row in rows[1:]] != model['values']:
                    continue
            return request
        except ValueError:
            continue
    raise ValueError('Could not derive an unambiguous canonical data request; inspect the datasheet.')

def inspect(a):
    from prepare_thinkcell_name import inventory, link_contract
    from update_thinkcell_json import model_of
    path = a.input.resolve()
    data = path.read_bytes()
    _, charts, _ = inventory(data)
    rows = []
    for c in charts:
        if a.slide_number is not None and c['doc']['slide_number'] != a.slide_number:
            continue
        if a.slide_id is not None and c['doc']['slide_id'] != a.slide_id:
            continue
        if a.shape_id is not None and not any(f['shape_id'] == a.shape_id for f in c['frames']):
            continue
        if a.shape_tag is not None and a.shape_tag not in c['tags']:
            continue
        row = {'slide_number':c['doc']['slide_number'], 'slide_id':c['doc']['slide_id'],
               'family':c['owner'].tag, 'automation_name':c['owner_name'] or None, 'exact_target':c['exact'], 'frames':c['frames']}
        try:
            link_contract(c)
            row['data_route'] = 'internal'
            from chart_semantics import details,feature_summary
            row['semantics']=details(c)
            row['existing_feature_counts_on_slide']=feature_summary(c)
        except ValueError as e:
            row['data_route'] = 'blocked'
            row['reason'] = str(e)
        if a.data:
            row['model'] = model_of(c)
        rows.append(row)
    result = {'source_sha256':hashlib.sha256(data).hexdigest().upper(), 'charts':rows, 'read_only':True}
    if a.request_out:
        req = baseline_request(path, select(charts,a))
        a.request_out.parent.mkdir(parents=True, exist_ok=True)
        write_json(a.request_out, req)
        result['request_template'] = str(a.request_out.resolve())
    if path.read_bytes() != data:
        raise ValueError('Source changed during inspection; inspect again.')
    return result

def update(a):
    from prepare_thinkcell_name import inventory, link_contract, need, SKILL
    from update_thinkcell_json import run, validate_output
    start = time.perf_counter()
    src = a.input.resolve()
    out = a.output.resolve()
    report_path = a.report.resolve() if a.report else out.with_suffix('.report.json')
    request_path = a.data_json.resolve()
    work = out.parent/(out.stem+'.work')
    need(src != out and out.suffix.lower() == '.pptx', 'Choose a distinct .pptx output.')
    need(not str(out).startswith('\\\\'), 'Use a local output path.')
    need(not out.is_relative_to(SKILL) and not report_path.is_relative_to(SKILL), 'Outputs must be outside the installed skill.')
    need(report_path not in {src,out,request_path} and report_path.suffix.lower() == '.json', 'Report must be a separate JSON file.')
    need(not report_path.is_relative_to(work), 'Report must be outside the internal working folder.')
    need(not out.exists() and not report_path.exists(), 'Output or report already exists; choose new paths.')
    need(out.parent.is_dir() and report_path.parent.is_dir(), 'Output and report folders must already exist.')
    need(hash_file(src) == a.expected_sha256.upper(), 'Source hash changed; inspect again.')
    _, charts, _ = inventory(src.read_bytes())
    target = select(charts,a)
    if a.named_only:
        need(bool(target['owner_name']), 'Named-only mode requires an existing automation name.')
    for c in charts:
        link_contract(c)  # Includes siblings: regeneration must not silently break their links.
    request = json.loads(request_path.read_text(encoding='utf-8-sig'))
    generated = work/'generated.pptx'
    if not a.execute:
        dry = run(src,a.expected_sha256,generated,request,a.slide_id,a.slide_number,a.shape_id,a.shape_tag,False)
        return {'status':'DRY_RUN_PASS','automation_name':dry['preparation']['automation_name'],
                'experimental_naming_needed':not dry['preparation']['reused_existing_name'], 'output':str(out),
                'data_contract':dry['json_layout_contract'], 'writes':False}
    need(os.name == 'nt', 'Execution requires Windows desktop PowerPoint and licensed think-cell.')
    exe = Path(a.ppttc).resolve() if a.ppttc else find_ppttc()
    need(exe is not None and exe.is_file(), 'Run doctor; ppttc.exe was not found.')
    need(not work.exists(), 'Working folder already exists; inspect it before retrying with a new output name.')
    work.mkdir()
    result = {'status':'STARTED', 'source_sha256':a.expected_sha256.upper(), 'output':str(out)}
    try:
        generation = run(src,a.expected_sha256,generated,request,a.slide_id,a.slide_number,a.shape_id,a.shape_tag,True,exe)
        write_json(work/'generation.json', generation)
        verified = work/'verified.pptx'
        native_report = work/'native.json'
        render = work/'preview.png'
        cmd = [powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(IMPL/'native_verify_scoped.ps1'),
               '-InputFile',str(generated),'-OutputFile',str(verified),'-ReportFile',str(native_report),'-RenderFile',str(render)]
        with (work/'native.stdout.txt').open('w') as stdout, (work/'native.stderr.txt').open('w') as stderr:
            process = subprocess.Popen(cmd,stdout=stdout,stderr=stderr,creationflags=subprocess.CREATE_NO_WINDOW,env=powershell_env())
            try:
                code = process.wait(timeout=180)
            except subprocess.TimeoutExpired:
                raise RuntimeError(f'Native verification still running, PID {process.pid}. No process was killed. Inspect work files before retrying.')
        need(code == 0 and native_report.is_file(), 'Native verification failed; inspect work/native.json and logs.')
        native = json.loads(native_report.read_text(encoding='utf-8-sig'))
        need(native['native_reopen_pass'] and native['other_presentations_unchanged'] and native['source_unchanged'], 'Native verification did not pass.')
        prepared = Path(generation['staging_directory'])/'prepared.pptx'
        validation = validate_output(prepared,verified,generation['preparation']['automation_name'],request)
        need(hash_file(src) == a.expected_sha256.upper(), 'Source changed during update; candidate not delivered.')
        with out.open('xb') as f:
            f.write(verified.read_bytes())
        result.update(status=('NATIVE_DATA_VERIFIED_SPECIALIZED_VISUAL_REVIEW_REQUIRED' if validation['integrity_scope']['specialized_visual_semantics_review_required'] else 'NATIVE_VERIFIED_VISUAL_REVIEW_REQUIRED'),output_sha256=hash_file(out),
                      source_unchanged=True,automation_name=generation['preparation']['automation_name'],
                      experimental_namer_used=not generation['preparation']['reused_existing_name'],
                      integrity_scope=validation['integrity_scope'],
                      render=str(render),work_directory=str(work),exact_data_pass=True,native_reopen_pass=True,
                      visual_review='Inspect the preview for clipping, labels and layout before delivery.',
                      elapsed_seconds=round(time.perf_counter()-start,2))
        write_json(work/'post-native-validation.json', validation)
        write_json(report_path,result)
        return result
    except Exception as e:
        result.update(status='FAILED', error=str(e), work_directory=str(work), elapsed_seconds=round(time.perf_counter()-start,2))
        write_json(report_path,result)
        raise

def selectors(p):
    p.add_argument('--slide-number',type=int)
    p.add_argument('--slide-id',type=int)
    p.add_argument('--shape-id',type=int)
    p.add_argument('--shape-tag')

def create(a):
    """Create a new native slide from a user-supplied donor, preserving features."""
    import zipfile
    from prepare_thinkcell_name import logical_slides, need
    source=a.input.resolve();destination=a.output_directory.resolve()
    need(not destination.exists(), 'Use a new output directory.')
    need(not destination.is_relative_to(HERE.parent), 'Output must be outside the installed skill.')
    need(hash_file(source)==a.expected_sha256.upper(), 'Source hash changed; inspect again.')
    with zipfile.ZipFile(source) as z:
        need(1<=a.slide_number<=len(logical_slides(z)), 'Slide number out of range.')
    if a.style_file:
        need(a.style_file.is_file(), 'Style file does not exist.')
    if not a.execute:
        return {'status':'CREATE_PREFLIGHT_PASS','writes':False,'method':'native_whole_slide_donor_copy',
                'data_update':'Inspect the new slide, then update its exact chart.',
                'style_scope':'Optional defaults for new elements; donor formatting is preserved.'}
    cmd=[powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(HERE/'extract_slide.ps1'),
         '-InputFile',str(source),'-SlideNumber',str(a.slide_number),'-OutputDirectory',str(destination)]
    proc=subprocess.run(cmd,capture_output=True,text=True,env=powershell_env(),timeout=240)
    need(proc.returncode==0, 'Native donor copy failed: '+proc.stdout[-1500:]+proc.stderr[-500:])
    result=json.loads((destination/'extraction.json').read_text(encoding='utf-8-sig'))
    if a.style_file:
        style_out=destination/'style'
        cmd=[powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(HERE/'load_style.ps1'),
             '-InputFile',result['output'],'-OutputDirectory',str(style_out),'-StyleFile',str(a.style_file.resolve())]
        proc=subprocess.run(cmd,capture_output=True,text=True,env=powershell_env(),timeout=180)
        need(proc.returncode==0, 'Style defaults failed: '+proc.stdout[-1500:]+proc.stderr[-500:])
        result['style']=json.loads((style_out/'style.json').read_text(encoding='utf-8-sig'))
        result['output']=result['style']['output']
    need(hash_file(source)==a.expected_sha256.upper(), 'Source changed during creation.')
    result.update(status='CREATED_FROM_DONOR_REVIEW_REQUIRED',output_sha256=hash_file(Path(result['output'])),
                  next_step='Inspect the created file, update chart data, then review the native preview.')
    return result

def verify(a):
    from update_thinkcell_json import validate_output
    request = json.loads(a.data_json.read_text(encoding='utf-8-sig'))
    result = validate_output(a.prepared,a.input,a.automation_name,request)
    return {'status':'EXACT_DATA_AND_NATIVE_STRUCTURE_PASS','integrity_scope':result['integrity_scope'], 'output_sha256':result['output_sha256'],
            'target_tag_preserved':result['resolved_target_tag_preserved'],
            'untargeted_chart_data_unchanged':result['untargeted_chart_data_and_streams_unchanged'],
            'theme_and_notes_unchanged':result['bound_theme_and_notes_unchanged'],
            'read_only':True, 'native_reopen_and_visual_review':'Separate checks; this command does not open Office.'}

def main():
    p = argparse.ArgumentParser(description=__doc__)
    sub = p.add_subparsers(dest='command',required=True)
    sub.add_parser('doctor',help='Read-only runtime and dependency discovery.')
    c = sub.add_parser('create',help='Create a new native slide from a donor; preserves its chart types and features.')
    c.add_argument('--input',type=Path,required=True)
    c.add_argument('--expected-sha256',required=True)
    c.add_argument('--slide-number',type=int,required=True)
    c.add_argument('--output-directory',type=Path,required=True)
    c.add_argument('--style-file',type=Path)
    c.add_argument('--execute',action='store_true')
    i = sub.add_parser('inspect',help='List charts; optionally write a canonical request template.')
    i.add_argument('--input',type=Path,required=True)
    i.add_argument('--data',action='store_true')
    i.add_argument('--request-out',type=Path)
    selectors(i)
    u = sub.add_parser('update',help='Prepare and validate by default; --execute creates a natively verified copy.')
    for arg in ('input','output','data-json'):
        u.add_argument('--'+arg,type=Path,required=True)
    u.add_argument('--expected-sha256',required=True)
    u.add_argument('--report',type=Path)
    u.add_argument('--ppttc',help='Optional installed ppttc.exe path.')
    u.add_argument('--named-only',action='store_true',help='Disable experimental naming; require an existing name.')
    u.add_argument('--execute',action='store_true')
    selectors(u)
    v = sub.add_parser('verify',help='Check an exported or reopened candidate against its prepared source and intended data.')
    v.add_argument('--prepared',type=Path,required=True)
    v.add_argument('--input',type=Path,required=True)
    v.add_argument('--data-json',type=Path,required=True)
    v.add_argument('--automation-name',required=True)
    a = p.parse_args()
    try:
        result = {'doctor':lambda:doctor(),'create':lambda:create(a),'inspect':lambda:inspect(a),'update':lambda:update(a),'verify':lambda:verify(a)}[a.command]()
        print(json.dumps(result,ensure_ascii=False,allow_nan=False))
    except (Exception, SystemExit) as e:
        print(json.dumps({'status':'REJECTED','error':str(e)}),file=sys.stderr)
        return 1
    return 0

if __name__ == '__main__':
    sys.exit(main())
