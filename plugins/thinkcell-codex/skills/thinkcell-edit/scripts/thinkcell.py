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
from office_operation_lock import serialized_office, run_locked_subprocess
from runtime import doctor, find_ppttc, powershell, powershell_env

def write_json(path, data):
    with Path(path).open('x', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False, allow_nan=False)

def hash_file(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest().upper()

def percent_contract():
    from percent_labels import native_percent_contract
    return native_percent_contract

def percent_target(charts, source_target):
    """Resolve a regenerated chart through the original exact frame tag."""
    from prepare_thinkcell_name import need
    frames = source_target.get('frames', [])
    need(len(frames) == 1 and frames[0].get('shape_tag'), 'Percentage target requires one exact frame tag.')
    tag = frames[0]['shape_tag']
    matches = [c for c in charts if any(f.get('shape_tag') == tag for f in c.get('frames', []))]
    need(len(matches) == 1, 'Native percentage target identity did not survive regeneration.')
    return matches[0]

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
        # A 100%-axis donor can encode its reserved row as all blanks.  The
        # model still exposes derived category extents, which belongs in the
        # canonical expected model even though the matrix row stays blank.
        if family=='CSequenceChartSE' and 'column_widths' not in model and len(rows)>1 and rows[1][0] not in model['series_names'] and (all(isinstance(v,(int,float)) for v in rows[1][1:]) or all(v is None for v in rows[1][1:])):
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
        fill_setting = c['table'].find('m_bExcelOnTop')
        row['datasheet_fill_enabled'] = fill_setting is not None and fill_setting.get('val') == '1'
        if row['datasheet_fill_enabled']:
            row['update_route'] = 'multi_chart_update.py preserves configured fills; cover every native chart'
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
        percent = percent_contract().snapshot(path, c)
        row['native_percentage_labels'] = {k:percent[k] for k in ('status', 'label_count')}
        from percent_labels import native_relative_contract
        relative = native_relative_contract.snapshot(path,c)
        row['native_relative_labels'] = {'status':relative['status'],'label_count':len(relative['labels'])}
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

@serialized_office
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
    percentage_before = percent_contract().snapshot(src, target)
    from percent_labels import native_relative_contract
    relative_before = native_relative_contract.snapshot(src, target)
    if getattr(a, 'require_native_percent', False):
        need(percentage_before['status'] == 'native_percentage', 'Target has no verified native percentage labels; select a native percentage donor.')
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
                'data_contract':dry['json_layout_contract'], 'native_percentage_labels':percentage_before, 'writes':False}
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
            code = run_locked_subprocess(cmd,operation='native-verify',timeout_seconds=180,stdout=stdout,stderr=stderr,creationflags=subprocess.CREATE_NO_WINDOW,env=powershell_env()).returncode
        need(code == 0 and native_report.is_file(), 'Native verification failed; inspect work/native.json and logs.')
        native = json.loads(native_report.read_text(encoding='utf-8-sig'))
        need(native['native_reopen_pass'] and native['other_presentations_unchanged'] and native['source_unchanged'], 'Native verification did not pass.')
        prepared = Path(generation['staging_directory'])/'prepared.pptx'
        validation = validate_output(prepared,verified,generation['preparation']['automation_name'],request)
        if percentage_before['status'] == 'native_percentage':
            _, after_charts, _ = inventory(verified.read_bytes())
            percentage_result = percent_contract().verify(percentage_before, verified, percent_target(after_charts, target))
        else:
            percentage_result = percentage_before
        if relative_before['status'] == 'native_relative_fields':
            _, relative_charts, _ = inventory(verified.read_bytes())
            relative_result = native_relative_contract.verify(relative_before, verified, percent_target(relative_charts, target))
        else:
            relative_result = relative_before
        need(hash_file(src) == a.expected_sha256.upper(), 'Source changed during update; candidate not delivered.')
        with out.open('xb') as f:
            f.write(verified.read_bytes())
        result.update(status=('NATIVE_DATA_VERIFIED_SPECIALIZED_VISUAL_REVIEW_REQUIRED' if validation['integrity_scope']['specialized_visual_semantics_review_required'] else 'NATIVE_VERIFIED_VISUAL_REVIEW_REQUIRED'),output_sha256=hash_file(out),
                      source_unchanged=True,automation_name=generation['preparation']['automation_name'],
                      experimental_namer_used=not generation['preparation']['reused_existing_name'],
                      integrity_scope=validation['integrity_scope'],
                      native_percentage_labels=percentage_result,
                      native_relative_labels=relative_result,
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

@serialized_office
def create(a):
    """Create a new native slide from a user-supplied donor, preserving features."""
    import zipfile
    from prepare_thinkcell_name import logical_slides, need, inventory
    source=a.input.resolve();destination=a.output_directory.resolve()
    need(not destination.exists(), 'Use a new output directory.')
    need(not destination.is_relative_to(HERE.parent), 'Output must be outside the installed skill.')
    need(hash_file(source)==a.expected_sha256.upper(), 'Source hash changed; inspect again.')
    with zipfile.ZipFile(source) as z:
        need(1<=a.slide_number<=len(logical_slides(z)), 'Slide number out of range.')
    if a.style_file:
        need(a.style_file.is_file(), 'Style file does not exist.')
    chart = None
    percentage_before = None
    data_path = getattr(a, 'data_json', None)
    if getattr(a, 'native_percent', False) or data_path:
        _, charts, _ = inventory(source.read_bytes())
        matches = [c for c in charts if c['doc']['slide_number'] == a.slide_number]
        need(len(matches) == 1, 'Feature creation requires exactly one chart on the donor slide.')
        chart = matches[0]
        percentage_before = percent_contract().snapshot(source, chart)
        if getattr(a, 'native_percent', False):
            need(percentage_before['status'] == 'native_percentage', 'Donor does not have verified native percentage labels.')
        if data_path:
            # Validate the complete request and link contract before opening Office.
            from update_thinkcell_json import validate_request, canonical_sequence_contract
            from prepare_thinkcell_name import link_contract
            link_contract(chart)
            data_request = json.loads(data_path.read_text(encoding='utf-8-sig'))
            validate_request(data_request, chart['owner'].tag)
            canonical_sequence_contract(source, data_request,
                target={'slide_id':chart['doc']['slide_id'], **chart['frames'][0]})
    if not a.execute:
        return {'status':'CREATE_PREFLIGHT_PASS','writes':False,'method':'native_whole_slide_donor_copy',
                'data_update':'Inspect the new slide, then update its exact chart.',
                'style_scope':'Optional defaults for new elements; donor formatting is preserved.',
                'native_percentage_labels':percentage_before, 'data_request_validated':bool(data_path)}
    cmd=[powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(HERE/'extract_slide.ps1'),
         '-InputFile',str(source),'-SlideNumber',str(a.slide_number),'-OutputDirectory',str(destination)]
    proc=run_locked_subprocess(cmd,operation='native-extract',timeout_seconds=240,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,env=powershell_env(),creationflags=subprocess.CREATE_NO_WINDOW)
    need(proc.returncode==0, 'Native donor copy failed: '+proc.stdout[-1500:]+proc.stderr[-500:])
    result=json.loads((destination/'extraction.json').read_text(encoding='utf-8-sig'))
    if a.style_file:
        style_out=destination/'style'
        cmd=[powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(HERE/'load_style.ps1'),
             '-InputFile',result['output'],'-OutputDirectory',str(style_out),'-StyleFile',str(a.style_file.resolve())]
        proc=run_locked_subprocess(cmd,operation='native-style',timeout_seconds=180,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,env=powershell_env(),creationflags=subprocess.CREATE_NO_WINDOW)
        need(proc.returncode==0, 'Style defaults failed: '+proc.stdout[-1500:]+proc.stderr[-500:])
        result['style']=json.loads((style_out/'style.json').read_text(encoding='utf-8-sig'))
        result['output']=result['style']['output']
    need(hash_file(source)==a.expected_sha256.upper(), 'Source changed during creation.')
    if chart is not None:
        created = Path(result['output'])
        _, copied_charts, _ = inventory(created.read_bytes())
        copied_target = percent_target(copied_charts, chart)
        if percentage_before['status'] == 'native_percentage':
            result['native_percentage_labels'] = percent_contract().verify(percentage_before, created, copied_target)
        if data_path:
            updated = update(argparse.Namespace(input=created, output=destination/'populated.pptx',
                data_json=data_path, report=destination/'populated.report.json',
                expected_sha256=hash_file(created), slide_id=None, slide_number=1,
                shape_id=None, shape_tag=copied_target['frames'][0]['shape_tag'],
                named_only=False, ppttc=None, execute=True,
                require_native_percent=getattr(a, 'native_percent', False)))
            updated['creation_method'] = 'native_whole_slide_donor_copy_then_json_update'
            updated['donor_sha256'] = a.expected_sha256.upper()
            need(hash_file(source)==a.expected_sha256.upper(), 'Donor source changed during creation and data update.')
            write_json(destination/'creation.json', updated)
            return updated
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

@serialized_office
def convert_percent(a):
    """Convert supported absolute labels to genuine bare relative fields."""
    from percent_labels.prepare_native_relative_candidate import prepare
    from percent_labels.native_relative_contract import geometry
    from prepare_thinkcell_name import inventory, need
    source=a.input.resolve();dest=a.output_directory.resolve()
    need(not dest.exists(), 'Use a new output directory.')
    need(not dest.is_relative_to(HERE.parent), 'Output must be outside the installed skill.')
    need(hash_file(source)==a.expected_sha256.upper(), 'Source hash changed; inspect again.')
    all_labels=getattr(a,'all_labels',False)
    if all_labels:
        need(a.category_index is None and a.series_index is None,'Use either --all-labels or exact indices.')
    else:
        need(a.category_index is not None and a.series_index is not None,'Provide both zero-based indices or --all-labels.')
    original_geometry=geometry(source)
    dest.mkdir(parents=True)
    report={'status':'STARTED','source_sha256':a.expected_sha256.upper()}
    try:
        targets=[(cat,ser) for cat in range(3) for ser in range(3)] if all_labels else [(a.category_index,a.series_index)]
        current=source
        for index,(category,series) in enumerate(targets):
            candidate=dest/('candidate-'+str(index)+'.pptx')
            preparation=prepare(current,candidate,dest/('preparation-'+str(index)+'.json'),expected_sha=hash_file(current),
                category_index=category,series_index=series)
            current=candidate
        _, charts, _=inventory(candidate.read_bytes());need(len(charts)==1,'Expected one candidate chart.')
        request=baseline_request(candidate,charts[0]);request_path=dest/'baseline-request.json';write_json(request_path,request)
        result=update(argparse.Namespace(input=candidate,output=dest/'verified-candidate.pptx',report=dest/'update.report.json',
            expected_sha256=hash_file(candidate),data_json=request_path,slide_id=None,slide_number=1,shape_id=None,
            shape_tag=charts[0]['frames'][0]['shape_tag'],execute=True,named_only=False,ppttc=None))
        need(result['native_relative_labels']['status']=='native_relative_fields_preserved','Converted relative binding failed.')
        final_geometry=geometry(Path(result['output']))
        frame_changed=final_geometry['bounds_emu']!=original_geometry['bounds_emu']
        # With every label moved into genuine native text fields, think-cell
        # recomputes its chart-cache container. Native plot coordinates, type,
        # data and axis must still match exactly. Partial conversion keeps the
        # stricter original frame requirement.
        invariant=lambda value:{k:v for k,v in value.items() if k!='bounds_emu'}
        need(invariant(final_geometry)==invariant(original_geometry),
             'Conversion changed chart type, axis or native plot geometry.')
        need(not frame_changed or (all_labels and len(result['native_relative_labels']['after']['labels'])==9),
             'Partial conversion unexpectedly changed the chart-cache frame.')
        need(hash_file(source)==a.expected_sha256.upper(),'Source changed during conversion.')
        final=dest/'converted.pptx'
        with final.open('xb') as output:
            output.write(Path(result['output']).read_bytes())
        result.update(conversion='absolute_to_bare_native_relative_field',source_sha256=a.expected_sha256.upper(),
            converted_label_count=len(targets),
            output=str(final),output_sha256=hash_file(final),chart_type_axis_geometry_preserved=True,
            chart_cache_frame_recomputed=frame_changed,
            geometry_before=original_geometry,geometry_after=final_geometry,
            candidate_sha256=preparation.get('output_sha256'),source_unchanged=True)
        write_json(dest/'conversion.json',result)
        return result
    except Exception as error:
        report.update(status='FAILED',error=str(error));write_json(dest/'conversion-failure.json',report)
        raise

def main():
    p = argparse.ArgumentParser(description=__doc__)
    sub = p.add_subparsers(dest='command',required=True)
    sub.add_parser('doctor',help='Read-only runtime and dependency discovery.')
    cp=sub.add_parser('convert-percent',help='Convert supported absolute segment labels to bare native percentages on a copy.')
    cp.add_argument('--input',type=Path,required=True)
    cp.add_argument('--expected-sha256',required=True)
    cp.add_argument('--category-index',type=int)
    cp.add_argument('--series-index',type=int)
    cp.add_argument('--all-labels',action='store_true',help='Convert all nine labels of the tested 3x3 chart in one regeneration.')
    cp.add_argument('--output-directory',type=Path,required=True)
    c = sub.add_parser('create',help='Create a new native slide from a donor; preserves its chart types and features.')
    c.add_argument('--input',type=Path,required=True)
    c.add_argument('--expected-sha256',required=True)
    c.add_argument('--slide-number',type=int,required=True)
    c.add_argument('--output-directory',type=Path,required=True)
    c.add_argument('--style-file',type=Path)
    c.add_argument('--native-percent',action='store_true',help='Require and verify native percentage labels on the donor chart.')
    c.add_argument('--data-json',type=Path,help='Populate the single donor chart and verify it in the same operation.')
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
    u.add_argument('--require-native-percent',action='store_true',help='Reject charts without verified native percentage labels.')
    u.add_argument('--execute',action='store_true')
    selectors(u)
    v = sub.add_parser('verify',help='Check an exported or reopened candidate against its prepared source and intended data.')
    v.add_argument('--prepared',type=Path,required=True)
    v.add_argument('--input',type=Path,required=True)
    v.add_argument('--data-json',type=Path,required=True)
    v.add_argument('--automation-name',required=True)
    a = p.parse_args()
    try:
        result = {'doctor':lambda:doctor(),'create':lambda:create(a),'inspect':lambda:inspect(a),'update':lambda:update(a),'verify':lambda:verify(a),'convert-percent':lambda:convert_percent(a)}[a.command]()
        print(json.dumps(result,ensure_ascii=False,allow_nan=False))
    except (Exception, SystemExit) as e:
        print(json.dumps({'status':'REJECTED','error':str(e)}),file=sys.stderr)
        return 1
    return 0

if __name__ == '__main__':
    sys.exit(main())
