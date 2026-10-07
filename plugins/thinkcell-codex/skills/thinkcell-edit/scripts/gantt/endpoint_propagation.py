"""Invoked native finish/start endpoint equality on the verified empty-sheet profile.

Preparation is not a delivery result. Regenerate the emitted job officially,
then save/reopen with native_verify_scoped.ps1 and run verify. Review its PNG.
No autonomous scheduling, scalar binding, lag or duration preservation.
"""
from pathlib import Path
from datetime import datetime
import sys, io, zipfile, json, re, hashlib, tempfile, subprocess, os, argparse
import xlrd, pyxlsb
from xml.etree import ElementTree as ET
import olefile

from dependency_adapter import inspect_native_anchors, anchor_edges, finish_start_model
from portable_gantt_adapter import GanttPackage, powershell, powershell_env, REPLACE_STREAM


def _date(value):
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}T00:00:00", value): raise ValueError("midnight date without timezone required")
    return datetime.fromisoformat(value)


def plan_endpoint_dates(bars, edges, requests):
    """Propagate only exact FS endpoint equality; keep every other endpoint fixed."""
    by_id = {str(b['bar_id']): b for b in bars}
    if len(by_id) != len(bars): raise ValueError('duplicate bar identity')
    result = {bid: list(b['dates']) for bid, b in by_id.items()}
    assignments = {}
    for bid, side, value in requests:
        bid = str(bid)
        if bid not in by_id or side not in ('Low', 'High'): raise ValueError('unknown endpoint')
        _date(value)
        key = (bid, side)
        if key in assignments and assignments[key] != value: raise ValueError('conflicting endpoint requests')
        assignments[key] = value
    queue = list(assignments)
    while queue:
        key = queue.pop(0)
        bid, side = key; value = assignments[key]
        result[bid][0 if side == 'Low' else 1] = value
        for parent, parent_side, child, child_side, anchor in edges:
            if (parent, parent_side) != key: continue
            if parent_side != 'High' or child_side != 'Low': raise ValueError('affected graph must be FS only')
            if child not in by_id: raise ValueError('affected endpoint is not a supported bar')
            child_key = (child, child_side)
            if child_key in assignments:
                if assignments[child_key] != value: raise ValueError('conflicting dependent endpoint request')
                continue
            assignments[child_key] = value; queue.append(child_key)
    for bid, dates in result.items():
        if _date(dates[0]) >= _date(dates[1]): raise ValueError('propagation would eliminate or invert a bar')
    for parent, parent_side, child, child_side, anchor in edges:
        if (child,child_side) not in assignments: continue
        if parent not in result or child not in result: raise ValueError('affected endpoint is not a supported bar')
        if result[parent][0 if parent_side=='Low' else 1] != result[child][0 if child_side=='Low' else 1]:
            raise ValueError('request would break an existing anchor')
    return {bid: dates for bid, dates in result.items() if dates != by_id[bid]['dates']}


def prepare(source, output, predecessor_guid, end_date, *, expected_sha256, report, job, successor_guid=None):
    source, output = Path(source).resolve(), Path(output).resolve()
    paths = guard_paths([source, output, report, job], fresh=[output, report, job])
    report, job = paths[2:]
    if not re.fullmatch(r'[0-9a-fA-F]{64}', expected_sha256): raise ValueError('expected SHA-256 required')
    if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', end_date): raise ValueError('end date must be YYYY-MM-DD')
    datetime.fromisoformat(end_date)
    if hashlib.sha256(source.read_bytes()).hexdigest() != expected_sha256.lower(): raise ValueError('source hash differs')
    before = hashlib.sha256(source.read_bytes()).hexdigest()
    state = inspect_native_anchors(source)
    if not state['consistent']: raise ValueError('source native owner must be consistent')
    if state['schema_version'] != '38790': raise ValueError('verified normalized schema 38790 required')
    if any(not b['visible'] for b in state['bars']): raise ValueError('all bars must already be visible; donor preparation is a separate operation')
    if len(state['bars']) != 4: raise ValueError('verified four-visible-bar profile required')
    _bar_index(state)
    if len({b['line_id'] for b in state['bars']}) != 4 or len({b['shape_tag'] for b in state['bars']}) != 4:
        raise ValueError('bar line/physical shape ownership must be exclusive')
    authenticated_empty_sheet(source)
    def bar_guid(guid, side):
        matches = [b for b in state['bars'] if b['date_guids'][side] == guid]
        if len(matches) != 1: raise ValueError('source GUID must resolve uniquely')
        return matches[0]
    parent = bar_guid(predecessor_guid, 1)
    p = GanttPackage.__new__(GanttPackage); p.path = source
    with zipfile.ZipFile(source) as z:
        p.entries = z.infolist(); p.blobs = {e.filename:z.read(e.filename) for e in p.entries}
    if len([n for n in p.blobs if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)]) != 1: raise ValueError('one-slide package required')
    found = []
    for name, blob in p.blobs.items():
        if not name.startswith('ppt/embeddings/') or not name.endswith('.bin'): continue
        with olefile.OleFileIO(io.BytesIO(blob)) as o:
            if not o.exists('think-cellXML'): continue
            model = o.openstream('think-cellXML').read().decode('utf8')
            if '<CGanttSE' in model: found.append((name, model))
    if len(found) != 1: raise ValueError('unique Gantt model required')
    p.ole_part, model = found[0]
    modelroot = ET.fromstring(model)
    if next(modelroot.iter('CGanttScalar'),None) is not None or next(modelroot.iter('CGanttScalarAnchor'),None) is not None:
        raise ValueError('scalar or scalar-anchor profile is outside this route')
    if successor_guid:
        child = bar_guid(successor_guid, 0)
        model, aid = finish_start_model(model, parent['bar_id'], child['bar_id'])
    edges = anchor_edges(model)
    if not any(e[0] == parent["bar_id"] and e[1:2] == ("High",) for e in edges):
        raise ValueError("existing finish/start anchor or explicit successor required")
    outgoing = [e for e in edges if e[0] == parent["bar_id"] and e[1] == "High"]
    if len(outgoing) != 1 or outgoing[0][3] != "Low": raise ValueError("unique reciprocal finish/start attachment required")
    changed = plan_endpoint_dates(state['bars'], edges, [(parent['bar_id'], 'High', end_date + 'T00:00:00')])
    root = ET.fromstring(model); ids = {n.get('id'):n for n in root.iter() if n.get('id')}
    owner = root.find('CGanttSE'); timerange = owner.find('m_timerange')
    if timerange.find('m_bFiveDayWeek').get('val') != '0': raise ValueError('only seven-day calendar geometry supported')
    interval = timerange.find('m_intvltp'); begin = datetime.fromisoformat(interval.find('begin').get('val')); end = datetime.fromisoformat(interval.find('end').get('val'))
    left = int(ids[owner.find('m_pptlineChartLeft').get('idref')].find('m_rectPPTShape').get('left'))
    right = int(ids[owner.find('m_pptlineChartRight').get('idref')].find('m_rectPPTShape').get('left'))
    if right <= left or end <= begin: raise ValueError('invalid native timeline geometry')
    for node in root:
        if node.tag not in {'CGanttBar','CGanttShade','CGanttBracket','CGanttProcess','CGanttMilestone'}: continue
        item_dates = [datetime.fromisoformat(d.get('val')) for d in node.iter('m_datetime')]
        if any(not begin <= d <= end for d in item_dates): raise ValueError('all native items must lie inside existing normalized timeline')
        placed = list(node.iter('m_bPlaced'))
        if item_dates and placed and all(d.get('val') == '0' for d in placed):
            raise ValueError('hidden native item requires separate donor preparation')
    def x_for(value):
        value = datetime.fromisoformat(value)
        if not begin <= value <= end: raise ValueError('date outside existing calendar')
        return round(left + (right-left)*(value-begin).total_seconds()/(end-begin).total_seconds())
    by_id = {b['bar_id']:b for b in state['bars']}
    # Authenticate calendar mapping and the native 1/8-point cache coordinate scale.
    for bid in by_id:
        b = by_id[bid]
        if not b['visible']: raise ValueError('affected bar must be visible')
        if any(abs(x_for(d)-b['model_rect'][i*2])>1 for i,d in enumerate(b['dates'])): raise ValueError('native date-to-geometry mapping differs')
        tr = b['physical_transform']; rect = b['model_rect']
        expected = [rect[0]*1587.5,rect[1]*1587.5,(rect[2]-rect[0])*1587.5,(rect[3]-rect[1])*1587.5]
        if any(abs(a-b)>1 for a,b in zip(tr,expected)): raise ValueError('physical native coordinate profile differs')
    slide = p.blobs['ppt/slides/slide1.xml'].decode()
    for bid, dates in changed.items():
        b = by_id[bid]
        match = re.search(r'<CGanttBar id="'+re.escape(bid)+r'".*?</CGanttBar>',model,re.S); block = match.group()
        for side, value in zip(('Low','High'),dates):
            pattern = r'(<m_varsrc'+side+r'>.*?<m_datetime val=")[^"]+("/>)'
            block,count = re.subn(pattern,lambda m:m.group(1)+value+m.group(2),block,count=1,flags=re.S)
            if count != 1: raise ValueError('date source closure mismatch')
        model = model[:match.start()]+block+model[match.end():]
        new_left,new_right = map(x_for,dates)
        match = re.search(r'<CPPTAutoShapeLine id="'+re.escape(b['line_id'])+r'".*?</CPPTAutoShapeLine>',model,re.S)
        block,count = re.subn(r'<m_rectPPTShape left="\d+" top="(\d+)" right="\d+" bottom="(\d+)"',lambda m:f'<m_rectPPTShape left="{new_left}" top="{m.group(1)}" right="{new_right}" bottom="{m.group(2)}"',match.group(),count=1)
        if count != 1: raise ValueError('line closure mismatch')
        model = model[:match.start()]+block+model[match.end():]
        shape,tr = p._shape(b['shape_tag']); old = re.search(r'<a:off x="\d+" y="(\d+)"/><a:ext cx="\d+" cy="(\d+)"',shape)
        sl,sr = round(new_left*1587.5),round(new_right*1587.5)
        shape2 = shape[:old.start()]+f'<a:off x="{sl}" y="{old.group(1)}"/><a:ext cx="{sr-sl}" cy="{old.group(2)}"'+shape[old.end():]
        slide = slide.replace(shape,shape2,1)
    p.blobs['ppt/slides/slide1.xml'] = slide.encode()
    # Name only the owner and owning table for the official empty-datasheet control.
    for tag in ('CGanttSE','CGanttTable'):
        match = re.search(r'<'+tag+r' id="[^"]+".*?</'+tag+r'>',model,re.S)
        block,count = re.subn(r'<m_strName\s*/>|<m_strName>.*?</m_strName>','<m_strName>TC_GANTT_ENDPOINT_CONTROL</m_strName>',match.group(),count=1,flags=re.S)
        if count != 1: raise ValueError('native owner/table naming profile missing')
        model = model[:match.start()]+block+model[match.end():]
    if anchor_edges(model) != edges: raise ValueError('graph changed during date propagation')
    with tempfile.TemporaryDirectory(dir=output.parent) as td:
        td=Path(td); carrier=td/'carrier.bin';xml=td/'model.xml';candidate=td/'candidate.pptx'
        carrier.write_bytes(p.blobs[p.ole_part]);xml.write_text(model,encoding='utf8')
        run=subprocess.run([powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(REPLACE_STREAM),'-StoragePath',str(carrier),'-StreamBytesPath',str(xml)],capture_output=True,text=True,env=powershell_env(),timeout=60)
        if run.returncode: raise RuntimeError(run.stderr or run.stdout)
        p.blobs[p.ole_part]=carrier.read_bytes()
        with zipfile.ZipFile(candidate,'x') as z:
            for entry in p.entries:z.writestr(entry,p.blobs[entry.filename])
        actual=inspect_native_anchors(candidate)
        if actual['edges'] != edges: raise ValueError('saved graph differs')
        for b in actual['bars']:
            if b['dates'] != changed.get(b['bar_id'],by_id[b['bar_id']]['dates']): raise ValueError('saved date mismatch')
        if native_semantics(source)[1] != native_semantics(candidate)[1]: raise ValueError('unrelated embedded streams changed')
        authenticated_empty_sheet(candidate)
        if grade(candidate,candidate,predecessor_guid)['status'] != 'ADAPTER_ENDPOINT_NATIVE_PASS': raise ValueError('prepared endpoint geometry validation failed')
        if hashlib.sha256(source.read_bytes()).hexdigest() != before: raise ValueError('source changed')
        payload=candidate.read_bytes()
    result = {'status':'ADAPTER_MANAGED_ENDPOINT_PROPAGATION_PREPARED','changed':changed,'edges':edges,'source':str(source),'output':str(output),'job':str(job),'source_sha256':before,'prepared_sha256':hashlib.sha256(payload).hexdigest(),'predecessor_end_guid':predecessor_guid,'successor_start_guid':successor_guid,'end_date':end_date,'source_unchanged':True,'native_verification_required':True,'visual_review_required':True,'engine_autonomous_scheduling':False,'duration_preservation':False,'lag_scheduling':False}
    publish_files([(output,payload),
        (job,(json.dumps([{'template':str(output),'data':[{'name':'TC_GANTT_ENDPOINT_CONTROL','table':[[]]}]}],indent=2)+'\n').encode('utf8')),
        (report,(json.dumps(result,indent=2)+'\n').encode('utf8'))], source=source, expected_sha256=before)
    return result




def publish_files(files, *, source=None, expected_sha256=None):
    """Create all public artifacts exclusively, cleaning only this call's creations."""
    created=[]
    try:
        if source is not None and hashlib.sha256(Path(source).read_bytes()).hexdigest()!=expected_sha256:
            raise ValueError('source changed before publication')
        for path, payload in files:
            with Path(path).open('xb') as stream:
                created.append(Path(path))
                stream.write(payload)
        if source is not None and hashlib.sha256(Path(source).read_bytes()).hexdigest()!=expected_sha256:
            raise ValueError('source changed during publication')
    except Exception:
        for path in reversed(created): path.unlink()
        raise

def guard_paths(paths, *, fresh):
    resolved = [Path(p).resolve() for p in paths]
    keys = [os.path.normcase(str(p)) for p in resolved]
    if len(set(keys)) != len(keys): raise ValueError('all input, output, job and report paths must be distinct')
    for p in map(lambda x: Path(x).resolve(), fresh):
        if p.exists(): raise ValueError('fresh output, job and report paths required')
        if not p.parent.is_dir(): raise ValueError('output parent must exist')
    return resolved


def _bar_index(state):
    bars = state['bars']
    result = {tuple(b['date_guids']): b for b in bars}
    if len(result) != len(bars): raise ValueError('duplicate date GUID endpoint pair')
    if len({g for b in bars for g in b['date_guids']}) != 2 * len(bars):
        raise ValueError('duplicate date source GUID')
    return result


def _model_and_streams(path):
    found = []; siblings = {}
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.startswith('ppt/embeddings/'): continue
            raw_part = archive.read(name)
            if not name.endswith('.bin'):
                siblings['package:' + name] = raw_part; continue
            with olefile.OleFileIO(io.BytesIO(raw_part)) as carrier:
                if not carrier.exists('think-cellXML'):
                    siblings['package:' + name] = raw_part; continue
                raw = carrier.openstream('think-cellXML').read()
                if b'<CGanttSE' not in raw:
                    siblings['package:' + name] = raw_part; continue
                found.append((raw.decode('utf8'), {'/'.join(s): carrier.openstream(s).read()
                    for s in carrier.listdir() if '/'.join(s) != 'think-cellXML'}))
    if len(found) != 1: raise ValueError('one native Gantt model required')
    model, streams = found[0]; streams.update(siblings)
    return model, streams


def authenticated_empty_sheet(path):
    """Inspect the actual embedded BIFF/XLSB worksheet; never infer an empty sheet."""
    model, streams = _model_and_streams(path)
    root = ET.fromstring(model); table = root.find('CGanttTable')
    if table is None or table.find('m_advisesink') is None or table.find('m_advisesink').get('idref') != '0':
        raise ValueError('linked datasheet is outside this route')
    if next(root.iter('CGanttScalar'),None) is not None or next(root.iter('CGanttScalarAnchor'),None) is not None:
        raise ValueError('datasheet scalar binding is outside this route')
    roots = [name for name in streams if name.endswith('/Workbook') or name.endswith('/Package')]
    if len(roots) != 1: raise ValueError('unique internal worksheet required')
    name = roots[0]; raw = streams[name]
    if name.endswith('/Workbook'):
        workbook = xlrd.open_workbook(file_contents=raw)
        sheets = [(sheet.name, sheet.nrows, sheet.ncols) for sheet in workbook.sheets()]
    else:
        with tempfile.TemporaryDirectory() as directory:
            workbook_path = Path(directory) / 'datasheet.xlsb'; workbook_path.write_bytes(raw)
            sheets = []
            with pyxlsb.open_workbook(workbook_path) as workbook:
                for sheet_name in workbook.sheets:
                    with workbook.get_sheet(sheet_name) as sheet:
                        rows = list(sheet.rows())
                        sheets.append((sheet_name, len(rows), max((len(row) for row in rows), default=0)))
    if sheets != [('Sheet1', 0, 0)]: raise ValueError('authenticated empty unbound Sheet1 required')
    return sheets


def native_semantics(path):
    """Read full graph and all native date source identities, plus exact sibling streams."""
    model, streams = _model_and_streams(path)
    root = ET.fromstring(model)
    nodes = [n for n in root.iter() if n.get('id')]
    ids = {n.get('id'): n for n in nodes}
    if len(ids) != len(nodes): raise ValueError('duplicate native object identity')
    def identity(bid):
        n = ids[bid]
        return (n.tag, tuple(g.get('val') for side in ('Low', 'High')
                            for g in n.findall('m_varsrc' + side + '/m_guid')))
    graph = sorted((identity(p), ps, identity(c), cs) for p, ps, c, cs, _ in anchor_edges(model))
    dates = {}
    for node in list(root):
        values = tuple(d.get('val') for d in node.iter('m_datetime'))
        guids = tuple(g.get('val') for g in node.iter('m_guid'))
        if not values or not guids: continue
        key = (node.tag, guids)
        if key in dates: raise ValueError('duplicate native date source identity')
        dates[key] = values
    return dates, {name: hashlib.sha256(raw).hexdigest() for name, raw in streams.items()}, graph


def grade(prepared, native, predecessor_guid):
    expected = inspect_native_anchors(prepared); actual = inspect_native_anchors(native)
    e, a = _bar_index(expected), _bar_index(actual)
    matches = [b for b in actual['bars'] if b['date_guids'][1] == predecessor_guid]
    if len(matches) != 1: raise ValueError('predecessor GUID must resolve uniquely')
    parent = matches[0]
    fs = [edge for edge in actual['edges'] if edge[0] == parent['bar_id'] and edge[1] == 'High']
    by_id = {b['bar_id']: b for b in actual['bars']}
    if len(fs) != 1 or any(edge[3] != 'Low' or edge[2] not in by_id for edge in fs):
        raise ValueError('selected finish/start graph must attach supported bar starts')
    ed, es, eg = native_semantics(prepared); ad, ast, ag = native_semantics(native)
    def close(x, y, tolerance):
        if x is None or y is None: return x == y
        return len(x) == len(y) and all(abs(i - j) <= tolerance for i, j in zip(x, y))
    same = set(e) == set(a)
    checks = {'native_consistent': actual['consistent'], 'all_bar_guid_identities_preserved': same,
        'all_exact_native_dates_match_request': same and all(e[g]['dates'] == a[g]['dates'] for g in e),
        'all_bar_visibility_preserved': same and all(e[g]['visible'] == a[g]['visible'] for g in e),
        'all_bar_native_rects_preserved': same and all(close(e[g]['model_rect'], a[g]['model_rect'], 1) for g in e),
        'all_bar_physical_transforms_preserved': same and all(close(e[g]['physical_transform'], a[g]['physical_transform'], 2) for g in e),
        'all_native_date_sources_preserved': ed == ad,
        'all_anchor_semantics_preserved': eg == ag, 'unrelated_datasheet_stream_hashes_equal': es == ast,
        'selected_endpoints_visible': parent['visible'] and all(by_id[c]['visible'] for _, _, c, _, _ in fs)}
    checks['exact_endpoint_date_equality'] = all(parent['dates'][1] == by_id[c]['dates'][0] for _, _, c, _, _ in fs)
    checks['model_endpoint_geometry_equal'] = checks['selected_endpoints_visible'] and all(parent['model_rect'][2] == by_id[c]['model_rect'][0] for _, _, c, _, _ in fs)
    checks['physical_endpoint_geometry_equal'] = checks['selected_endpoints_visible'] and all(abs(parent['physical_transform'][0] + parent['physical_transform'][2] - by_id[c]['physical_transform'][0]) <= 1 for _, _, c, _, _ in fs)
    return {'status': 'ADAPTER_ENDPOINT_NATIVE_PASS' if all(checks.values()) else 'ADAPTER_ENDPOINT_NATIVE_FAIL',
        'checks': checks, 'parent_dates': parent['dates'], 'child_dates': [by_id[c]['dates'] for _, _, c, _, _ in fs],
        'expected_bar_count': len(e), 'native_bar_count': len(a),
        'prepared_sha256': hashlib.sha256(Path(prepared).read_bytes()).hexdigest(),
        'native_sha256': hashlib.sha256(Path(native).read_bytes()).hexdigest(),
        'geometry_tolerance': {'native_model_units': 1, 'physical_emu': 2},
        'engine_autonomous_scheduling': False, 'duration_preservation': False, 'lag_scheduling': False,
        'visual_review_required': True}


def assert_candidate_matches(rebuilt, prepared):
    expected_model, expected_streams = _model_and_streams(rebuilt)
    actual_model, actual_streams = _model_and_streams(prepared)
    if expected_model != actual_model or expected_streams != actual_streams:
        raise ValueError("prepared candidate differs from source reconstruction")
    with zipfile.ZipFile(rebuilt) as expected, zipfile.ZipFile(prepared) as actual:
        if expected.namelist() != actual.namelist(): raise ValueError("prepared package parts differ")
        for name in expected.namelist():
            if name.startswith("ppt/embeddings/") and name.endswith(".bin"): continue
            if expected.read(name) != actual.read(name): raise ValueError("prepared unrelated package part differs")

def verify(preparation_report, native_report, report):
    """Validate source/job/official/native lineage and strict saved model readback."""
    preparation_report, native_report, report = guard_paths([preparation_report, native_report, report], fresh=[report])
    preparation = json.loads(preparation_report.read_text(encoding='utf-8-sig'))
    native_record = json.loads(native_report.read_text(encoding='utf-8-sig'))
    source = Path(preparation['source']).resolve(); prepared = Path(preparation['output']).resolve()
    job = Path(preparation['job']).resolve(); generated = Path(native_record['source']).resolve()
    native = Path(native_record['output']).resolve(); render = Path(native_record['render']).resolve()
    guard_paths([source, prepared, job, generated, native, render, preparation_report, native_report, report], fresh=[report])
    for file in [source, prepared, job, generated, native, render]:
        if not file.is_file(): raise ValueError('required native lineage file missing')
    if hashlib.sha256(source.read_bytes()).hexdigest() != preparation['source_sha256'].lower():
        raise ValueError('original source changed')
    if hashlib.sha256(prepared.read_bytes()).hexdigest() != preparation['prepared_sha256'].lower():
        raise ValueError('prepared package changed')
    if hashlib.sha256(generated.read_bytes()).hexdigest() != native_record['source_sha256'].lower():
        raise ValueError('official generated package changed')
    wanted_job = [{'template': str(prepared), 'data': [{'name': 'TC_GANTT_ENDPOINT_CONTROL', 'table': [[]]}]}]
    if json.loads(job.read_text(encoding='utf8')) != wanted_job: raise ValueError('official job differs from preparation')
    if preparation['status'] != 'ADAPTER_MANAGED_ENDPOINT_PROPAGATION_PREPARED': raise ValueError('invalid preparation status')
    for gate in ['native_reopen_pass', 'source_unchanged', 'other_presentations_unchanged', 'cleanup_release_complete']:
        if native_record.get(gate) is not True: raise ValueError('native source/user/reopen/cleanup gate failed')
    # Rebuild the candidate from the preserved source and selectors. Reports
    # cannot waive a source/candidate mismatch by changing their stored hashes.
    with tempfile.TemporaryDirectory(dir=prepared.parent) as directory:
        folder = Path(directory); rebuilt = folder / 'reconstructed.pptx'
        prepare(source, rebuilt, preparation['predecessor_end_guid'], preparation['end_date'],
            expected_sha256=preparation['source_sha256'], report=folder / 'plan.json', job=folder / 'job.ppttc',
            successor_guid=preparation.get('successor_start_guid'))
        assert_candidate_matches(rebuilt, prepared)
    authenticated_empty_sheet(prepared); authenticated_empty_sheet(native)
    result = grade(prepared, native, preparation['predecessor_end_guid'])
    generated_grade = grade(prepared, generated, preparation['predecessor_end_guid'])
    result['checks']['official_regeneration_semantics'] = generated_grade['status'] == 'ADAPTER_ENDPOINT_NATIVE_PASS'
    result['checks']['original_source_unchanged'] = True
    result['checks']['native_reopen_source_user_cleanup_gates'] = True
    result['render'] = str(render)
    result['status'] = 'ADAPTER_ENDPOINT_NATIVE_PASS_VISUAL_REVIEW_REQUIRED' if all(result['checks'].values()) else 'ADAPTER_ENDPOINT_NATIVE_FAIL'
    publish_files([(report,(json.dumps(result, indent=2)+'\n').encode('utf8'))])
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    p = commands.add_parser('prepare')
    for flag in ['input', 'expected-sha256', 'output', 'report', 'job', 'predecessor-end-guid', 'end-date']:
        p.add_argument('--' + flag, required=True)
    p.add_argument('--successor-start-guid')
    v = commands.add_parser('verify')
    for flag in ['preparation-report', 'native-report', 'report']: v.add_argument('--' + flag, required=True)
    args = parser.parse_args()
    if args.command == 'prepare':
        result = prepare(args.input, args.output, args.predecessor_end_guid, args.end_date,
            expected_sha256=args.expected_sha256, report=args.report, job=args.job,
            successor_guid=args.successor_start_guid)
    else: result = verify(args.preparation_report, args.native_report, args.report)
    print(json.dumps(result, indent=2))
    return 2 if result['status'] == 'ADAPTER_ENDPOINT_NATIVE_FAIL' else 0


if __name__ == '__main__': sys.exit(main())
