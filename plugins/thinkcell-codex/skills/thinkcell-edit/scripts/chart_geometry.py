"""Experimental exclusive-anchor geometry on selected native charts.

Only coordinate values are patched. Official regeneration and native readback
are required before this intermediate file can be delivered.
"""
from pathlib import Path
import copy
import io
import math
import subprocess
import sys
import tempfile
import zipfile
from lxml import etree as E

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE/'thinkcell_no_click/implementation'))
from prepare_thinkcell_name import inventory, link_contract, need, sha, streams, xml
from audit_thinkcell_integrity import NS, logical_slides, relationship_map
from runtime import powershell, powershell_env

SIDES = ('left', 'top', 'right', 'bottom')
FAMILIES = {'CPieChartSE', 'CSequenceChartSE', 'CScatterChartSE'}


def active(node, version):
    return int(node.get('reqver', '0')) <= version < int(node.get('endver', '999999'))


def _chart_identity(chart):
    """Stable selectors carried from planning through preparation and verify."""
    frame = chart['frames'][0]
    return {'carrier': chart['doc']['part'], 'owner_id': chart['owner'].get('id'),
            'shape_id': frame['shape_id'], 'shape_tag': frame['shape_tag'],
            'automation_name': chart['owner_name']}


def _select_chart(charts, selection=None, allow_shape_id_drift=False):
    """Resolve exactly one native chart; multi-chart inputs require a selector."""
    if selection is None:
        need(len(charts) == 1, 'Multiple native charts require an exact selected chart identity')
        return charts[0]
    need(isinstance(selection, dict), 'Invalid selected chart identity')
    keys = ('carrier', 'owner_id', 'shape_id', 'shape_tag', 'automation_name')
    if allow_shape_id_drift:
        keys = tuple(key for key in keys if key != 'shape_id')
    supplied = {k: selection[k] for k in keys if selection.get(k) not in (None, '')}
    need(supplied, 'Selected chart identity is missing')
    matches = []
    for chart in charts:
        if not chart['exact'] or not chart['frames']:
            continue
        actual = _chart_identity(chart)
        if all(actual[k] == v for k, v in supplied.items()):
            matches.append(chart)
    need(len(matches) == 1, 'Selected chart is missing, ambiguous, or changed')
    return matches[0]


def tagged_bounds(path, shape_tags=None):
    """Read visible bounds for the selected native bundle only.

    A detached legend cannot be attributed safely when sibling charts exist, so
    the multi-chart route rejects it before this function is called.
    """
    bounds = []
    with zipfile.ZipFile(path) as z:
        slide = logical_slides(z)[0]['part']
        root, rels = xml(z.read(slide)), relationship_map(z, slide)
        for shape in root.find('p:cSld/p:spTree', NS):
            cn = shape.find('.//p:cNvPr', NS)
            if cn is not None and cn.get('hidden') in {'1','true'}:
                continue
            tagged = False
            for ref in shape.findall('.//p:tags', NS):
                tags = xml(z.read(rels[ref.get('{'+NS['r']+'}id')]['resolved']))
                tagged |= any(n.get('name', '').upper() == 'THINKCELLSHAPEDONOTDELETE' for n in tags)
            if not tagged:
                continue
            if shape_tags is not None:
                values = []
                for ref in shape.findall('.//p:tags', NS):
                    tags = xml(z.read(rels[ref.get('{'+NS['r']+'}id')]['resolved']))
                    values.extend(n.get('val') for n in tags if n.get('name', '').upper() == 'THINKCELLSHAPEDONOTDELETE')
                if not any(value in shape_tags for value in values):
                    continue
            grouped = shape.tag == '{'+NS['p']+'}grpSp'
            if grouped:
                need(shape.find('.//c:chart',NS) is None,'Grouped native chart frame unsupported')
                for leaf in shape.iter():
                    if leaf.tag in {'{'+NS['p']+'}sp','{'+NS['p']+'}pic','{'+NS['p']+'}cxnSp'}:
                        refs=leaf.findall('.//p:tags',NS)
                        need(any(any(n.get('name','').upper()=='THINKCELLSHAPEDONOTDELETE' for n in
                             xml(z.read(rels[ref.get('{'+NS['r']+'}id')]['resolved']))) for ref in refs),
                             'Mixed ordinary/native legend group unsupported')
            xf = shape.find('p:grpSpPr/a:xfrm', NS) if grouped else shape.find('p:xfrm', NS)
            if xf is None: xf = shape.find('p:spPr/a:xfrm', NS)
            need(xf is not None and int(xf.get('rot', '0')) == 0, 'Unknown native shape transform')
            off, ext = xf.find('a:off', NS), xf.find('a:ext', NS)
            x, y = int(off.get('x'))/12700, int(off.get('y'))/12700
            w, h = int(ext.get('cx'))/12700, int(ext.get('cy'))/12700
            bounds.append((x, y, x+w, y+h))
    need(bounds, 'No native bundle bounds')
    return [min(b[0] for b in bounds), min(b[1] for b in bounds),
            max(b[2] for b in bounds), max(b[3] for b in bounds)]


def inspect_geometry(path, selection=None, allow_shape_id_drift=False):
    original = Path(path).read_bytes()
    docs, charts, names = inventory(original)
    c = _select_chart(charts, selection, allow_shape_id_drift)
    with zipfile.ZipFile(io.BytesIO(original)) as z:
        need(len(logical_slides(z)) == 1, 'Geometry requires a single-slide donor')
    d = c['doc']
    need(c['exact'] and c['owner'].tag in FAMILIES, 'Unsupported exact chart identity')
    link_contract(c)
    named = c['owner_name'] and sum(n['name'].casefold() == c['owner_name'].casefold() for n in names) == 1
    # A generated name is preferred for downstream JSON regeneration.  The
    # geometry-only route can still safely inspect an unnamed chart when its
    # caller supplied every durable native selector.
    explicit_identity = isinstance(selection, dict) and all(selection.get(k) not in (None, '')
        for k in ('carrier', 'owner_id', 'shape_id', 'shape_tag'))
    need(named or explicit_identity,
         'Use automatic naming or provide carrier, owner, shape ID, and shape tag for geometry')
    root, ids = d['root'], d['ids']
    version = int(root.find('version').get('val'))
    need(version >= 36196, 'Regenerate an older model before geometry')
    need(not c['owner'].findall('m_vecgidGroup/elem'), 'Grouped chart unsupported')
    legend_types = {'CSequenceChartLegendSE','CScatterChartLegendSE'}

    def anchors(owner, plot):
        result = []
        owned = {n.get('idref') for n in owner.findall('m_agrdlnanchor/elem')}
        for index, side in enumerate(SIDES):
            if plot:
                bindings = owner.findall('m_rectgrdlnanchorPlot/'+side+'/m_grdlnBinding')
                bindings = [b for b in bindings if active(b, version)]
            else:
                refs = owner.findall('m_agrdlnanchor/elem')
                need(len(refs) == 4, 'Unknown legend anchor order')
                anchor = ids[refs[index].get('idref')]
                bindings = [b for b in anchor.findall('m_grdlnBinding') if active(b, version)]
            need(len(bindings) == 1, 'Ambiguous coordinate binding')
            binding = bindings[0]; grid = ids[binding.get('idref')]
            val = grid.find('m_gveps/m_gvValue')
            need(grid.tag == 'CGridline' and val is not None, 'Unknown coordinate grid')
            orient = int(grid.find('m_eorient').get('val'))
            need(orient == index % 2, 'Unexpected coordinate orientation')
            for ref in root.iter():
                if ref is binding or ref.get('idref') != grid.get('id'):
                    continue
                parent = ref.getparent()
                if parent.get('id') in owned and not active(ref, version):
                    continue
                need(False, 'Coordinate is shared outside selected element')
            result.append({'grid_id': grid.get('id'), 'side': side,
                           'old_string': val.get('val'), 'value': float(val.get('val'))/8})
        need(len({a['grid_id'] for a in result}) == 4, 'Repeated coordinate grid')
        return result

    main = anchors(c['owner'], True)
    orientation=c['owner'].find('m_eorient')
    outer=anchors(c['owner'],False) if c['owner'].tag=='CSequenceChartSE' and orientation is not None and orientation.get('val')=='1' else []
    selected_legends = []
    legend_tags = set()
    for legend in (n for n in root if n.tag in legend_types):
        owner = legend.find('m_cse')
        need(owner is not None and owner.get('idref') in ids,
             'Detached legend has no resolvable chart owner')
        if owner.get('idref') != c['owner'].get('id'):
            continue
        selected_legends.append(legend)
        rect_ref = legend.find('m_pptrect')
        need(rect_ref is not None and rect_ref.get('idref') in ids,
             'Selected detached legend has no native frame')
        tag = ids[rect_ref.get('idref')].findtext('m_bstrShapeName')
        # Some current models render the legend inside the chart's tagged frame
        # and leave the rectangle without a shape name. In that form the chart
        # tag already covers its visible bundle; only add a separate tag when
        # the model supplies one.
        if tag:
            legend_tags.add(tag)
    legends = [anchors(n, False) for n in selected_legends]
    need(len(legends) <= 1, 'Multiple legends unsupported')
    selected_ids = ({c['owner'].get('id'), c['table'].get('id')} |
                    {n.get('id') for n in selected_legends} |
                    {a['grid_id'] for a in main + outer + [a for group in legends for a in group]})
    for node in root.iter():
        if 'constraint' not in node.tag.lower():
            continue
        refs = {desc.get('idref') for desc in node.iter() if desc.get('idref') not in (None, '0')}
        need(not refs.intersection(selected_ids), 'Selected chart participates in a layout constraint')
    visible_bundle = tagged_bounds(path, {f['shape_tag'] for f in c['frames']} | legend_tags)
    # Legacy detached legends can have no separately tagged drawing shape. The
    # model anchors remain authoritative for their extent, so include them in
    # the planning frame before calculating plot margins or allowing a clamp.
    logical_legend_bounds = [[a['value'] for a in legend] for legend in legends]
    bundle_bounds = [visible_bundle] + logical_legend_bounds
    bundle = [min(bounds[0] for bounds in bundle_bounds), min(bounds[1] for bounds in bundle_bounds),
              max(bounds[2] for bounds in bundle_bounds), max(bounds[3] for bounds in bundle_bounds)]
    return {'source_sha256': sha(original), 'name': c['owner_name'], 'family': c['owner'].tag,
            'plot': main, 'outer':outer, 'legends': legends,
            'bundle': bundle,
            'carrier': d['part'], 'chart': c, 'document': d,
            'target': _chart_identity(c), 'sibling_chart_count': len(charts)-1}


def make_plan(path, frame, selection=None, translation=None, plot_bounds=None, legend_bounds=None):
    g = inspect_geometry(path, selection)
    need(set(frame) == {'left','top','width','height'}, 'Require a complete outer frame')
    need(all(not isinstance(v, bool) and isinstance(v, (int,float)) and math.isfinite(v) for v in frame.values()), 'Invalid frame')
    x, y, w, h = (frame[k] for k in ('left','top','width','height'))
    need(x >= 0 and y >= 0 and w > 0 and h > 0, 'Invalid frame extent')
    l, t, r, b = (a['value'] for a in g['plot'])
    bl, bt, br, bb = g['bundle']
    if plot_bounds is not None:
        need(translation is None and isinstance(plot_bounds, (list, tuple)) and len(plot_bounds) == 4 and
             all(not isinstance(value, bool) and isinstance(value, (int, float)) and math.isfinite(value)
                 for value in plot_bounds), 'Exact plot bounds require four finite values')
        nl, nt, nr, nb = plot_bounds
    elif translation is None:
        # Keep label/legend margins in points. Native text is not scaled down.
        nl, nt = x+max(0,l-bl), y+max(0,t-bt)
        nr, nb = x+w-max(0,br-r), y+h-max(0,bb-b)
    else:
        need(isinstance(translation, dict) and set(translation) == {'x', 'y'},
             'Translation requires x and y')
        need(all(not isinstance(v, bool) and isinstance(v, (int, float)) and math.isfinite(v)
                 for v in translation.values()), 'Invalid translation')
        nl, nt, nr, nb = l+translation['x'], t+translation['y'], r+translation['x'], b+translation['y']
    need(nr-nl >= 36 and nb-nt >= 36, 'Frame is too small for donor labels and legend')
    if g['family'] == 'CPieChartSE' and translation is None and plot_bounds is None:
        # Outside pie labels reflow as diameter changes. Keep extra label room.
        nl += 8; nt += 8; nr -= 8; nb -= 8
        side = min(nr-nl, nb-nt)
        need(side >= 36, 'Pie frame is too small for its labels')
        nl += ((nr-nl)-side)/2; nt += ((nb-nt)-side)/2
        nr, nb = nl+side, nt+side
    edits = [dict(a, new_value=v*8) for a,v in zip(g['plot'], (nl,nt,nr,nb))]
    if g.get('outer'):
        ol,ot,orr,ob=(a['value'] for a in g['outer'])
        values=(nl-(l-ol),nt-(t-ot),nr+(orr-r),nb+(ob-b))
        edits.extend(dict(a,new_value=v*8) for a,v in zip(g['outer'],values))
    expected_legends = []
    need(legend_bounds is None or (isinstance(legend_bounds,list) and len(legend_bounds)==len(g['legends'])), 'Explicit legend bounds must cover every selected legend')
    for legend_index,legend in enumerate(g['legends']):
        ll,lt,lr,lb = (a['value'] for a in legend)
        lw,lh = lr-ll,lb-lt
        need(lw <= w and lh <= h, 'Existing legend does not fit; choose a more compact donor')
        if legend_bounds is not None:
            explicit=legend_bounds[legend_index]
            need(isinstance(explicit,list) and len(explicit)==4 and all(not isinstance(v,bool) and isinstance(v,(int,float)) and math.isfinite(v) for v in explicit),'Invalid legend coordinates')
            dl,dt,dr,db=explicit
            need(abs((dr-dl)-lw)<0.01 and abs((db-dt)-lh)<0.01,'Move existing legends without resizing them')
        elif translation is not None:
            dl, dt = ll+translation['x'], lt+translation['y']
        elif lt >= b:
            dl,dt = nl+(ll-l), nb+(lt-b)
        elif ll >= r:
            dl,dt = nr+(ll-r), nt+(lt-t)
        else:
            dl = nl+(ll-l)*(nr-nl)/(r-l)
            dt = nt+(lt-t)*(nb-nt)/(b-t)
        if translation is None and legend_bounds is None:
            dl = min(max(dl,x),x+w-lw); dt = min(max(dt,y),y+h-lh)
        need(dl >= x and dt >= y and dl+lw <= x+w and dt+lh <= y+h,
             'Legend exceeds the planned outer frame')
        expected_legends.append([dl,dt,dl+lw,dt+lh])
        edits.extend(dict(a,new_value=v*8) for a,v in zip(legend,(dl,dt,dl+lw,dt+lh)))
    need(len({e['grid_id'] for e in edits}) == len(edits), 'Shared chart and legend coordinates')
    plan = {'source_sha256':g['source_sha256'], 'automation_name':g['name'], 'outer_frame':frame,
            'expected_plot':[nl,nt,nr,nb], 'expected_legends':expected_legends,
            'translation':translation, 'plot_bounds':list(plot_bounds) if plot_bounds is not None else None, 'legend_bounds':legend_bounds,
            'edits':edits, 'experimental':True}
    # Retain a mock-friendly legacy make_plan surface; real inspection always
    # supplies these fields and therefore produces the stricter selected plan.
    if 'target' in g:
        plan['target'] = g['target']
        plan['sibling_chart_count'] = g['sibling_chart_count']
    return plan


def assert_only_selected_coordinates_changed(before, payload, edits):
    """Prove that resetting selected grid values restores the whole model.

    This is deliberately whole-document comparison: siblings may be complex,
    but none of their model XML is permitted to change in this route.
    """
    original, candidate = xml(before), xml(payload)
    for edit in edits:
        node = candidate.find("./CGridline[@id='%s']/m_gveps/m_gvValue" % edit['grid_id'])
        need(node is not None and float(node.get('val')) == edit['new_value'], 'Wrong new coordinate')
        node.set('val', edit['old_string'])
    need(E.tostring(original, method='c14n') == E.tostring(candidate, method='c14n'),
         'Non-coordinate model changes')


def prepare(path, output, plan):
    path, output = Path(path), Path(output)
    need(isinstance(plan.get('target'), dict), 'Geometry plan is missing selected chart identity')
    g = inspect_geometry(path, plan['target']); d = g['document']
    need(not output.exists() and path.resolve() != output.resolve(), 'Distinct new output required')
    need(plan == make_plan(path, plan['outer_frame'], plan['target'], plan.get('translation'), plan.get('plot_bounds'), plan.get('legend_bounds')),
         'Stale or modified geometry plan')
    before = d['streams'][('think-cellXML',)]
    root = xml(before)
    for e in plan['edits']:
        node = root.find("./CGridline[@id='%s']/m_gveps/m_gvValue" % e['grid_id'])
        need(node is not None and node.get('val') == e['old_string'], 'Coordinate identity changed')
        node.set('val', format(e['new_value'], '.20E'))
    after = E.tostring(root, encoding='utf-8')

    assert_only_selected_coordinates_changed(before, after, plan['edits'])
    with tempfile.TemporaryDirectory(prefix='tc_geometry_', dir=output.parent) as td:
        td=Path(td); carrier=td/'carrier.bin'; payload=td/'model.xml'
        carrier.write_bytes(d['ole']); payload.write_bytes(after)
        p=subprocess.run([powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',
            str(HERE/'thinkcell_no_click/implementation/replace_ole_stream.ps1'),
            '-StoragePath',str(carrier),'-StreamBytesPath',str(payload)],
            capture_output=True,text=True,env=powershell_env(),timeout=60)
        need(p.returncode == 0, 'Coordinate preparation failed: '+p.stderr[-500:])
        changed=carrier.read_bytes(); ss=streams(changed)
        need(set(ss)==set(d['streams']), 'OLE stream inventory changed')
        need(all(v==ss[k] for k,v in d['streams'].items() if k!=('think-cellXML',)), 'Other OLE stream changed')
        assert_only_selected_coordinates_changed(before, ss[('think-cellXML',)], plan['edits'])
        original=path.read_bytes(); need(sha(original)==plan['source_sha256'], 'Source changed')
        with zipfile.ZipFile(io.BytesIO(original)) as src, zipfile.ZipFile(output,'x') as dst:
            dst.comment=src.comment
            for item in src.infolist():
                dst.writestr(copy.copy(item),changed if item.filename==d['part'] else src.read(item.filename))
        with zipfile.ZipFile(io.BytesIO(original)) as src,zipfile.ZipFile(output) as dst:
            need(src.namelist()==dst.namelist() and dst.testzip() is None, 'ZIP inventory or CRC changed')
            need(all(src.read(n)==dst.read(n) for n in src.namelist() if n!=d['part']), 'Other package parts changed')
    return {'status':'GEOMETRY_PREPARED_REGENERATION_REQUIRED','source_sha256':g['source_sha256'],
            'output_sha256':sha(output.read_bytes()),'only_exclusive_coordinates_changed':True}


def verify(path, plan, tolerance=1.01):
    need(not isinstance(tolerance,bool) and isinstance(tolerance,(int,float)) and
         math.isfinite(tolerance) and 0 <= tolerance <= 2, 'Invalid geometry tolerance')
    need(isinstance(plan.get('expected_plot'),list) and len(plan['expected_plot'])==4 and
         all(not isinstance(v,bool) and isinstance(v,(int,float)) and math.isfinite(v) for v in plan['expected_plot']),
         'Invalid expected plot bounds')
    if plan.get('target') is None:
        g = inspect_geometry(path)
    else:
        need(isinstance(plan['target'], dict), 'Geometry plan has an invalid selected chart identity')
        g = inspect_geometry(path, plan['target'], allow_shape_id_drift=True)
    actual = [a['value'] for a in g['plot']]
    if plan.get('target') is not None:
        stable_keys = ('carrier', 'owner_id', 'shape_tag', 'automation_name')
        need(all(g['target'][key] == plan['target'][key] for key in stable_keys) and
             g['name'] == plan['automation_name'], 'Chart identity changed')
    else:
        need(g['name'] == plan['automation_name'], 'Chart identity changed')
    need(len(actual)==4 and all(math.isfinite(v) for v in actual), 'Invalid native plot bounds')
    need(actual[2]-actual[0]>=36 and actual[3]-actual[1]>=36, 'Native plot is too small')
    if 'expected_legends' in plan:
        expected_legends = plan['expected_legends']
        need(isinstance(expected_legends, list) and len(expected_legends) == len(g['legends']),
             'Native legend identity changed')
        for actual_legend, expected_legend in zip(g['legends'], expected_legends):
            actual_values = [anchor['value'] for anchor in actual_legend]
            need(all(abs(a-b) <= tolerance for a, b in zip(actual_values, expected_legend)),
                 'Native legend did not retain requested geometry')
    # The frame specifies available outer space, not an exact plot size. A native
    # horizontal sequence chart can modestly contract its right/bottom plot edge.
    horizontal=g['family']=='CSequenceChartSE' and len(g.get('outer',[]))==4
    deltas=[a-b for a,b in zip(actual,plan['expected_plot'])]
    need(all(-(4 if horizontal and i>=2 else tolerance)<=delta<=tolerance
             for i,delta in enumerate(deltas)), 'Native plot did not retain requested geometry')
    l,t,r,b = g['bundle']; f = plan['outer_frame']
    need(l>=f['left']-tolerance and t>=f['top']-tolerance and
         r<=f['left']+f['width']+tolerance and b<=f['top']+f['height']+tolerance,
         'Native labels or legend exceed the planned outer frame; revise geometry or donor')
    return {'native_plot_bounds':actual,'native_bundle_bounds':g['bundle'], 'fits_outer_frame':True,
            'native_plot_adjustment_points':dict(zip(SIDES,deltas)),
            'native_inward_reflow_accepted':horizontal and any(d<-tolerance for d in deltas[2:]),
            'inward_reflow_limit_points':4 if horizontal else 0}
