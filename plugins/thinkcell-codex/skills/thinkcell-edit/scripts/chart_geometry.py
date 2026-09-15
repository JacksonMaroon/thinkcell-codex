"""Experimental exclusive-anchor geometry on one-chart native donor copies.

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


def tagged_bounds(path):
    """Read the entire visible native bundle, including detached legends."""
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


def inspect_geometry(path):
    original = Path(path).read_bytes()
    docs, charts, names = inventory(original)
    need(len(docs) == len(charts) == 1, 'Geometry requires one chart and one native carrier')
    with zipfile.ZipFile(io.BytesIO(original)) as z:
        need(len(logical_slides(z)) == 1, 'Geometry requires a single-slide donor')
    c, d = charts[0], docs[0]
    need(c['exact'] and c['owner'].tag in FAMILIES, 'Unsupported exact chart identity')
    link_contract(c)
    need(c['owner_name'] and sum(n['name'].casefold() == c['owner_name'].casefold() for n in names) == 1,
         'Use automatic naming and native generation before geometry on an unnamed donor')
    root, ids = d['root'], d['ids']
    version = int(root.find('version').get('val'))
    need(version >= 36196, 'Regenerate an older model before geometry')
    need(not c['owner'].findall('m_vecgidGroup/elem'), 'Grouped chart unsupported')
    legend_types = {'CSequenceChartLegendSE','CScatterChartLegendSE'}
    allowed = FAMILIES | {'CContainerSE'} | legend_types
    need(all(n.tag in allowed for n in root if n.tag.endswith('SE')), 'Other native elements on donor slide')
    for n in root.iter():
        if 'constraint' in n.tag.lower():
            need(not any(x.get('idref') not in (None, '0') for x in n.iter()), 'Shared layout constraints unsupported')

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
    legends = [anchors(n, False) for n in root if n.tag in legend_types]
    need(len(legends) <= 1, 'Multiple legends unsupported')
    return {'source_sha256': sha(original), 'name': c['owner_name'], 'family': c['owner'].tag,
            'plot': main, 'outer':outer, 'legends': legends, 'bundle': tagged_bounds(path),
            'carrier': d['part'], 'chart': c, 'document': d}


def make_plan(path, frame):
    g = inspect_geometry(path)
    need(set(frame) == {'left','top','width','height'}, 'Require a complete outer frame')
    need(all(not isinstance(v, bool) and isinstance(v, (int,float)) and math.isfinite(v) for v in frame.values()), 'Invalid frame')
    x, y, w, h = (frame[k] for k in ('left','top','width','height'))
    need(x >= 0 and y >= 0 and w > 0 and h > 0, 'Invalid frame extent')
    l, t, r, b = (a['value'] for a in g['plot'])
    bl, bt, br, bb = g['bundle']
    # Keep label/legend margins in points. Native text is not scaled down.
    nl, nt = x+max(0,l-bl), y+max(0,t-bt)
    nr, nb = x+w-max(0,br-r), y+h-max(0,bb-b)
    need(nr-nl >= 36 and nb-nt >= 36, 'Frame is too small for donor labels and legend')
    if g['family'] == 'CPieChartSE':
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
    for legend in g['legends']:
        ll,lt,lr,lb = (a['value'] for a in legend)
        lw,lh = lr-ll,lb-lt
        need(lw <= w and lh <= h, 'Existing legend does not fit; choose a more compact donor')
        if lt >= b:
            dl,dt = nl+(ll-l), nb+(lt-b)
        elif ll >= r:
            dl,dt = nr+(ll-r), nt+(lt-t)
        else:
            dl = nl+(ll-l)*(nr-nl)/(r-l)
            dt = nt+(lt-t)*(nb-nt)/(b-t)
        dl = min(max(dl,x),x+w-lw); dt = min(max(dt,y),y+h-lh)
        edits.extend(dict(a,new_value=v*8) for a,v in zip(legend,(dl,dt,dl+lw,dt+lh)))
    need(len({e['grid_id'] for e in edits}) == len(edits), 'Shared chart and legend coordinates')
    return {'source_sha256':g['source_sha256'], 'automation_name':g['name'], 'outer_frame':frame,
            'expected_plot':[nl,nt,nr,nb], 'edits':edits, 'experimental':True}


def prepare(path, output, plan):
    path, output = Path(path), Path(output)
    g = inspect_geometry(path); d = g['document']
    need(not output.exists() and path.resolve() != output.resolve(), 'Distinct new output required')
    need(plan == make_plan(path, plan['outer_frame']), 'Stale or modified geometry plan')
    before = d['streams'][('think-cellXML',)]
    root = xml(before)
    for e in plan['edits']:
        node = root.find("./CGridline[@id='%s']/m_gveps/m_gvValue" % e['grid_id'])
        need(node is not None and node.get('val') == e['old_string'], 'Coordinate identity changed')
        node.set('val', format(e['new_value'], '.20E'))
    after = E.tostring(root, encoding='utf-8')

    def prove(payload):
        original, candidate = xml(before), xml(payload)
        for e in plan['edits']:
            node = candidate.find("./CGridline[@id='%s']/m_gveps/m_gvValue" % e['grid_id'])
            need(float(node.get('val')) == e['new_value'], 'Wrong new coordinate')
            node.set('val',e['old_string'])
        need(E.tostring(original,method='c14n') == E.tostring(candidate,method='c14n'), 'Non-coordinate model changes')
    prove(after)
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
        prove(ss[('think-cellXML',)])
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
    g = inspect_geometry(path)
    actual = [a['value'] for a in g['plot']]
    need(g['name'] == plan['automation_name'], 'Chart identity changed')
    need(len(actual)==4 and all(math.isfinite(v) for v in actual), 'Invalid native plot bounds')
    need(actual[2]-actual[0]>=36 and actual[3]-actual[1]>=36, 'Native plot is too small')
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
