"""Read-only retention gate for bare native relative fields on an absolute axis."""
import math
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP

from .discover_percent_semantics import discover
from .native_percent_contract import _identity, _target_rows
from chart_geometry import inventory, streams, xml, active


def geometry(path):
    """Conversion-only baseline signature; data updates may legitimately rescale."""
    _, charts, _ = inventory(Path(path).read_bytes())
    if len(charts) != 1 or len(charts[0]['frames']) != 1:
        raise ValueError('Conversion requires exactly one native chart frame')
    chart = charts[0]
    root = xml(streams(chart['doc']['ole'])[('think-cellXML',)])
    ids = {n.get('id'):n for n in root if n.get('id')}
    owner = chart['owner']
    axis = ids[owner.find('m_daxisPrimaryValue').get('idref')]
    version=int(root.find('version').get('val'))
    plot={}
    for side in ('left','top','right','bottom'):
        bindings=[b for b in owner.findall('m_rectgrdlnanchorPlot/'+side+'/m_grdlnBinding') if active(b,version)]
        if len(bindings)!=1:
            raise ValueError('Conversion plot binding is ambiguous')
        grid=ids[bindings[0].get('idref')]
        plot[side]=(grid.get('id'),grid.find('m_gveps/m_gvValue').get('val'),grid.find('m_eorient').get('val'))
    def field(node,name):
        value=node.find(name)
        return None if value is None else (value.text,dict(value.attrib))
    return {'owner_type':owner.tag,'chart_type':field(owner,'m_ect'), 'plot':plot,
        'owner_id':owner.get('id'),'table':field(owner,'m_dtable'),
        'native_chart':field(owner,'m_pptseqchart'),'plot_area':field(owner,'m_pptpolylinePlotArea'),
        'orientation':field(owner,'m_eorient'),'gap_width':field(owner,'m_nGapWidth'),
        'series_type':field(owner,'m_estDefault'),'bounds_emu':chart['frames'][0]['bounds_emu'],
        'axis':{name:field(axis,name) for name in ('m_fMinValue','m_fMaxValue','m_fUserScaleUnit',
                  'm_edaxistype','m_euseraxistype','m_eaxisextentMin','m_eaxisextentMax')},
        'axis_precision': __import__('lxml').etree.tostring(axis.find('m_precUser')).decode()}


def snapshot(path, target):
    path = Path(path)
    if _identity(target).get('owner_type') not in (None, 'CSequenceChartSE'):
        return {'status':'not_applicable', 'labels':[]}
    rows, identity = _target_rows(discover(path), target, path)
    selected = [r for r in rows if r['relative_text_variable'] is not None
                and r['absolute_text_variable'] is None and r['relative_suffix'] == '%'
                and len(r['physical_shapes']) == 1
                and len(r['physical_shapes'][0]['fields']) == 1
                and r['physical_shapes'][0]['literal_texts'] == [r['physical_shapes'][0]['fields'][0]['text']]]
    if not selected:
        return {'status':'not_applicable', 'labels':[]}
    _, charts, _ = inventory(path.read_bytes())
    matches = [c for c in charts if c['doc']['part'] == selected[0]['chart_part']
               and any(f.get('shape_tag') == identity.get('shape_tag') for f in c['frames'])]
    if len(matches) != 1:
        raise ValueError('Bare relative label owner is ambiguous')
    root = xml(streams(matches[0]['doc']['ole'])[('think-cellXML',)])
    ids = {n.get('id'):n for n in root if n.get('id')}
    result = []
    for row in selected:
        source = ids[row['relative_source_id']]
        text = ids[row['relative_text_variable']]
        digits = int(row['relative_decimal_digits'])
        numerator, denominator = row['numerator'], row['denominator']
        if not 0 <= digits <= 3 or denominator is None or denominator <= 0:
            raise ValueError('Unsupported bare relative denominator/precision')
        ratio = float(numerator) / denominator
        val = source.find('m_varval')
        # Derived scalar-relative sources omit a persisted m_varval on this
        # profile. Their live binding plus regenerated field value is checked.
        if val is not None and not math.isclose(float(val.get('val')),ratio,rel_tol=1e-9,abs_tol=1e-9):
            raise ValueError('Native relative source is stale')
        expected = str((Decimal(str(numerator))*100/Decimal(str(denominator))).quantize(
            Decimal('1').scaleb(-digits),rounding=ROUND_HALF_UP))+'%'
        physical = row['physical_shapes'][0]
        field = physical['fields'][0]
        if field['text'] != expected or field['type'] != 'datetime'+text.findtext('m_bstrFormat'):
            raise ValueError('Bare relative field does not display its current native ratio')
        if text.findtext('m_prec17834/m_strPrefix') not in (None,''):
            raise ValueError('Bare relative field has an unsupported prefix')
        result.append({'category_index':row['category_index'],'series_index':row['series_index'],
            'shape_tag':row['shape_tag'],'source_guid':source.find('m_guid').get('val'),
            'digits':digits,'numerator':numerator,'denominator':denominator,'display':expected})
    return {'status':'native_relative_fields','chart_tag':identity.get('shape_tag'),'labels':result}


def verify(before, path, target):
    if before['status'] != 'native_relative_fields':
        return before
    after = snapshot(path, target)
    keys = ('category_index','series_index','shape_tag','source_guid','digits')
    expected = {tuple(r[k] for k in keys) for r in before['labels']}
    actual = {tuple(r[k] for k in keys) for r in after['labels']}
    if (after['status'] != 'native_relative_fields' or before['chart_tag'] != after.get('chart_tag')
            or expected != actual or len(after['labels']) != len(before['labels'])):
        raise ValueError('Bare native relative labels lost identity, binding or precision')
    return {'status':'native_relative_fields_preserved','before':before,'after':after}
