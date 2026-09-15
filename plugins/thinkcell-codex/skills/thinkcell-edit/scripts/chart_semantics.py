"""Read-only, version-bounded chart semantics. Never patches chart values."""
from collections import Counter
from audit_thinkcell_integrity import _variable_value


def kind(c):
    node = c['owner'].find('m_ect')
    code = node.get('val') if node is not None else None
    return {'4':'waterfall','11':'mekko-percent','12':'mekko-units'}.get(code, c['owner'].tag)


def details(c):
    ids = c['doc']['ids']
    vectors = [ids[x.get('idref')] for x in c['table'].findall('ocol/elem')]
    result = {'kind':kind(c)}
    if kind(c).startswith('mekko'):
        result['column_widths'] = [_variable_value(ids,v,'m_varsrcAbsoluteExtent') for v in vectors]
    if kind(c) == 'waterfall':
        anchors, equals, grounds = {}, [], []
        for ci,v in enumerate(vectors):
            scalars = [x.get('idref') for x in v.findall('ocol/elem')]
            eq = v.find('m_scdscEquals')
            if eq is not None and eq.get('idref') != '0':
                equals.append([scalars.index(eq.get('idref')),ci])
            for gi,gr in enumerate(v.findall('m_cscdscgrp/elem')):
                g = ids[gr.get('idref')]
                ground = g.find('m_eground')
                grounds.append([ci,gi,ground.get('val') if ground is not None else None])
                for child in g:
                    if child.tag.startswith('m_ascanchor'):
                        for ai,a in enumerate(child):
                            anchors[a.get('idref')] = [ci,gi,child.tag,ai]
                    elif child.tag.startswith('m_scanchor'):
                        anchors[child.get('idref')] = [ci,gi,child.tag,0]
        connections = []
        for conn in c['doc']['root'].findall('CWaterfallConnector'):
            ends = [anchors.get(conn.find(tag).get('idref')) for tag in ('m_anchorSource','m_anchorSink')]
            if any(x is None for x in ends):
                raise ValueError('Waterfall connector topology is outside the supported group-anchor contract.')
            connections.append(ends)
        result.update(equals_slots=equals,connectors=sorted(connections),grounds=grounds)
    return result


def audit_scope(path, audit):
    """Specialized charts use model/datasheet + semantic + native-render checks.

    The generic Office cache parity algorithm does not describe waterfall
    offset series or Mekko shapes. Do not report those checks as passing.
    Only a single specialized chart is allowed, so no sibling is exempted.
    """
    from prepare_thinkcell_name import inventory, need
    _,cs,_ = inventory(path.read_bytes())
    special = len(cs)==1 and kind(cs[0]) in {'waterfall','mekko-percent','mekko-units'}
    exclusions = {'no_model_visible_chart_mismatch','strict_parity_available','strict_parity_pass'} if special else set()
    failures = [k for k,v in audit['assertions'].items() if not v and k not in exclusions]
    need(not failures, 'Integrity checks failed: '+', '.join(failures))
    return {'native_cache_parity': 'not_implemented_for_specialized_chart' if special else 'pass',
            'specialized_visual_semantics_review_required':special,
            'excluded_inapplicable_checks':sorted(exclusions)}


def feature_summary(c):
    counts=Counter(x.tag for x in c['doc']['root'])
    return {k:v for k,v in counts.items() if any(s in k.lower() for s in ('connector','difference','cagr','trend','errorbar','legend','axisbreak'))}
