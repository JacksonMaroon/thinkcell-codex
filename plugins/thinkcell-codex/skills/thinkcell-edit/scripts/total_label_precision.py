"""Experimental coherent total formatting on one-category native sequence charts.

Install beside the chart/update helpers. Preparation never changes chart data;
official regeneration and native reopen are required before verify/delivery.
"""
from __future__ import annotations
import argparse
import copy
import io
import json
import math
from pathlib import Path
import subprocess
import sys
import tempfile
import zipfile
from lxml import etree as E

HERE = Path(__file__).resolve().parent
sys.path[:0] = [str(HERE), str(HERE / 'thinkcell_no_click' / 'implementation')]
from chart_geometry import _chart_identity, _select_chart, inventory, need, sha, streams, xml
from prepare_thinkcell_name import link_contract
from audit_thinkcell_integrity import logical_slides, NS
from runtime import powershell, powershell_env

PRECISION = 'm_prec17834/m_nDecimalDigits17909'


def valid_digits(value):
    need(isinstance(value, int) and not isinstance(value, bool) and 0 <= value <= 3,
         'Decimal digits must be an integer from 0 through 3')


def finite(value):
    need(isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value),
         'Expected total must be a finite number')
    return float(value)


def binding(chart):
    root, ids = chart['doc']['root'], chart['doc']['ids']
    need(chart['owner'].tag == 'CSequenceChartSE', 'Require a native sequence chart')
    columns = chart['table'].findall('ocol/elem')
    need(len(columns) == 1, 'Require exactly one category')
    vector = ids.get(columns[0].get('idref'))
    need(vector is not None and vector.tag == 'CSequenceChartDataVector', 'Unknown category binding')
    refs = vector.findall('m_cscdscgrp/elem')
    groups = root.findall('CSequenceChartDataScalarGroup')
    need(len(refs) == len(groups) == 1 and refs[0].get('idref') == groups[0].get('id'),
         'Require exactly one scalar group owned by the selected category')
    group = groups[0]
    ref = group.find('m_varsrcAbsoluteSum')
    source = ids.get(ref.get('idref')) if ref is not None else None
    need(source is not None and source.tag == 'CVariableSource', 'Missing absolute total source')
    value = source.find('m_varval')
    need(value is not None and value.get('type') == '1', 'Total source must be numeric')
    total = finite(float(value.get('val')))
    refs = source.findall('m_ctextvar/elem')
    need(len(refs) == 1 and refs[0].get('idref') in ids, 'Total needs one linked text variable')
    text = ids[refs[0].get('idref')]
    need(text.tag == 'CTextVariable', 'Unexpected total text binding')
    version = int(root.find('version').get('val'))
    active = lambda n: int(n.get('reqver', '0')) <= version < int(n.get('endver', '999999'))
    precision = [n for n in text.findall('m_prec17834') if active(n)]
    need(len(precision) == 1, 'Ambiguous active precision structure')
    digits = [n for n in precision[0].findall('m_nDecimalDigits17909') if active(n)]
    formats = [n for n in text.findall('m_bstrFormat') if active(n)]
    need(len(digits) == len(formats) == 1 and formats[0].text, 'Missing active total format')
    # This bounded adapter does not guess units, magnitudes, locale or date styles.
    p = precision[0]
    need(p.findtext('m_strPrefix', '') == '' and p.findtext('m_strSuffix17909', '') == '', 'Prefixed/suffixed totals unsupported')
    for key in ('m_nMagnitude17909', 'm_bNumberIsYear', 'm_esigndisplay'):
        node = p.find(key)
        need(node is not None and node.get('val') == '0', 'Scaled/date/special-sign totals unsupported')
    need(p.findtext('m_chDecimalSymbol17909') == '.', 'Non-dot decimal format unsupported')
    grouping = p.find('m_nGroupingDigits17909')
    need(grouping is not None and grouping.get('val') in {'3', '2147483647'}, 'Unknown grouping format')
    grouped = grouping.get('val') == '3'
    need(not grouped or p.findtext('m_chGroupingSymbol17909') == ',', 'Non-comma grouping unsupported')
    return {'group': group, 'source': source, 'text': text, 'digits': digits[0],
            'format': formats[0].text, 'total': total, 'grouped': grouped}


def inspect(path, selection, native_reopen=False):
    raw = Path(path).read_bytes()
    docs, charts, _ = inventory(raw)
    need(len(docs) == len(charts) == 1, 'Require exactly one native chart and carrier')
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        slides = logical_slides(z)
    need(len(slides) == 1, 'Require one slide')
    chart = _select_chart(charts, selection, allow_shape_id_drift=native_reopen)
    need(chart['exact'], 'Chart identity is not exact')
    link_contract(chart)
    return raw, chart, binding(chart), slides[0]['part']


def formatted_total(value, digits, grouped):
    return format(value, (',' if grouped else '') + '.' + str(digits) + 'f')


def bound_field(raw, slide_part, format_key):
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        root = xml(z.read(slide_part))
    fields = [n for n in root.findall('.//a:fld', NS) if n.get('type') == 'datetime' + format_key]
    need(len(fields) == 1 and len(fields[0].findall('a:t', NS)) == 1,
         'Selected total format must bind exactly one dynamic field with one text node')
    need(fields[0].get('id'), 'Selected field identity missing')
    return root, fields[0]


def make_plan(path, selection, digits=1):
    valid_digits(digits)
    raw, chart, b, slide_part = inspect(path, selection)
    slide, field = bound_field(raw, slide_part, b['format'])
    display = formatted_total(b['total'], digits, b['grouped'])
    key = ''.join("'" + char + "'" for char in display)
    need(not any(n.getparent().get('id') != b['text'].get('id') and n.text == key
                 for n in chart['doc']['root'].findall('CTextVariable/m_bstrFormat')), 'New total key collides with another text variable')
    need(not any(n is not field and n.get('type') == 'datetime' + key
                 for n in slide.findall('.//a:fld', NS)), 'New total key collides with another field')
    return {'schema_version': 2, 'source_sha256': sha(raw), 'target': _chart_identity(chart),
            'group_id': b['group'].get('id'), 'source_id': b['source'].get('id'),
            'text_id': b['text'].get('id'), 'old_digits': b['digits'].get('val'),
            'new_digits': digits, 'original_total': b['total'], 'existing_bstrFormat': b['format'],
            'slide_part': slide_part, 'field_id': field.get('id'), 'old_field_type': field.get('type'),
            'old_field_text': field.findtext('a:t', namespaces=NS), 'new_field_text': display,
            'new_bstrFormat': key,
            'edit': 'coherent_total_precision_format_and_bound_field', 'requires_official_regeneration': True,
            'requires_native_reopen_and_visual_review': True}


def assert_only(before, after, plan):
    original, candidate = xml(before), xml(after)
    text = candidate.find("./CTextVariable[@id='%s']" % plan['text_id'])
    need(text is not None, 'Selected text variable disappeared')
    matches = [n for n in text.findall(PRECISION) if n.get('val') == str(plan['new_digits'])]
    need(len(matches) == 1, 'Wrong or ambiguous prepared precision')
    matches[0].set('val', plan['old_digits'])
    fmt = text.find('m_bstrFormat')
    need(fmt is not None and fmt.text == plan['new_bstrFormat'], 'Wrong prepared total format key')
    fmt.text = plan['existing_bstrFormat']
    need(E.tostring(original, method='c14n') == E.tostring(candidate, method='c14n'),
         'Changed more than the selected active precision and format key')


def assert_slide_only(before, after, plan):
    original, candidate = xml(before), xml(after)
    hits = [n for n in candidate.findall('.//a:fld', NS) if n.get('id') == plan['field_id']]
    need(len(hits) == 1, 'Selected field identity changed')
    field = hits[0]
    need(field.get('type') == 'datetime' + plan['new_bstrFormat'] and
         field.findtext('a:t', namespaces=NS) == plan['new_field_text'], 'Wrong prepared total field')
    field.set('type', plan['old_field_type']); field.find('a:t', NS).text = plan['old_field_text']
    need(E.tostring(original, method='c14n') == E.tostring(candidate, method='c14n'),
         'Changed more than the selected dynamic field type/text')


def prepare(path, output, plan):
    path, output = Path(path), Path(output)
    need(not output.exists() and path.resolve() != output.resolve(), 'Distinct new output required')
    need(plan == make_plan(path, plan['target'], plan['new_digits']), 'Stale or modified precision plan')
    raw, chart, b, _ = inspect(path, plan['target'])
    doc = chart['doc']; before = doc['streams'][('think-cellXML',)]
    root = xml(before)
    # Use the exact active node index, including versioned sibling structures.
    original_node = b['digits']
    xpath = original_node.getroottree().getpath(original_node)
    matches = root.xpath(xpath)
    need(len(matches) == 1, 'Active precision path changed')
    matches[0].set('val', str(plan['new_digits']))
    fmt = root.find("./CTextVariable[@id='%s']/m_bstrFormat" % plan['text_id'])
    need(fmt is not None and fmt.text == plan['existing_bstrFormat'], 'Original format changed')
    fmt.text = plan['new_bstrFormat']
    after = E.tostring(root, encoding='utf-8'); assert_only(before, after, plan)
    with zipfile.ZipFile(io.BytesIO(raw)) as src: slide_before = src.read(plan['slide_part'])
    slide, field = bound_field(raw, plan['slide_part'], plan['existing_bstrFormat'])
    field.set('type', 'datetime' + plan['new_bstrFormat']); field.find('a:t', NS).text = plan['new_field_text']
    slide_after = E.tostring(slide, encoding='utf-8'); assert_slide_only(slide_before, slide_after, plan)
    with tempfile.TemporaryDirectory(prefix='tc_total_precision_', dir=output.parent) as td:
        td = Path(td); carrier = td / 'carrier.bin'; payload = td / 'model.xml'
        carrier.write_bytes(doc['ole']); payload.write_bytes(after)
        result = subprocess.run([powershell(), '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File',
            str(HERE / 'thinkcell_no_click/implementation/replace_ole_stream.ps1'),
            '-StoragePath', str(carrier), '-StreamBytesPath', str(payload)],
            capture_output=True, text=True, env=powershell_env(), timeout=60)
        need(result.returncode == 0, 'OLE precision preparation failed: ' + result.stderr[-500:])
        changed = carrier.read_bytes(); updated = streams(changed)
        need(set(updated) == set(doc['streams']), 'OLE stream inventory changed')
        need(all(updated[k] == v for k, v in doc['streams'].items() if k != ('think-cellXML',)), 'Other OLE stream changed')
        assert_only(before, updated[('think-cellXML',)], plan)
        need(sha(path.read_bytes()) == plan['source_sha256'], 'Source changed during preparation')
        with zipfile.ZipFile(io.BytesIO(raw)) as src, zipfile.ZipFile(output, 'x') as dst:
            dst.comment = src.comment
            for entry in src.infolist():
                payload = changed if entry.filename == doc['part'] else slide_after if entry.filename == plan['slide_part'] else src.read(entry.filename)
                dst.writestr(copy.copy(entry), payload)
        with zipfile.ZipFile(io.BytesIO(raw)) as src, zipfile.ZipFile(output) as dst:
            need(src.namelist() == dst.namelist() and dst.testzip() is None, 'ZIP inventory or CRC changed')
            need(all(src.read(n) == dst.read(n) for n in src.namelist() if n not in (doc['part'], plan['slide_part'])), 'Other package part changed')
            assert_slide_only(slide_before, dst.read(plan['slide_part']), plan)
    # Reversal proofs above preserve every numeric source/cache/model value. The
    # untouched datasource streams also preserve the original workbook exactly.
    need(sha(path.read_bytes()) == plan['source_sha256'], 'Source changed during preparation')
    return {'status': 'TOTAL_PRECISION_PREPARED_REGENERATION_REQUIRED', 'source_unchanged': True,
            'output_sha256': sha(output.read_bytes()), 'only_total_format_representations_changed': True,
            'numeric_sources_and_caches_unchanged': True, 'embedded_datasource_unchanged': True}


def verify(path, plan, expected_total):
    valid_digits(plan['new_digits']); expected = finite(expected_total)
    raw, chart, b, slide_part = inspect(path, plan['target'], native_reopen=True)
    need((b['group'].get('id'), b['source'].get('id'), b['text'].get('id')) ==
         (plan['group_id'], plan['source_id'], plan['text_id']), 'Absolute total binding changed')
    need(b['digits'].get('val') == str(plan['new_digits']), 'Requested active precision did not survive')
    numeric_tolerance = 8 * max(math.ulp(b['total']), math.ulp(expected))
    need(abs(b['total'] - expected) <= numeric_tolerance, 'Selected total model value differs from expected')
    field_type = 'datetime' + b['format']
    _, selected = bound_field(raw, slide_part, b['format'])
    text = selected.findtext('a:t', namespaces=NS)
    wanted = formatted_total(expected, plan['new_digits'], b['grouped'])
    need(text == wanted, 'Selected dynamic total field has wrong text or precision')
    return {'status': 'TOTAL_PRECISION_BOUND_FIELD_VERIFIED', 'output_sha256': sha(raw),
            'target': _chart_identity(chart), 'source_id': plan['source_id'], 'text_id': plan['text_id'],
            'model_total': b['total'], 'text': text, 'field_id': selected.get('id'),
            'native_shape_id_changed': _chart_identity(chart)['shape_id'] != plan['target']['shape_id'],
            'field_type': field_type, 'requires_native_reopen_and_visual_review': True}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    m = sub.add_parser('make-plan'); m.add_argument('--input', required=True); m.add_argument('--selection-json', required=True)
    m.add_argument('--digits', type=int, default=1); m.add_argument('--output', required=True)
    for name in ('prepare', 'verify'):
        p = sub.add_parser(name); p.add_argument('--input', required=True); p.add_argument('--plan', required=True)
        p.add_argument('--output', required=True) if name == 'prepare' else p.add_argument('--expected-total', type=float, required=True)
    a = parser.parse_args()
    read = lambda p: json.loads(Path(p).read_text(encoding='utf-8-sig'))
    if a.command == 'make-plan':
        result = make_plan(a.input, read(a.selection_json), a.digits)
        with Path(a.output).open('x', encoding='utf-8') as f: json.dump(result, f, indent=2, allow_nan=False)
    elif a.command == 'prepare': result = prepare(a.input, a.output, read(a.plan))
    else: result = verify(a.input, read(a.plan), a.expected_total)
    print(json.dumps(result, allow_nan=False))


if __name__ == '__main__': main()
