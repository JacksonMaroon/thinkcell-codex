"""Prepare a bounded new persistent Excel link from an authenticated native adviser.

Offline preparation only. Native save/reopen and changed-data evidence are required
before claiming persistence. No client assets are bundled by this adapter.
"""
from __future__ import annotations
import argparse
import base64
import copy
import io
import json
import math
from pathlib import Path
import sys
import uuid
import zipfile

import olefile
from lxml import etree as E
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from audit_thinkcell_integrity import sequence_tables, logical_slides, relationship_map, NS, R
import portable_rebind as rebind
from native_moniker import serialize_composite, composite_file, item_name


def one(root, path):
    nodes = root.findall(path)
    if len(nodes) != 1:
        raise RuntimeError(f'expected one {path}, found {len(nodes)}')
    return nodes[0]


def profile(xml, chart_name, linked):
    root = E.fromstring(xml)
    if one(root, 'version').get('val') != '38732':
        raise RuntimeError('only model version 38732 is supported')
    chart, table, grid = [one(root, tag) for tag in ('CSequenceChartSE', 'CSequenceChartDataTable', 'CSmartGrid')]
    ids = [n.get('id') for n in root.iter() if n.get('id') is not None]
    if len(set(ids)) != len(ids) or any(not v.isdigit() or v == '0' for v in ids):
        raise RuntimeError('duplicate or invalid model IDs')
    if one(chart, 'm_ect').get('val') != '0' or one(chart, 'm_dtable').get('idref') != table.get('id'):
        raise RuntimeError('unsupported chart type or table ownership')
    actual_name = (table.findtext('m_strName') or '').strip()
    if actual_name != chart_name:
        raise RuntimeError('chart automation name mismatch')
    if one(table, 'm_bstrRangeName').get('val') != 'think-cellChild0' or one(table, 'm_bExternalStorage').get('val') != '1':
        raise RuntimeError('unsupported child storage identity')
    data = sequence_tables(root)[0]
    if len(data['series_names']) != 4 or len(data['categories']) != 2 or any(len(v) != 2 for v in data['series_values']):
        raise RuntimeError('only four series by two categories are supported')
    names = data['series_names']
    if len(set(names)) != 4 or any(not isinstance(n, str) or not n for n in names):
        raise RuntimeError('series names must be unique nonempty strings')
    if any(v is None or not math.isfinite(v) for row in data['series_values'] for v in row):
        raise RuntimeError('model has unsupported numeric values')
    sinks, registry, ref = root.findall('CAdviseSink'), one(grid, 'm_cadvisesink'), one(table, 'm_advisesink')
    if linked:
        if len(sinks) != 1 or len(registry) != 1 or ref.get('idref') != sinks[0].get('id') or registry[0].get('idref') != sinks[0].get('id'):
            raise RuntimeError('donor adviser table/grid registration does not resolve')
        rebind._fields(xml)
        sink = sinks[0]
        # The bounded adviser is self-contained, never import references to a donor chart.
        if any(n.get('idref') not in (None, '0', sink.get('id')) for n in sink.iter()):
            raise RuntimeError('donor adviser has references outside its closure')
    else:
        if sinks or len(registry) or ref.get('idref') != '0':
            raise RuntimeError('target is already linked or adviser registry is nonempty')
        sink = None
    return root, table, grid, sink, data


def read_carrier(raw):
    with olefile.OleFileIO(io.BytesIO(raw)) as ole:
        streams = {tuple(p) for p in ole.listdir()}
        if not {('think-cellChild0', 'Package'), ('think-cellChild0', 'think-cellXML'), ('think-cellXML',)} <= streams:
            raise RuntimeError('carrier lacks authentic Package child closure')
        return ole.openstream('think-cellXML').read()


def owner(package, target, slide_number):
    matches = []
    for index, slide in enumerate(logical_slides(package), 1):
        rels = relationship_map(package, slide['part'])
        tree = E.fromstring(package.read(slide['part']))
        nodes = [n for n in tree.xpath('.//p:oleObj', namespaces=NS) if rels.get(n.get(R+'id'), {}).get('resolved') == target]
        if nodes:
            # Choice/Fallback duplicates are one owner only when their object identities match.
            shapes = set()
            for n in nodes:
                ancestor = n
                while ancestor is not None and ancestor.tag != '{'+NS['p']+'}graphicFrame':
                    ancestor = ancestor.getparent()
                props = ancestor.find('.//{'+NS['p']+'}cNvPr') if ancestor is not None else None
                shapes.add(props.get('id') if props is not None else None)
            if len(shapes) != 1 or None in shapes:
                raise RuntimeError('carrier has multiple logical chart owners')
            matches.append(index)
    if matches != [slide_number]:
        raise RuntimeError(f'carrier slide ownership mismatch: {matches}')


def workbook_identity(path, hidden, sheet, data):
    info = rebind.named_range(path, hidden, sheet)
    with zipfile.ZipFile(path) as package:
        book = E.fromstring(package.read('xl/workbook.xml'))
        native_names = [n for n in book.iter() if n.tag.endswith('}definedName') and n.get('name', '').startswith('___thinkcell')]
        if len(native_names) != 1:
            raise RuntimeError('workbook has multiple native think-cell aliases')
    if info['hidden'] != '1' or info['range'] != 'A1:C5' or sheet != 'ChartData':
        raise RuntimeError('requires native hidden ChartData!A1:C5 name')
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = list(wb[sheet].iter_rows(min_row=1, max_row=5, min_col=1, max_col=3, values_only=True))
        if list(rows[0][1:]) != data['categories']:
            raise RuntimeError('workbook category labels differ from model')
        labels = [row[0] for row in rows[1:]]
        if len(labels) != 4 or len(set(labels)) != 4 or labels != data['series_names']:
            raise RuntimeError('workbook ordered series labels differ from model')
        values = [list(row[1:]) for row in rows[1:]]
        for actual_row, expected_row in zip(values, data['series_values']):
            for actual, wanted in zip(actual_row, expected_row):
                if isinstance(actual, bool) or not isinstance(actual, (int, float)) or not math.isfinite(actual) or not math.isclose(actual, wanted, rel_tol=1e-12, abs_tol=1e-12):
                    raise RuntimeError('workbook numeric values differ from model')
    finally:
        wb.close()
    return info


def fresh_identity(header, excluded):
    for _ in range(256):
        value = uuid.uuid4()
        link = base64.urlsafe_b64encode(value.bytes).decode().rstrip('=')
        name = '___thinkcell' + base64.b32encode(header + value.bytes_le).decode().rstrip('=')
        if link.isalnum() and not ({str(value), link, name} & excluded):
            return str(value), link, name
    raise RuntimeError('could not allocate a fresh alphanumeric link identity')


def prepare(args):
    inputs = [args.source_presentation.resolve(), args.donor_presentation.resolve(), args.source_workbook.resolve()]
    outputs = [args.output_presentation.resolve(), args.output_workbook.resolve(), args.report.resolve()]
    if len(set(inputs + outputs)) != 6 or any(p.exists() for p in outputs):
        raise RuntimeError('inputs/outputs must be distinct and outputs must be new')
    hashes = [rebind.sha256(p) for p in inputs]
    if hashes != [args.source_sha256.upper(), args.donor_sha256.upper(), args.source_workbook_sha256.upper()]:
        raise RuntimeError('input SHA-256 mismatch')
    if args.sheet_name != 'ChartData' or rebind._norm_ref(args.range) != 'A1:C5':
        raise RuntimeError('only ChartData!A1:C5 is supported')
    with zipfile.ZipFile(inputs[0]) as src, zipfile.ZipFile(inputs[1]) as donor:
        owner(src, args.target_part, args.slide_number)
        owner(donor, args.donor_part, args.donor_slide_number or args.slide_number)
        source_raw, donor_raw = src.read(args.target_part), donor.read(args.donor_part)
        root, table, grid, _, data = profile(read_carrier(source_raw), args.chart_name, False)
        _, _, _, sink, _ = profile(read_carrier(donor_raw), args.chart_name, True)
        moniker = rebind._moniker(sink.findtext('m_vecbMoniker'))[1]
        _, end, old_workbook = composite_file(moniker)
        item = item_name(moniker[end:])
        if item.startswith('!'):
            item = item[1:]
        sheet, hidden = item.split('!', 1)
        if not hidden.startswith('___thinkcell'):
            raise RuntimeError('donor lacks native hidden-name identity')
        encoded = hidden[len('___thinkcell'):]
        native_name = base64.b32decode(encoded + '=' * (-len(encoded) % 8))
        if len(native_name) != 28 or native_name[12:] != uuid.UUID(sink.find('m_guidLink').get('val')).bytes_le:
            raise RuntimeError('donor hidden name does not match adviser GUID')
        workbook_identity(inputs[2], hidden, sheet, data)
        with zipfile.ZipFile(inputs[2]) as wb:
            workbook_xml = E.fromstring(wb.read('xl/workbook.xml'))
            hidden_nodes = workbook_xml.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}definedName')
            if len([n for n in hidden_nodes if n.get('name', '').startswith('___thinkcell')]) != 1:
                raise RuntimeError('workbook has multiple native think-cell aliases')
            excluded = set(workbook_xml.xpath('//@name'))
            for package in (src, donor):
                for part in package.namelist():
                    if part.startswith('ppt/embeddings/') and part.endswith('.bin'):
                        try:
                            model = E.fromstring(rebind._xml(package.read(part)))
                        except (OSError, E.XMLSyntaxError):
                            continue
                        excluded.update(model.xpath('//m_guidLink/@val | //m_lnkid/text()'))
            guid, link, name = fresh_identity(native_name[:12], excluded)
            node = next(n for n in hidden_nodes if n.get('name') == hidden)
            node.set('name', name)
            workbook_parts = {n: wb.read(n) for n in wb.namelist()}
            workbook_parts['xl/workbook.xml'] = E.tostring(workbook_xml, xml_declaration=True, encoding='UTF-8')
        new_sink = copy.deepcopy(sink)
        new_id = str(max(int(n.get('id')) for n in root.iter() if n.get('id')) + 1)
        old_id = new_sink.get('id')
        new_sink.set('id', new_id)
        for n in new_sink.iter():
            if n.get('idref') == old_id:
                n.set('idref', new_id)
        new_sink.find('m_guidLink').set('val', guid)
        new_sink.find('m_lnkid').text = link
        new_sink.find('m_bAutoUpdate').set('val', '1')
        new_sink.find('m_vecbMoniker').text = base64.b64encode(serialize_composite(str(outputs[1]), sheet+'!'+name)).decode().rstrip('=')
        root.append(new_sink)
        one(table, 'm_advisesink').set('idref', new_id)
        E.SubElement(one(grid, 'm_cadvisesink'), 'elem', idref=new_id)
        xml = E.tostring(root, encoding='UTF-8')
        profile(xml, args.chart_name, True)
        expected_sink = E.tostring(new_sink)
        replacement = rebind.replace_xml_stream(source_raw, xml)
        parts = {n: src.read(n) for n in src.namelist()}
        parts[args.target_part] = replacement
    if [rebind.sha256(p) for p in inputs] != hashes:
        raise RuntimeError('input changed during preparation')
    # All validation completes before deliverable paths are created.
    written = []
    try:
        for output, content in zip(outputs[:2], (parts, workbook_parts)):
            output.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(output, 'x', zipfile.ZIP_DEFLATED) as package:
                written.append(output)
                for n, value in content.items():
                    package.writestr(n, value)
        workbook_identity(outputs[1], name, sheet, data)
        with zipfile.ZipFile(outputs[0]) as result:
            actual = read_carrier(result.read(args.target_part))
            _, _, _, readback_sink, _ = profile(actual, args.chart_name, True)
            if E.tostring(readback_sink) != expected_sink:
                raise RuntimeError('adviser closure differs after package readback')
            fields, _, serialized = rebind._fields(actual)
            _, item_offset, actual_workbook = composite_file(serialized)
            if actual_workbook != str(outputs[1]) or item_name(serialized[item_offset:]).lstrip('!') != sheet+'!'+name:
                raise RuntimeError('new workbook/name moniker differs after package readback')
            changed = [n for n in parts if result.read(n) != (source_raw if n == args.target_part else parts[n])]
            if changed != [args.target_part]:
                raise RuntimeError('unexpected changed package parts')
        if [rebind.sha256(p) for p in inputs] != hashes:
            raise RuntimeError('input changed before final readback')
        report = dict(schema='thinkcell-new-excel-link-v1', status='PREPARED_OFFLINE_NEW_EXTERNAL_LINK',
                      source_presentation=str(inputs[0]), output_presentation=str(outputs[0]),
                      output_workbook=str(outputs[1]), selected_part=args.target_part, changed_parts=changed,
                      guid_link=guid, link_id=link, hidden_name=name, adviser_id=new_id, moniker_display_name=str(outputs[1])+'!'+sheet+'!'+name,
                      input_sha256=hashes, output_sha256=[rebind.sha256(p) for p in outputs[:2]],
                      native_certification='NOT_PERFORMED_OFFLINE_ADAPTER_ONLY',
                      guards=dict(model_version=38732, type=0, series=4, categories=2, chart_name=args.chart_name,
                                  slide_number=args.slide_number, table_id=table.get("id"), range="ChartData!A1:C5",
                                  child_storage="think-cellChild0/Package", table_grid_registration=True,
                                  source_values_compatible=True, only_selected_carrier_changed=True))
        outputs[2].parent.mkdir(parents=True, exist_ok=True)
        with outputs[2].open('x', encoding='utf-8') as handle:
            written.append(outputs[2]); json.dump(report, handle, indent=2)
    except Exception:
        for path in written:
            path.unlink(missing_ok=True)
        raise
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ('source-presentation', 'donor-presentation', 'source-workbook', 'output-presentation', 'output-workbook', 'report'):
        parser.add_argument('--'+name, type=Path, required=True)
    for name in ('source-sha256', 'donor-sha256', 'source-workbook-sha256', 'target-part', 'donor-part'):
        parser.add_argument('--'+name, required=True)
    parser.add_argument('--slide-number', type=int, required=True)
    parser.add_argument('--donor-slide-number', type=int)
    parser.add_argument('--chart-name', default='')
    parser.add_argument('--sheet-name', default='ChartData')
    parser.add_argument('--range', default='A1:C5')
    print(json.dumps(prepare(parser.parse_args(argv)), indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
