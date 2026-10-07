"""Inspect native Gantt anchors and prepare experimental endpoint candidates.

Preparation writes model/cache date equality, not certified native reflow.
Prepared packages require native regeneration and changed-date verification.
Automatic reflow needs an authenticated date-scalar anchor profile, currently
unavailable. No lag scheduling or duration preservation is claimed.
"""
from __future__ import annotations

from pathlib import Path
import argparse
import hashlib
import io
import json
import re
import subprocess
import tempfile
import zipfile
from xml.etree import ElementTree as ET

from portable_gantt_adapter import GanttPackage, REPLACE_STREAM, powershell, powershell_env, require_distinct_paths


def inspect_native_anchors(path):
    """Read old and current native bar identities without requiring weekly scales."""
    package = GanttPackage.__new__(GanttPackage)
    package.path = Path(path)
    with zipfile.ZipFile(package.path) as archive:
        package.blobs = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    import olefile
    found = []
    for name, blob in package.blobs.items():
        if not name.startswith('ppt/embeddings/') or not name.endswith('.bin'):
            continue
        with olefile.OleFileIO(io.BytesIO(blob)) as carrier:
            if not carrier.exists('think-cellXML'):
                continue
            model = carrier.openstream('think-cellXML').read().decode('utf8')
            if '<CGanttSE' in model:
                found.append((name, model))
    if len(found) != 1:
        raise ValueError('expected one native Gantt model stream')
    package.ole_part, package.model = found[0]
    root = ET.fromstring(package.model)
    ids = {n.get('id'): n for n in root.iter() if n.get('id')}
    edges = anchor_edges(package.model)
    owners = list(root.iter('CGanttSE'))
    if len(owners) != 1:
        raise ValueError('expected one native Gantt owner')
    consistency = owners[0].find('m_bConsistent')
    bars = []
    for bar in root.iter('CGanttBar'):
        sources = [bar.find('m_varsrc' + side) for side in ('Low', 'High')]
        dates, guids = [], []
        for source in sources:
            if source is None:
                raise ValueError('bar date source is missing')
            value = source.find('m_varval')
            dt = value.find('m_datetime') if value is not None else None
            guid = source.find('m_guid')
            if value is None or value.get('type') != '6' or dt is None or guid is None:
                raise ValueError('bar date source lacks typed date and GUID identity')
            dates.append(dt.get('val')); guids.append(guid.get('val'))
        generic = bar.find('m_pptgenline')
        if generic is not None and generic.get('idref'):
            generic = ids.get(generic.get('idref'))
        ptr = generic.find('m_pptautoshpline') if generic is not None else None
        line = ids.get(ptr.get('idref')) if ptr is not None else None
        rect = line.find('m_rectPPTShape') if line is not None else None
        name = line.findtext('m_bstrShapeName') if line is not None else None
        placed = line.find('m_bPlaced') if line is not None else None
        visible = placed is not None and placed.get('val') == '1'
        if not name or (visible and rect is None):
            raise ValueError('bar native line closure is missing')
        transform = None
        if visible:
            _, transform = package._shape(name)
            if transform is None:
                raise ValueError('bar native physical transform is missing')
        bars.append({'bar_id': bar.get('id'), 'dates': dates, 'date_guids': guids,
                     'line_id': line.get('id'), 'shape_tag': name, 'visible': visible,
                     'model_rect': [int(rect.get(key)) for key in ('left', 'top', 'right', 'bottom')] if rect is not None else None,
                     'physical_transform': transform})
    return {'schema_version': root.find('version').get('val'),
            'consistent': consistency is not None and consistency.get('val') == '1',
            'bars': bars, 'edges': edges}


def anchor_edges(model):
    """Return resolved native edges; reject ambiguous or dangling ownership."""
    root = ET.fromstring(model)
    nodes = [n for n in root.iter() if 'id' in n.attrib]
    by_id = {n.get('id'): n for n in nodes}
    if len(by_id) != len(nodes):
        raise ValueError('duplicate native object identity')
    offered = {}
    for n in nodes:
        for edge in ('Low', 'High'):
            a = n.find('m_anchorOffered' + edge)
            if a is not None and a.get('idref') != '0':
                aid = a.get('idref')
                if aid in offered:
                    raise ValueError('anchor has multiple offering endpoints')
                offered[aid] = (n.get('id'), edge)
    edges = []
    for aid, (parent, parent_edge) in offered.items():
        anchor = by_id.get(aid)
        if anchor is None or anchor.tag != 'CGanttRangeItemAnchor':
            raise ValueError('offered range anchor is missing')
        features = anchor.find('m_cfeature')
        if features is None:
            raise ValueError('range anchor feature collection is missing')
        children = [e.get('idref') for e in features]
        if len(children) != len(set(children)):
            raise ValueError('duplicate attached child')
        for child in children:
            node = by_id.get(child)
            if node is None:
                raise ValueError('attached child is missing')
            attached = [edge for edge in ('Low', 'High')
                        if node.find('m_anchorAttached' + edge) is not None
                        and node.find('m_anchorAttached' + edge).get('idref') == aid]
            if len(attached) != 1:
                raise ValueError('child attachment is not reciprocal and unique')
            edges.append((parent, parent_edge, child, attached[0], aid))
    for n in nodes:
        for edge in ('Low', 'High'):
            a = n.find('m_anchorAttached' + edge)
            if a is not None and a.get('idref') != '0':
                if not any(e[2] == n.get('id') and e[3] == edge and e[4] == a.get('idref') for e in edges):
                    raise ValueError('attachment has no reciprocal offered anchor')
    adjacency = {}
    for parent, _, child, _, _ in edges:
        adjacency.setdefault(parent, []).append(child)
    def visit(node, pending, done):
        if node in pending:
            raise ValueError('cyclic native anchor dependency')
        if node in done:
            return
        for child in adjacency.get(node, []):
            visit(child, pending | {node}, done)
        done.add(node)
    done = set()
    for node in adjacency:
        visit(node, set(), done)
    return edges


def finish_start_model(model, predecessor, successor):
    """Add one reciprocal native range anchor, preserving unrelated XML bytes."""
    if predecessor == successor:
        raise ValueError('dependency endpoints must be distinct')
    anchor_edges(model)
    def block(bid):
        matches = list(re.finditer(r'<CGanttBar id="' + re.escape(str(bid)) + r'".*?</CGanttBar>', model, re.S))
        if len(matches) != 1:
            raise ValueError('bar identity must match exactly once')
        return matches[0].group()
    parent, child = block(predecessor), block(successor)
    if '<m_anchorOfferedHigh' in parent:
        raise ValueError('predecessor already offers an end anchor')
    if '<m_anchorAttachedLow idref="0"/>' not in child:
        raise ValueError('successor start is already attached or unsupported')
    ids = re.findall(r'\bid="([^\"]+)"', model)
    if any(not i.isdecimal() for i in ids):
        raise ValueError('native profile requires numeric object identities')
    aid = str(max(map(int, ids), default=0) + 1)
    parent2 = parent.replace('<m_anchorAttachedHigh idref="0"/>', '<m_anchorAttachedHigh idref="0"/><m_anchorOfferedHigh idref="' + aid + '"/>', 1)
    if parent2 == parent:
        raise ValueError('predecessor end attachment profile is unsupported')
    child2 = child.replace('<m_anchorAttachedLow idref="0"/>', '<m_anchorAttachedLow idref="' + aid + '"/>', 1)
    candidate = model.replace(parent, parent2, 1).replace(child, child2, 1)
    candidate = candidate.replace('</root>', '<CGanttRangeItemAnchor id="' + aid + '"><m_cfeature><elem idref="' + str(successor) + '"/></m_cfeature></CGanttRangeItemAnchor></root>', 1)
    anchor_edges(candidate)
    return candidate, aid


def prepare_finish_start(source, output, predecessor_selector, successor_selector):
    """Align the successor endpoint and create its native finish/start anchor."""
    source, output = Path(source), Path(output)
    if output.resolve() == source.resolve() or output.exists():
        raise ValueError('output must be a fresh path distinct from source')
    package = GanttPackage(source)
    parent = package.resolve_bar(**predecessor_selector)
    child = package.resolve_bar(**successor_selector)
    # Validate topology before preparing any package or changing geometry.
    finish_start_model(package.model, parent.id, child.id)
    endpoint = parent.end[:10]
    if endpoint > child.end[:10]:
        raise ValueError('anchoring would put successor start after its end')
    with tempfile.TemporaryDirectory(dir=output.parent) as folder:
        folder = Path(folder)
        aligned = folder / 'aligned.pptx'
        package.edit_bar(aligned, selector={'bar_id': child.id}, new_start=endpoint, new_end=child.end[:10])
        candidate = GanttPackage(aligned)
        model, aid = finish_start_model(candidate.model, parent.id, child.id)
        # Endpoint anchoring requires geometric equality too. Calendar bins in
        # the date adapter can span a week; use the exact native parent endpoint.
        line = re.search(r'<CPPTAutoShapeLine id="' + re.escape(child.line_id) + r'".*?</CPPTAutoShapeLine>', model, re.S)
        rect = re.search(r'<m_rectPPTShape left="(\d+)" top="(\d+)" right="(\d+)" bottom="(\d+)"', line.group())
        if parent.line_rect[2] >= int(rect.group(3)):
            raise ValueError('anchored successor geometry has no positive width')
        replacement = rect.group().replace('left="' + rect.group(1) + '"', 'left="' + str(parent.line_rect[2]) + '"', 1)
        line2 = line.group().replace(rect.group(), replacement, 1)
        model = model[:line.start()] + line2 + model[line.end():]
        _, parent_transform = package._shape(parent.shape_name)
        child_shape, child_transform = candidate._shape(child.shape_name)
        if parent_transform is None or child_transform is None:
            raise ValueError('native dependency endpoint lacks a physical transform')
        endpoint_x = parent_transform[0] + parent_transform[2]
        right_x = child_transform[0] + child_transform[2]
        if endpoint_x >= right_x:
            raise ValueError('anchored physical successor has no positive width')
        transform = re.search(r'<a:off x="(\d+)" y="(\d+)"/><a:ext cx="(\d+)" cy="(\d+)"', child_shape)
        shape2 = child_shape[:transform.start()] + f'<a:off x="{endpoint_x}" y="{transform.group(2)}"/><a:ext cx="{right_x-endpoint_x}" cy="{transform.group(4)}"' + child_shape[transform.end():]
        candidate.blobs['ppt/slides/slide1.xml'] = candidate.blobs['ppt/slides/slide1.xml'].decode('utf8').replace(child_shape, shape2, 1).encode('utf8')
        carrier, xml = folder / 'carrier.bin', folder / 'model.xml'
        carrier.write_bytes(candidate.blobs[candidate.ole_part])
        xml.write_text(model, encoding='utf8')
        result = subprocess.run([powershell(), '-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', str(REPLACE_STREAM), '-StoragePath', str(carrier), '-StreamBytesPath', str(xml)], capture_output=True, text=True, env=powershell_env(), timeout=60)
        if result.returncode:
            raise RuntimeError(result.stderr or result.stdout)
        candidate.blobs[candidate.ole_part] = carrier.read_bytes()
        with zipfile.ZipFile(output, 'x') as z:
            for entry in candidate.entries:
                z.writestr(entry, candidate.blobs[entry.filename])
    return {'status': 'NATIVE_ENDPOINT_ANCHOR_PREPARED', 'predecessor_bar': parent.id,
            'successor_bar': child.id, 'anchor_id': aid, 'date': endpoint,
            'native_verification_required': True, 'duration_preservation': False}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', required=True)
    parser.add_argument('--inspect', action='store_true')
    parser.add_argument('--expected-sha256')
    parser.add_argument('--output')
    parser.add_argument('--report')
    parser.add_argument('--predecessor-bar')
    parser.add_argument('--successor-bar')
    args = parser.parse_args()
    if args.inspect:
        if any((args.output, args.report, args.predecessor_bar, args.successor_bar)):
            parser.error('inspection does not accept preparation or output arguments')
        source = Path(args.source)
        before = hashlib.sha256(source.read_bytes()).hexdigest()
        if args.expected_sha256 and before.lower() != args.expected_sha256.lower():
            raise ValueError('source SHA256 mismatch')
        result = inspect_native_anchors(source)
        if hashlib.sha256(source.read_bytes()).hexdigest() != before:
            raise RuntimeError('source changed during inspection')
        result.update(status='NATIVE_ANCHOR_STATE_OBSERVED', source_sha256=before,
                      source_unchanged=True, automatic_reflow_certified=False)
        print(json.dumps(result, indent=2))
        return 0
    if not all((args.expected_sha256, args.output, args.report, args.predecessor_bar, args.successor_bar)):
        parser.error('preparation requires expected hash, output, report and both bar identities')
    source, output, report = map(Path, (args.source, args.output, args.report))
    require_distinct_paths(source, output, report)
    if output.exists() or report.exists():
        raise ValueError('output and report must be fresh')
    before = hashlib.sha256(source.read_bytes()).hexdigest()
    if before.lower() != args.expected_sha256.lower():
        raise ValueError('source SHA256 mismatch')
    result = prepare_finish_start(source, output, {'bar_id': args.predecessor_bar}, {'bar_id': args.successor_bar})
    after = hashlib.sha256(source.read_bytes()).hexdigest()
    if after != before:
        raise RuntimeError('source changed during dependency preparation')
    result.update(source=str(source.resolve()), output=str(output.resolve()), source_sha256=before,
                  source_unchanged=True, output_sha256=hashlib.sha256(output.read_bytes()).hexdigest())
    with report.open('x', encoding='utf8') as stream:
        stream.write(json.dumps(result, indent=2) + '\n')
    print(json.dumps(result, indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
