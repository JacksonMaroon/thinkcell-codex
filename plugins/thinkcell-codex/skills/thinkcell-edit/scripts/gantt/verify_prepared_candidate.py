"""Portable offline verification of a prepared edit against its source/report.

This gate proves package readback and preservation, never native certification.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from portable_gantt_adapter import GanttPackage
from milestone_adapter import Package


def verify(source, output, report):
    source, output = Path(source), Path(output)
    evidence = json.loads(Path(report).read_text(encoding='utf8'))
    checks = {}
    before, after = GanttPackage(source), GanttPackage(output)
    edit = evidence['edit']
    checks['source_digest'] = hashlib.sha256(source.read_bytes()).hexdigest() == evidence['source_sha256_before']
    checks['output_identity'] = output.resolve() == Path(evidence['output']).resolve()
    checks['source_identity'] = source.resolve() == Path(evidence['source']).resolve()
    checks['owner_identity'] = before.ole_part == after.ole_part
    checks['package_parts'] = set(before.blobs) == set(after.blobs)
    checks['unrelated_parts'] = all(before.blobs[n] == after.blobs.get(n) for n in before.blobs if n not in (before.ole_part, 'ppt/slides/slide1.xml'))
    old_root, new_root = ET.fromstring(before.model), ET.fromstring(after.model)
    # Native records are keyed by type and ID. Select only the exact records
    # documented by the edit report; every other model record must be identical.
    old_records = {(n.tag, n.get('id')): ET.tostring(n) for n in old_root if n.get('id')}
    new_records = {(n.tag, n.get('id')): ET.tostring(n) for n in new_root if n.get('id')}
    checks['record_identity'] = old_records.keys() == new_records.keys()
    if 'bar_id' in edit:
        old = before.resolve_bar(bar_id=edit['bar_id'])
        new = after.resolve_bar(bar_id=edit['bar_id'])
        allowed = {('CGanttBar', old.id), ('CPPTAutoShapeLine', old.line_id)}
        checks['dates'] = [new.start, new.end] == edit['new_dates']
        checks['model_geometry'] = list(new.line_rect) == edit['new_model_line']
        _, transform = after._shape(new.shape_name)
        checks['visible_geometry'] = [transform[0], transform[0]+transform[2]] == edit['new_visible_bounds']
        checks['binding'] = (new.line_id, new.shape_name) == (old.line_id, old.shape_name)
    else:
        old = Package(source).resolve(milestone_id=edit['milestone_id'])
        new_package = Package(output)
        new = new_package.resolve(milestone_id=edit['milestone_id'])
        allowed = {('CGanttMilestone', old.id)}
        checks['dates'] = new.when == edit['new_date']
        checks['model_geometry'] = list(new.marker_rect) == edit['new_model_bounds']
        _, transform = new_package.visible(new.shape_name)
        checks['visible_geometry'] = transform[0]+transform[2]//2 == edit['new_visible_center']
        checks['binding'] = (new.style, new.shape_name) == (old.style, old.shape_name)
    checks['unrelated_model_records'] = all(value == new_records.get(key) for key, value in old_records.items() if key not in allowed)
    # Exactly one visible shape may change. Replace the selected shape with a
    # sentinel in both XML strings to detect changes to unrelated slide content.
    old_shape, _ = before._shape(old.shape_name)
    new_shape, _ = after._shape(new.shape_name)
    checks['unrelated_slide_content'] = before.blobs['ppt/slides/slide1.xml'].decode().replace(old_shape, '<selected/>', 1) == after.blobs['ppt/slides/slide1.xml'].decode().replace(new_shape, '<selected/>', 1)
    return {'status': 'OFFLINE_GANTT_READBACK_PASS' if all(checks.values()) else 'OFFLINE_GANTT_READBACK_FAIL', 'native_verified': False, 'checks': checks, 'pass': all(checks.values())}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', required=True)
    parser.add_argument('--output', required=True)
    parser.add_argument('--report', required=True)
    args = parser.parse_args()
    try:
        result = verify(args.source, args.output, args.report)
    except (ValueError, KeyError, OSError, ET.ParseError) as error:
        result = {'pass': False, 'native_verified': False, 'error': str(error)}
    print(json.dumps(result, indent=2))
    return 0 if result['pass'] else 2


if __name__ == '__main__':
    raise SystemExit(main())
