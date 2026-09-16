"""Prepare a tightly scoped think-cell primary-axis range change.

This is an intermediate, non-deliverable mutation.  The prepared file must be
regenerated through think-cell's official JSON path and then native-saved,
reopened, audited, and checked for the requested range before it is used.
"""
from __future__ import annotations

import argparse
import copy
import io
import json
import math
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree as E

HERE = Path(__file__).resolve().parent
GEOMETRY = HERE
sys.path.insert(0, str(GEOMETRY))
from chart_geometry import (  # noqa: E402
    _chart_identity,
    _select_chart,
    inventory,
    need,
    sha,
    streams,
    xml,
)
from runtime import powershell, powershell_env  # noqa: E402


def _number(value, name, *, positive=False):
    need(not isinstance(value, bool) and isinstance(value, (int, float)) and math.isfinite(value),
         f"Invalid {name}")
    if positive:
        need(value > 0, f"Invalid {name}")
    return float(value)


def _primary_axis(chart):
    from chart_semantics import kind
    from update_thinkcell_json import model_of
    owner = chart["owner"]
    need(owner.tag == "CSequenceChartSE", "Only sequence-chart primary value axes are supported")
    need(kind(chart)=='CSequenceChartSE' and model_of(chart).get('percent_axis') in (None,False), 'Only ordinary absolute axes are supported')
    refs = owner.findall("m_daxisPrimaryValue")
    need(len(refs) == 1 and refs[0].get("idref"), "Missing or ambiguous primary value-axis reference")
    axis_id = refs[0].get("idref")
    axis = chart["doc"]["ids"].get(axis_id)
    need(axis is not None and axis.tag == "CSequenceChartDataAxis", "Unexpected primary value-axis type")
    need(not axis.findall('m_cdaxisbreak/elem') and owner.find('m_daxisSecondaryValue').get('idref')=='0','Axis breaks and secondary axes require a separate contract')
    # An axis is referenced by the selected owner and its selected data-table's
    # series group.  No other owner may use it, so a range change cannot
    # silently affect a sibling chart.
    root = chart["doc"]["root"]
    selected_group = None
    for ref in root.iter():
        if ref is refs[0] or ref.get("idref") != axis_id:
            continue
        parent = ref.getparent()
        if ref.tag == "m_scdaxis" and parent.tag == "CSequenceChartDataSeriesGroup":
            need(selected_group is None or selected_group is parent,
                 "Primary value axis is shared by multiple series groups")
            selected_group = parent
            continue
        need(False, "Primary value axis is shared outside selected chart")
    need(selected_group is not None and selected_group.get("id"),
         "Selected sequence chart has no primary-axis series group")
    group_refs = [ref for ref in root.iter() if ref.get("idref") == selected_group.get("id")]
    table_refs = [ref for ref in chart["table"].iter() if ref.get("idref") == selected_group.get("id")]
    need(len(group_refs) == 1 and group_refs == table_refs,
         "Primary-axis series group is shared outside selected chart")
    fields = ("m_fMaxValue", "m_fUserMaxValue", "m_fUserScaleUnit")
    result = {}
    for field in fields:
        nodes = axis.findall(field)
        need(len(nodes) == 1 and nodes[0].get("val") is not None, f"Missing or ambiguous {field}")
        result[field] = nodes[0].get("val")
    minimum = axis.findall("m_fMinValue")
    need(len(minimum) == 1 and minimum[0].get("val") is not None,
         "Missing or ambiguous m_fMinValue")
    need(math.isclose(float(minimum[0].get("val")), 0.0, abs_tol=1e-9),
         "Only zero-minimum axes are supported")
    return axis, axis_id, result


def make_plan(path, selection, maximum, major_unit):
    """Create a hash-bound plan for exactly one exclusive primary value axis."""
    maximum = _number(maximum, "maximum", positive=True)
    major_unit = _number(major_unit, "major unit", positive=True)
    need(maximum > major_unit, "Maximum must exceed major unit")
    ratio = maximum / major_unit
    need(math.isclose(ratio, round(ratio), abs_tol=1e-9), "Maximum must be an exact multiple of major unit")
    raw = Path(path).read_bytes()
    docs, charts, _ = inventory(raw)
    chart = _select_chart(charts, selection)
    from update_thinkcell_json import model_of
    values=model_of(chart)['series_values']
    need(all(v is None or v>=0 for row in values for v in row),'Negative values need a separate axis contract')
    need(max(sum(row[i] or 0 for row in values) for i in range(len(values[0])))<=maximum,'Requested maximum would hide plotted data')
    axis, axis_id, old = _primary_axis(chart)
    new = {
        "m_fMaxValue": maximum,
        "m_fUserMaxValue": maximum,
        "m_fUserScaleUnit": major_unit,
    }
    edits = [{"field": field, "old_string": old[field], "new_value": value}
             for field, value in new.items()]
    return {
        "schema_version": 1,
        "source_sha256": sha(raw),
        "target": _chart_identity(chart),
        "axis_id": axis_id,
        "requested": {"minimum": 0.0, "maximum": maximum, "major_unit": major_unit},
        "edits": edits,
        "required_followup": [
            "official JSON regeneration of every named native chart",
            "exact generated data/model and sibling-preservation verification",
            "native save, reopen, strict package audit, and axis readback",
        ],
    }


def _assert_only_axis_changed(before, after, plan):
    original, candidate = xml(before), xml(after)
    axis = candidate.find("./CSequenceChartDataAxis[@id='%s']" % plan["axis_id"])
    need(axis is not None, "Primary axis changed identity")
    for edit in plan["edits"]:
        nodes = axis.findall(edit["field"])
        need(len(nodes) == 1 and math.isclose(float(nodes[0].get("val")), edit["new_value"], abs_tol=1e-9),
             "Wrong new axis value")
        nodes[0].set("val", edit["old_string"])
    need(E.tostring(original, method="c14n") == E.tostring(candidate, method="c14n"),
         "Non-axis model changes")


def prepare(path, output, plan):
    path, output = Path(path), Path(output)
    need(path.resolve() != output.resolve() and not output.exists(), "Distinct new output required")
    need(isinstance(plan.get("target"), dict), "Axis plan is missing selected chart identity")
    expected = make_plan(path, plan["target"], plan["requested"]["maximum"], plan["requested"]["major_unit"])
    need(plan == expected, "Stale or modified axis plan")
    raw = path.read_bytes()
    docs, charts, _ = inventory(raw)
    chart = _select_chart(charts, plan["target"])
    axis, axis_id, old = _primary_axis(chart)
    need(axis_id == plan["axis_id"], "Primary axis identity changed")
    before = chart["doc"]["streams"][("think-cellXML",)]
    root = xml(before)
    changed_axis = root.find("./CSequenceChartDataAxis[@id='%s']" % axis_id)
    need(changed_axis is not None, "Primary axis changed identity")
    for edit in plan["edits"]:
        node = changed_axis.find(edit["field"])
        need(node is not None and node.get("val") == edit["old_string"], "Axis value changed")
        node.set("val", format(edit["new_value"], ".20E"))
    after = E.tostring(root, encoding="utf-8")
    _assert_only_axis_changed(before, after, plan)

    doc = chart["doc"]
    with tempfile.TemporaryDirectory(prefix="tc_axis_", dir=output.parent) as td:
        td = Path(td)
        carrier, payload = td / "carrier.bin", td / "model.xml"
        carrier.write_bytes(doc["ole"])
        payload.write_bytes(after)
        result = subprocess.run(
            [powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File",
             str(GEOMETRY / "thinkcell_no_click/implementation/replace_ole_stream.ps1"),
             "-StoragePath", str(carrier), "-StreamBytesPath", str(payload)],
            capture_output=True, text=True, env=powershell_env(), timeout=60,
        )
        need(result.returncode == 0, "Axis preparation failed: " + result.stderr[-500:])
        carrier_bytes = carrier.read_bytes()
        new_streams = streams(carrier_bytes)
        need(set(new_streams) == set(doc["streams"]), "OLE stream inventory changed")
        need(all(value == new_streams[key] for key, value in doc["streams"].items()
                 if key != ("think-cellXML",)), "Other OLE stream changed")
        _assert_only_axis_changed(before, new_streams[("think-cellXML",)], plan)
        need(sha(path.read_bytes()) == plan["source_sha256"], "Source changed")
        with zipfile.ZipFile(io.BytesIO(raw)) as src, zipfile.ZipFile(output, "x") as dst:
            dst.comment = src.comment
            for item in src.infolist():
                dst.writestr(copy.copy(item), carrier_bytes if item.filename == doc["part"] else src.read(item.filename))
        with zipfile.ZipFile(io.BytesIO(raw)) as src, zipfile.ZipFile(output) as dst:
            need(src.namelist() == dst.namelist() and dst.testzip() is None, "ZIP inventory or CRC changed")
            need(all(src.read(name) == dst.read(name) for name in src.namelist() if name != doc["part"]),
                 "Other package parts changed")
    return {
        "status": "AXIS_RANGE_PREPARED_REGENERATION_REQUIRED",
        "source_sha256": plan["source_sha256"],
        "output_sha256": sha(output.read_bytes()),
        "only_exclusive_primary_axis_fields_changed": [edit["field"] for edit in plan["edits"]],
    }

def verify(path, plan):
    _,cs,_=inventory(Path(path).read_bytes())
    c=_select_chart(cs,{'shape_tag':plan['target']['shape_tag'],'automation_name':plan['target']['automation_name']})
    axis,_,fields=_primary_axis(c)
    for e in plan['edits']:need(math.isclose(float(fields[e['field']]),e['new_value'],abs_tol=1e-8),'Axis range did not survive regeneration')
    ns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
    with zipfile.ZipFile(path) as z:root=E.fromstring(z.read(c['frames'][0]['native_chart_part']))
    maxima=root.xpath('.//c:valAx/c:scaling/c:max/@val',namespaces=ns)
    need(len(maxima)==1 and math.isclose(float(maxima[0]),plan['requested']['maximum'],abs_tol=1e-8),'Native cache maximum differs from request')
    return {'status':'AXIS_NATIVE_VERIFIED','range':plan['requested']}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    make = commands.add_parser("make-plan")
    make.add_argument("--input", required=True)
    make.add_argument("--selection-json", required=True)
    make.add_argument("--maximum", required=True, type=float)
    make.add_argument("--major-unit", required=True, type=float)
    make.add_argument("--plan-out", required=True)
    run = commands.add_parser("prepare")
    run.add_argument("--input", required=True)
    run.add_argument("--plan", required=True)
    run.add_argument("--output", required=True)
    args = parser.parse_args()
    if args.command == "make-plan":
        plan = make_plan(args.input, json.loads(Path(args.selection_json).read_text(encoding="utf-8")),
                         args.maximum, args.major_unit)
        target = Path(args.plan_out)
        need(not target.exists(), "Plan output exists")
        target.write_text(json.dumps(plan, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(plan, indent=2))
    else:
        plan = json.loads(Path(args.plan).read_text(encoding="utf-8"))
        print(json.dumps(prepare(args.input, args.output, plan), indent=2))


if __name__ == "__main__":
    main()
