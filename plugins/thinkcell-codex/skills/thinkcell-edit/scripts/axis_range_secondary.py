"""Prepare a guarded native think-cell value-axis range change.

This production candidate keeps the original primary-axis route as the
default and adds a secondary-axis route for ordinary sequence charts.  It
edits only existing axis scalar fields; official JSON regeneration and native
save/reopen remain required follow-up operations.
"""
from __future__ import annotations

import argparse
import copy
import io
import json
import math
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree as E


HERE = Path(__file__).resolve().parent
_requested_scripts = os.environ.get("THINKCELL_SCRIPT_DIR")
_roots = [Path(_requested_scripts).resolve()] if _requested_scripts else []
_roots.extend([HERE, HERE.parent])
SCRIPT_DIR = next((root for root in _roots if (root / "chart_geometry.py").is_file()), None)
if SCRIPT_DIR is None:
    raise ImportError("chart_geometry.py is required beside this adapter or in THINKCELL_SCRIPT_DIR")
IMPLEMENTATION_DIR = SCRIPT_DIR / "thinkcell_no_click" / "implementation"
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(IMPLEMENTATION_DIR))

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


def _selected_axis(chart, axis_role="primary"):
    """Return the requested value axis and native series indices.

    Primary retains the previous route's exclusive-primary restriction.  The
    secondary route requires the owner reference and resolves its series from
    actual native group/index references.
    """
    from chart_semantics import kind
    from update_thinkcell_json import model_of

    owner = chart["owner"]
    need(owner.tag == "CSequenceChartSE", "Only sequence-chart value axes are supported")
    need(kind(chart) == "CSequenceChartSE" and model_of(chart).get("percent_axis") in (None, False),
         "Only ordinary absolute axes are supported")
    need(axis_role in {"primary", "secondary"}, "Axis role must be primary or secondary")

    primary_refs = owner.findall("m_daxisPrimaryValue")
    secondary_refs = owner.findall("m_daxisSecondaryValue")
    need(len(primary_refs) == 1 and primary_refs[0].get("idref"),
         "Missing or ambiguous primary value-axis reference")
    need(len(secondary_refs) == 1, "Missing or ambiguous secondary value-axis reference")
    if axis_role == "primary":
        need(secondary_refs[0].get("idref") == "0",
             "Secondary axes require axis_role=secondary")
        reference = primary_refs[0]
    else:
        need(secondary_refs[0].get("idref") not in (None, "0"),
             "Chart has no secondary value axis")
        reference = secondary_refs[0]

    axis_id = reference.get("idref")
    axis = chart["doc"]["ids"].get(axis_id)
    need(axis is not None and axis.tag == "CSequenceChartDataAxis",
         f"Unexpected {axis_role} value-axis type")
    need(not axis.findall("m_cdaxisbreak/elem"),
         "Axis breaks require a separate axis contract")

    root = chart["doc"]["root"]
    owner_refs = set(primary_refs + secondary_refs)
    groups = []
    for ref in root.iter():
        if ref.get("idref") != axis_id:
            continue
        parent = ref.getparent()
        if ref in owner_refs:
            continue
        if ref.tag == "m_scdaxis" and parent is not None and parent.tag == "CSequenceChartDataSeriesGroup":
            groups.append(parent)
            continue
        raise ValueError(f"Selected {axis_role} value axis is shared outside selected chart")
    need(groups, f"Selected {axis_role} value axis has no native series group")
    group_ids = {group.get("id") for group in groups}
    need(None not in group_ids, "Selected series group has no native ID")
    group_refs = [ref for ref in root.iter() if ref.get("idref") in group_ids]
    table_refs = [ref for ref in chart["table"].iter() if ref.get("idref") in group_ids]
    need(len(group_refs) == len(table_refs) and all(ref in table_refs for ref in group_refs),
         "Selected series group is shared outside the selected chart")

    series_count = len(model_of(chart)["series_values"])
    indices = []
    for group in groups:
        refs = group.findall("m_vecnIndex/elem")
        need(refs, "Selected series group has no native series indices")
        for ref in refs:
            raw = ref.get("val")
            need(raw is not None and raw.lstrip("-").isdigit(), "Series index is not an integer")
            index = int(raw)
            need(0 <= index < series_count, "Series index is outside the data model")
            indices.append(index)
    need(indices and len(indices) == len(set(indices)),
         f"Selected {axis_role} axis has ambiguous series ownership")

    fields = {}
    for field in ("m_fMinValue", "m_fMaxValue", "m_fUserScaleUnit", "m_fUserMaxValue"):
        nodes = axis.findall(field)
        need(len(nodes) == 1 and nodes[0].get("val") is not None, f"Missing or ambiguous {field}")
        fields[field] = nodes[0].get("val")
    if axis_role == "primary":
        need(math.isclose(float(fields["m_fMinValue"]), 0.0, abs_tol=1e-9),
             "Only zero-minimum primary axes are supported")
    return axis, axis_id, fields, sorted(indices)


def make_plan(path, selection, maximum, major_unit, *, axis_role="primary"):
    """Create a hash-bound plan for one exclusive value axis."""
    maximum = _number(maximum, "maximum", positive=True)
    major_unit = _number(major_unit, "major unit", positive=True)
    need(maximum > major_unit, "Maximum must exceed major unit")
    need(math.isclose(maximum / major_unit, round(maximum / major_unit), abs_tol=1e-9),
         "Maximum must be an exact multiple of major unit")
    raw = Path(path).read_bytes()
    _, charts, _ = inventory(raw)
    chart = _select_chart(charts, selection)
    from update_thinkcell_json import model_of
    model = model_of(chart)
    values = model["series_values"]
    axis, axis_id, old, indices = _selected_axis(chart, axis_role)
    need(all(v is None or v >= 0 for row in values for v in row),
         "Negative values need a separate axis contract")
    observed_max = max(sum(values[index][column] or 0 for index in indices)
                       for column in range(len(values[0])))
    need(observed_max <= maximum, "Requested maximum would hide plotted data")
    new = {"m_fMaxValue": maximum, "m_fUserMaxValue": maximum, "m_fUserScaleUnit": major_unit}
    edits = [{"field": field, "old_string": old[field], "new_value": value}
             for field, value in new.items()]
    return {
        "schema_version": 2,
        "source_sha256": sha(raw),
        "target": _chart_identity(chart),
        "axis_id": axis_id,
        "axis_role": axis_role,
        "series_indices": indices,
        "requested": {"minimum": 0.0, "maximum": maximum, "major_unit": major_unit},
        "edits": edits,
        "native_model_contract": {
            "axis_tag": axis.tag,
            "existing_axis_child_order": [child.tag for child in axis],
            "secondary_axis_ref_present": chart["owner"].find("m_daxisSecondaryValue").get("idref") not in (None, "0"),
        },
        "required_followup": [
            "official JSON regeneration of every named native chart",
            "exact generated data/model and sibling-preservation verification",
            "native save, reopen, strict package audit, and axis readback",
        ],
    }


def _assert_only_axis_changed(before, after, plan):
    original, candidate = xml(before), xml(after)
    axis = candidate.find(f"./CSequenceChartDataAxis[@id='{plan['axis_id']}']")
    need(axis is not None, "Selected axis changed identity")
    for edit in plan["edits"]:
        nodes = axis.findall(edit["field"])
        need(len(nodes) == 1 and math.isclose(float(nodes[0].get("val")), edit["new_value"], abs_tol=1e-9),
             f"Wrong new axis value for {edit['field']}")
        nodes[0].set("val", edit["old_string"])
    need(E.tostring(original, method="c14n") == E.tostring(candidate, method="c14n"),
         "Non-axis model changes or XML order drift")


def prepare(path, output, plan):
    path, output = Path(path), Path(output)
    need(path.resolve() != output.resolve() and not output.exists(), "Distinct new output required")
    axis_role = plan.get("axis_role", "primary")
    expected = make_plan(path, plan["target"], plan["requested"]["maximum"],
                         plan["requested"]["major_unit"], axis_role=axis_role)
    need(plan == expected, "Stale or modified axis plan")
    raw = path.read_bytes()
    _, charts, _ = inventory(raw)
    chart = _select_chart(charts, plan["target"])
    before = chart["doc"]["streams"][("think-cellXML",)]
    root = xml(before)
    axis = root.find(f"./CSequenceChartDataAxis[@id='{plan['axis_id']}']")
    need(axis is not None, "Selected axis changed identity")
    for edit in plan["edits"]:
        node = axis.find(edit["field"])
        need(node is not None and node.get("val") == edit["old_string"], "Axis value changed before preparation")
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
             str(IMPLEMENTATION_DIR / "replace_ole_stream.ps1"),
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
        with zipfile.ZipFile(io.BytesIO(raw)) as src, zipfile.ZipFile(output, "x") as dst:
            dst.comment = src.comment
            for item in src.infolist():
                dst.writestr(copy.copy(item), carrier_bytes if item.filename == doc["part"] else src.read(item.filename))
        with zipfile.ZipFile(io.BytesIO(raw)) as src, zipfile.ZipFile(output) as dst:
            need(src.namelist() == dst.namelist() and dst.testzip() is None, "ZIP inventory or CRC changed")
            need(all(src.read(name) == dst.read(name) for name in src.namelist() if name != doc["part"]),
                 "Other package parts changed")
    return {"status": "AXIS_RANGE_SECONDARY_PREPARED_REGENERATION_REQUIRED",
            "source_sha256": plan["source_sha256"], "output_sha256": sha(output.read_bytes()),
            "axis_role": axis_role, "series_indices": plan["series_indices"]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    make = commands.add_parser("make-plan")
    make.add_argument("--input", required=True)
    make.add_argument("--selection-json", required=True)
    make.add_argument("--maximum", required=True, type=float)
    make.add_argument("--major-unit", required=True, type=float)
    make.add_argument("--axis-role", choices=("primary", "secondary"), default="primary")
    make.add_argument("--plan-out", required=True)
    run = commands.add_parser("prepare")
    run.add_argument("--input", required=True)
    run.add_argument("--plan", required=True)
    run.add_argument("--output", required=True)
    args = parser.parse_args()
    if args.command == "make-plan":
        plan = make_plan(args.input, json.loads(Path(args.selection_json).read_text(encoding="utf-8")),
                         args.maximum, args.major_unit, axis_role=args.axis_role)
        target = Path(args.plan_out)
        need(not target.exists(), "Plan output exists")
        target.write_text(json.dumps(plan, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(plan, indent=2))
    else:
        plan = json.loads(Path(args.plan).read_text(encoding="utf-8"))
        print(json.dumps(prepare(args.input, args.output, plan), indent=2))


if __name__ == "__main__":
    main()
