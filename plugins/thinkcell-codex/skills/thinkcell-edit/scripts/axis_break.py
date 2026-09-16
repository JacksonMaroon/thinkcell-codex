"""Safely reposition one existing native think-cell value-axis break.

This adapter changes the `CDataAxisBreak.m_fFraction` field owned by one
exclusive primary value axis.  It does not draw a substitute shape or invent a
break in an arbitrary chart.  The candidate must subsequently go through the
official JSON regeneration and native save/reopen gates, then `verify`.
"""
from __future__ import annotations

import copy
import io
import math
import os
import subprocess
import tempfile
import zipfile
from pathlib import Path

from lxml import etree as E

HERE = Path(__file__).resolve().parent
import sys
sys.path.insert(0, str(HERE))
from chart_geometry import _chart_identity, _select_chart, inventory, need, sha, streams, xml
from runtime import powershell, powershell_env


def _replacement_helper():
    local = HERE / "thinkcell_no_click/implementation/replace_ole_stream.ps1"
    configured = os.environ.get("THINKCELL_IMPLEMENTATION")
    helper = local if local.is_file() else Path(configured or "") / "replace_ole_stream.ps1"
    need(helper.is_file(), "Missing guarded OLE stream helper")
    return helper


def _break(chart):
    owner, ids = chart["owner"], chart["doc"]["ids"]
    need(owner.tag == "CSequenceChartSE", "Only ordinary sequence-chart value axes are supported")
    refs = owner.findall("m_daxisPrimaryValue")
    need(len(refs) == 1 and refs[0].get("idref") in ids, "Missing primary value axis")
    axis = ids[refs[0].get("idref")]
    need(axis.tag == "CSequenceChartDataAxis", "Unexpected primary value-axis type")
    items = axis.findall("m_cdaxisbreak/elem")
    need(len(items) == 1 and items[0].get("idref") in ids,
         "This adapter requires exactly one existing native value-axis break")
    br = ids[items[0].get("idref")]
    need(br.tag == "CDataAxisBreak", "Unexpected native axis-break type")
    fraction = br.find("m_fFraction")
    need(fraction is not None and fraction.get("val") is not None, "Axis-break fraction is unavailable")
    # Repositioning a shared value axis could alter a sibling chart.  The only
    # permitted additional reference is the selected data-table's series group.
    root = chart["doc"]["root"]
    axis_id = axis.get("id")
    selected_group = None
    for ref in root.iter():
        if ref is refs[0] or ref.get("idref") != axis_id:
            continue
        parent = ref.getparent()
        if ref.tag == "m_scdaxis" and parent.tag == "CSequenceChartDataSeriesGroup":
            need(selected_group is None or selected_group is parent, "Primary value axis is shared by multiple series groups")
            selected_group = parent
            continue
        need(False, "Primary value axis is shared outside the selected chart")
    need(selected_group is not None and selected_group.get("id"), "Selected chart has no primary-axis series group")
    group_refs = [ref for ref in root.iter() if ref.get("idref") == selected_group.get("id")]
    table_refs = [ref for ref in chart["table"].iter() if ref.get("idref") == selected_group.get("id")]
    need(len(group_refs) == 1 and group_refs == table_refs, "Primary-axis series group is shared outside selected chart")
    return axis, br, fraction


def make_plan(path, selection, fraction):
    need(not isinstance(fraction, bool) and isinstance(fraction, (int, float)) and math.isfinite(fraction),
         "Axis-break fraction must be a finite number")
    need(0.05 <= fraction <= 0.95, "Axis-break fraction must be between 0.05 and 0.95")
    raw = Path(path).read_bytes()
    _, charts, _ = inventory(raw)
    chart = _select_chart(charts, selection)
    axis, br, old = _break(chart)
    return {
        "schema_version": 1,
        "source_sha256": sha(raw),
        "target": _chart_identity(chart),
        "axis_id": axis.get("id"),
        "axis_break_id": br.get("id"),
        "requested": {"fraction": float(fraction)},
        "edit": {"field": "m_fFraction", "old_string": old.get("val"), "new_value": float(fraction)},
        "required_followup": ["official JSON regeneration", "native save/reopen", "axis-break readback"],
    }


def _assert_only_break_changed(before, after, plan):
    original, candidate = xml(before), xml(after)
    br = candidate.find("./CDataAxisBreak[@id='%s']" % plan["axis_break_id"])
    need(br is not None, "Axis-break identity changed")
    node = br.find(plan["edit"]["field"])
    need(node is not None and math.isclose(float(node.get("val")), plan["edit"]["new_value"], abs_tol=1e-9),
         "Wrong new axis-break fraction")
    node.set("val", plan["edit"]["old_string"])
    need(E.tostring(original, method="c14n") == E.tostring(candidate, method="c14n"), "Non-axis-break model changes")


def prepare(path, output, plan):
    path, output = Path(path), Path(output)
    need(path.resolve() != output.resolve() and not output.exists(), "Distinct new output required")
    expected = make_plan(path, plan["target"], plan["requested"]["fraction"])
    need(plan == expected, "Stale or modified axis-break plan")
    raw = path.read_bytes()
    _, charts, _ = inventory(raw)
    chart = _select_chart(charts, plan["target"])
    axis, br, node = _break(chart)
    need(axis.get("id") == plan["axis_id"] and br.get("id") == plan["axis_break_id"], "Axis-break identity changed")
    before = chart["doc"]["streams"][("think-cellXML",)]
    root = xml(before)
    changed = root.find("./CDataAxisBreak[@id='%s']/%s" % (plan["axis_break_id"], plan["edit"]["field"]))
    need(changed is not None and changed.get("val") == plan["edit"]["old_string"], "Axis-break fraction changed")
    changed.set("val", format(plan["edit"]["new_value"], ".20E"))
    after = E.tostring(root, encoding="utf-8")
    _assert_only_break_changed(before, after, plan)
    doc = chart["doc"]
    with tempfile.TemporaryDirectory(prefix="tc_axis_break_", dir=output.parent) as temporary:
        temporary = Path(temporary)
        carrier, payload = temporary / "carrier.bin", temporary / "model.xml"
        carrier.write_bytes(doc["ole"]); payload.write_bytes(after)
        process = subprocess.run([powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File",
                                  str(_replacement_helper()),
                                  "-StoragePath", str(carrier), "-StreamBytesPath", str(payload)],
                                 capture_output=True, text=True, env=powershell_env(), timeout=60)
        need(process.returncode == 0, "Axis-break preparation failed: " + process.stderr[-500:])
        changed_ole = carrier.read_bytes(); changed_streams = streams(changed_ole)
        need(set(changed_streams) == set(doc["streams"]), "OLE stream inventory changed")
        need(all(value == changed_streams[key] for key, value in doc["streams"].items() if key != ("think-cellXML",)),
             "Other OLE stream changed")
        _assert_only_break_changed(before, changed_streams[("think-cellXML",)], plan)
        need(sha(path.read_bytes()) == plan["source_sha256"], "Source changed")
        with zipfile.ZipFile(io.BytesIO(raw)) as source, zipfile.ZipFile(output, "x") as destination:
            destination.comment = source.comment
            for item in source.infolist():
                destination.writestr(copy.copy(item), changed_ole if item.filename == doc["part"] else source.read(item.filename))
    return {"status": "AXIS_BREAK_PREPARED_REGENERATION_REQUIRED", "source_sha256": plan["source_sha256"],
            "output_sha256": sha(output.read_bytes()), "only_existing_native_axis_break_changed": True}


def verify(path, plan):
    _, charts, _ = inventory(Path(path).read_bytes())
    chart = _select_chart(charts, {"shape_tag": plan["target"]["shape_tag"], "automation_name": plan["target"]["automation_name"]})
    _, br, fraction = _break(chart)
    need(math.isclose(float(fraction.get("val")), plan["requested"]["fraction"], abs_tol=1e-8),
         "Axis-break fraction did not survive regeneration")
    return {"status": "AXIS_BREAK_NATIVE_VERIFIED", "fraction": plan["requested"]["fraction"],
            "axis_break_id_recreated": br.get("id") != plan["axis_break_id"]}
