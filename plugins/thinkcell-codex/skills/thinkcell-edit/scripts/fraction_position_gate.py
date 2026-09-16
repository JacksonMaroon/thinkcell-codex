"""Post-native proof gate for an existing think-cell value-axis break.

This gate is intentionally scoped to the donor profile proved in the v2
canary: a sequence chart with one user break and two native CPPTBreakShape
objects. It proves that the stored fraction survived native reopen and that
the shape geometry moved, rather than accepting metadata-only edits.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path[:0] = [str(HERE), str(HERE / "thinkcell_no_click" / "implementation")]
from prepare_thinkcell_name import inventory, need  # noqa: E402


def snapshot(path, selector):
    _, charts, _ = inventory(Path(path).read_bytes())
    selected = [c for c in charts if c["frames"][0]["shape_tag"] == selector["shape_tag"]]
    need(len(selected) == 1, "Require one tagged native chart")
    chart = selected[0]; ids = chart["doc"]["ids"]
    owner = chart["owner"]
    need(owner.tag == "CSequenceChartSE", "Require a sequence-chart donor")
    axis_ref = owner.find("m_daxisPrimaryValue")
    need(axis_ref is not None and axis_ref.get("idref") in ids, "Missing value axis")
    axis = ids[axis_ref.get("idref")]
    entries = axis.findall("m_cdaxisbreak/elem")
    need(len(entries) == 1 and entries[0].get("idref") in ids, "Require exactly one break")
    br = ids[entries[0].get("idref")]
    need(br.tag == "CDataAxisBreak" and br.find("m_bUser").get("val") == "1", "Require user native break")
    fraction = float(br.find("m_fFraction").get("val"))
    refs = br.findall("m_cpptbreakshp/elem")
    need(len(refs) == 2 and all(x.get("idref") in ids for x in refs), "Require two native shape references")
    shapes = []
    for item in refs:
        node = ids[item.get("idref")]
        need(node.tag == "CPPTBreakShape", "Break reference has wrong type")
        rects = []
        for rect in node.findall(".//m_rectPPTShape"):
            rects.append(tuple(int(rect.get(key)) for key in ("left", "top", "right", "bottom")))
        need(len(rects) == 3, "Expected low, high, and body rectangles")
        shapes.append({"id": node.get("id"), "rectangles": rects})
    return {"fraction": fraction, "axis_id": axis.get("id"), "break_id": br.get("id"), "shapes": shapes}


def verify(before_path, after_path, selector, requested_fraction, allow_unchanged=False):
    before, after = snapshot(before_path, selector), snapshot(after_path, selector)
    need(math.isfinite(requested_fraction) and 0.05 <= requested_fraction <= 0.95,
         "Requested fraction outside guarded range")
    need(math.isclose(after["fraction"], requested_fraction, abs_tol=1e-8), "Native fraction did not persist")
    # Official regeneration can allocate new internal IDs. The tagged chart,
    # exact one-break/two-shape topology, and fraction readback above are the
    # durable identities; record ID recreation rather than rejecting it.
    before_rects = [rect for shape in before["shapes"] for rect in shape["rectangles"]]
    after_rects = [rect for shape in after["shapes"] for rect in shape["rectangles"]]
    need(allow_unchanged or before_rects != after_rects, "Fraction metadata changed without rendered break-shape movement")
    return {"status": "NATIVE_BREAK_POSITION_PROVEN", "before": before, "after": after,
            "requested_fraction": requested_fraction,
            "native_ids_recreated": {"axis": before["axis_id"] != after["axis_id"],
                                     "break": before["break_id"] != after["break_id"]},
            "moved_rectangles": sum(a != b for a, b in zip(before_rects, after_rects))}
