"""Portable offline metric-rail adapter for detached think-cell legends.

The implementation discovers the installed think-cell scripts from the
package's sibling scripts rail.  It performs no Office calls.
"""
from __future__ import annotations

import copy
import io
import math
import subprocess
import tempfile
import zipfile
from pathlib import Path

from lxml import etree as E


HERE = Path(__file__).resolve()
_script_candidates = []
for ancestor in (HERE.parent, *HERE.parents):
    _script_candidates.extend((ancestor, ancestor / "scripts",
                               ancestor / "skills/thinkcell-edit/scripts"))
SCRIPTS = next((p.resolve() for p in _script_candidates
                if ((p / "chart_geometry.py").is_file() and
                    (p / "thinkcell_no_click/implementation/prepare_thinkcell_name.py").is_file())), None)
if SCRIPTS is None:
    raise RuntimeError("Installed think-cell scripts were not found beside this package")
IMPL = SCRIPTS / "thinkcell_no_click" / "implementation"
import sys
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(IMPL))

from chart_geometry import _chart_identity, _select_chart, tagged_bounds  # noqa: E402
from prepare_thinkcell_name import inventory, need, sha, streams, xml  # noqa: E402
from runtime import powershell, powershell_env  # noqa: E402


SIDES = ("left", "top", "right", "bottom")
LEGEND_TYPES = {"CSequenceChartLegendSE", "CScatterChartLegendSE", "CPieChartLegendSE"}


def _active(node, version):
    return int(node.get("reqver", "0")) <= version < int(node.get("endver", "999999"))


def _selected(path, selector):
    raw = Path(path).read_bytes()
    _, charts, _ = inventory(raw)
    chart = _select_chart(charts, selector)
    return raw, chart


def _legend(root, ids, owner_id, selector, version):
    candidates = [n for n in root if n.tag in LEGEND_TYPES
                  and n.find("m_cse") is not None
                  and n.find("m_cse").get("idref") == owner_id]
    if selector and selector.get("legend_id") not in (None, ""):
        candidates = [n for n in candidates if str(n.get("id")) == str(selector["legend_id"])]
    need(len(candidates) == 1, "Legend selector is missing, ambiguous, or changed")
    legend = candidates[0]
    rect_ref = legend.find("m_pptrect")
    need(rect_ref is not None and rect_ref.get("idref") in ids, "Legend frame is unresolved")
    if selector:
        need(selector.get("owner_id") in (None, "", owner_id), "Legend owner changed")
        need(selector.get("pptrect_id") in (None, "", rect_ref.get("idref")),
             "Legend frame selector changed")
    model = []
    for i, ref in enumerate(legend.findall("m_agrdlnanchor/elem")):
        anchor = ids.get(ref.get("idref"))
        need(anchor is not None, "Legend anchor is unresolved")
        bindings = [b for b in anchor.findall("m_grdlnBinding") if _active(b, version)]
        need(len(bindings) == 1, "Legend anchor binding is ambiguous")
        grid_id = bindings[0].get("idref")
        grid = ids.get(grid_id)
        value = grid.find("m_gveps/m_gvValue") if grid is not None else None
        need(grid is not None and value is not None and math.isfinite(float(value.get("val"))),
             "Legend coordinate is invalid")
        model.append({"anchor_id": anchor.get("id"), "grid_id": grid_id, "side": SIDES[i],
                      "old_string": value.get("val"), "value": float(value.get("val")) / 8.0})
    need(len(model) == 4, "Legend does not have four model anchors")
    tags = []
    for ref in legend.findall("m_clegendentry/elem"):
        entry = ids.get(ref.get("idref"))
        if entry is None:
            continue
        entry_rect = entry.find("m_pptrect")
        placed = ids.get(entry_rect.get("idref")) if entry_rect is not None else None
        entry_name = placed.findtext("m_bstrShapeName") if placed is not None else None
        if entry_name and entry_name not in tags:
            tags.append(entry_name)
        for n in entry.iter("m_bstrShapeName"):
            if n.text and n.text not in tags:
                tags.append(n.text)
        label_ref = entry.find("m_legendlabel")
        label = ids.get(label_ref.get("idref")) if label_ref is not None else None
        label_name = label.findtext("m_ppttb/m_bstrShapeName") if label is not None else None
        if label_name and label_name not in tags:
            tags.append(label_name)
    outer_name = ids[rect_ref.get("idref")].findtext("m_bstrShapeName")
    if outer_name and outer_name not in tags:
        tags.insert(0, outer_name)
    return legend, model, tags, rect_ref.get("idref")


def metric_anchors(path, selector=None, legend_selector=None):
    _, chart = _selected(path, selector)
    doc = chart["doc"]
    root, ids = doc["root"], doc["ids"]
    version = int(root.find("version").get("val"))
    legend, model, _, _ = _legend(root, ids, chart["owner"].get("id"), legend_selector, version)
    node = legend.find("m_agrdlnanchorText")
    need(node is not None, "Legend has no text-anchor rail")
    out = []
    for i, elem in enumerate(node.findall("elem")):
        binding = next((x for x in elem.findall("m_grdlnBinding") if _active(x, version)), None)
        need(binding is not None, "Text anchor has no active native binding")
        grid = ids.get(binding.get("idref"))
        val = grid.find("m_gveps/m_gvValue") if grid is not None else None
        need(val is not None and math.isfinite(float(val.get("val"))), "Text anchor grid is invalid")
        out.append({"grid_id": binding.get("idref"), "side": SIDES[i],
                    "old_string": val.get("val"), "value": float(val.get("val")) / 8.0,
                    "metric_offset": float(val.get("val")) / 8.0 - model[i]["value"]})
    need(len(out) == 4, "Text anchor rail is incomplete")
    return out


def _physical(path, chart, tags):
    bounds = tagged_bounds(path, set(tags))
    need(bounds is not None, "Tagged physical legend union is unavailable")
    return [float(x) for x in bounds]


def _constraints(root, selected):
    constraints = []
    selected = set(selected)
    for node in root.iter():
        if "constraint" not in E.QName(node).localname.lower():
            continue
        refs = sorted({x.get("idref") for x in node.iter()
                       if x.get("idref") not in (None, "0")})
        if refs:
            constraints.append({"tag": E.QName(node).localname, "id": node.get("id"),
                                "idrefs": refs})
    constrained = {ref for item in constraints for ref in item["idrefs"]}
    return {"active": bool(constraints), "constraints": constraints,
            "shared_coordinate_refs": [], "editable": not bool(selected & constrained),
            "shared_layout_compatible": False}


def inspect(path, selector=None, legend_selector=None):
    raw, chart = _selected(path, selector)
    doc = chart["doc"]
    root, ids = doc["root"], doc["ids"]
    version = int(root.find("version").get("val"))
    legend, model, tags, rect_id = _legend(root, ids, chart["owner"].get("id"), legend_selector, version)
    physical = _physical(path, chart, tags)
    consistent = lambda node: (node.find("m_bConsistent") is not None and
                               node.find("m_bConsistent").get("val") == "1")
    constraints = _constraints(root, [x["grid_id"] for x in model])
    return {"source_sha256": sha(raw), "target": _chart_identity(chart),
            "legend": {"id": legend.get("id"), "type": legend.tag,
                       "owner_id": chart["owner"].get("id"), "pptrect_id": rect_id,
                       "shape_tags": tags, "anchors": model,
                       "entry_order": [x.get("idref") for x in legend.findall("m_clegendentry/elem")],
                       "sort_mode": (legend.find("m_elegendsortmode").get("val")
                                     if legend.find("m_elegendsortmode") is not None else None),
                       "adopted_size": ({k: legend.find("m_szgvAdopted").get(k)
                                         for k in ("cx", "cy") if legend.find("m_szgvAdopted") is not None}),
                       "bounds": [x["value"] for x in model]},
            "physical": {"legend_bounds": physical},
            "consistency": {"chart": consistent(chart["owner"]), "legend": consistent(legend)},
            "constraints": constraints, "same_carrier": True,
            "native_group_nesting": False, "model_version": version}


def make_plan(path, selector, legend_selector, translation):
    base = inspect(path, selector, legend_selector)
    need(base["constraints"]["editable"], "Selected legend participates in a shared constraint")
    dx, dy = float(translation["x"]), float(translation["y"])
    old = base["legend"]["bounds"]
    expected = [old[0] + dx, old[1] + dy, old[2] + dx, old[3] + dy]
    edits = [dict(a, new_value=v * 8.0) for a, v in zip(base["legend"]["anchors"], expected)]
    text = metric_anchors(path, selector, legend_selector)
    base["operation"] = {"translation": {"x": dx, "y": dy}, "bounds": None}
    base["old_bounds"] = old
    base["expected_bounds"] = expected
    base["old_physical_bounds"] = base["physical"]["legend_bounds"]
    base["edits"] = edits
    base["same_carrier"] = True
    # Keep the established metric-plan schema; readback metadata is available
    # through inspect() and does not belong in the immutable edit plan.
    base.pop("consistency", None)
    base.pop("model_version", None)
    base["metric_anchor"] = {"method": "native_text_grid_minus_model_grid", "anchors": text,
                              "expected_text_bounds": [text[0]["value"] + dx,
                                                        text[1]["value"] + dy,
                                                        text[2]["value"] + dx,
                                                        text[3]["value"] + dy]}
    base.pop("physical", None)
    base["status"] = "NATIVE_LEGEND_PLAN_REVIEW_REQUIRED"
    base["legend_selector"] = {"legend_id": base["legend"]["id"],
                                "owner_id": base["legend"]["owner_id"],
                                "pptrect_id": base["legend"]["pptrect_id"]}
    base.pop("legend", None)
    base["requires_official_regeneration"] = True
    base["requires_native_reopen"] = True
    base["requires_visual_review"] = True
    base["physical_gate"] = {"baseline_bounds": base["old_physical_bounds"],
                              "expected_bounds": base["old_physical_bounds"],
                              "tolerance_pt": 0.05,
                              "source_method": "tagged_physical_legend_union"}
    return base


def prepare(path, output, plan):
    path, output = Path(path), Path(output)
    need(not output.exists() and path.resolve() != output.resolve(), "Distinct new output required")
    need(sha(path.read_bytes()) == plan["source_sha256"], "Source hash changed; inspect again")
    raw, chart = _selected(path, plan["target"])
    doc = chart["doc"]
    root = xml(doc["streams"][("think-cellXML",)])
    for edit in plan["edits"]:
        n = root.find("./CGridline[@id='%s']/m_gveps/m_gvValue" % edit["grid_id"])
        need(n is not None and n.get("val") == edit["old_string"], "Legend grid identity changed")
        n.set("val", format(edit["new_value"], ".20E"))
    for edit in plan["metric_anchor"]["anchors"]:
        n = root.find("./CGridline[@id='%s']/m_gveps/m_gvValue" % edit["grid_id"])
        need(n is not None and n.get("val") == edit["old_string"], "Text grid identity changed")
        delta = plan["operation"]["translation"]["x"] if edit["side"] in ("left", "right") else plan["operation"]["translation"]["y"]
        n.set("val", format((float(edit["old_string"]) / 8.0 + delta) * 8.0, ".20E"))
    payload = E.tostring(root, encoding="utf-8")
    with tempfile.TemporaryDirectory(prefix="tc_legend_portable_", dir=output.parent) as td:
        carrier, model = Path(td) / "carrier.bin", Path(td) / "model.xml"
        carrier.write_bytes(doc["ole"]); model.write_bytes(payload)
        proc = subprocess.run([powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File",
                               str(IMPL / "replace_ole_stream.ps1"), "-StoragePath", str(carrier),
                               "-StreamBytesPath", str(model)], capture_output=True, text=True,
                              env=powershell_env(), timeout=60)
        need(proc.returncode == 0, "Metric stream replacement failed")
        changed = carrier.read_bytes()
        with zipfile.ZipFile(io.BytesIO(raw)) as src, zipfile.ZipFile(output, "x") as dst:
            for item in src.infolist():
                dst.writestr(copy.copy(item), changed if item.filename == doc["part"] else src.read(item.filename))
    return {"status": "NATIVE_LEGEND_METRIC_PREPARED_REGENERATION_REQUIRED",
            "output": str(output), "plan": str(plan)}


def _core(plan):
    out = copy.deepcopy(plan)
    out.pop("physical_gate", None)
    out.pop("data_requests", None)
    out.pop("baseline_model_version", None)
    return out


def grade(path, plan):
    actual = inspect(path, {**plan["target"], "shape_id": None}, plan["legend_selector"])
    text = metric_anchors(path, {**plan["target"], "shape_id": None}, plan["legend_selector"])
    model_ok = all(math.isclose(a, b, abs_tol=.01)
                   for a, b in zip(actual["legend"]["bounds"], plan["expected_bounds"]))
    text_ok = all(math.isclose(a["value"], b, abs_tol=.01)
                  for a, b in zip(text, plan["metric_anchor"]["expected_text_bounds"]))
    return {"status": "PASS" if model_ok and text_ok else "FAIL",
            "model_bounds": actual["legend"]["bounds"], "expected_model_bounds": plan["expected_bounds"],
            "text_bounds": [x["value"] for x in text],
            "expected_text_bounds": plan["metric_anchor"]["expected_text_bounds"],
            "metric_offsets": [x["metric_offset"] for x in text],
            "consistency": actual["consistency"]}
