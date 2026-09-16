"""Dedicated readback for the exploratory two-break candidate."""
from __future__ import annotations
import json, math, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
import run_variant_control as variant


def need(condition, message):
    if not condition:
        raise RuntimeError(message)


def verify(path: Path, plan: dict, expected_shape_counts=None, expected_gaps=None):
    data = plan["targets"][0]["data"]
    chart, _ = variant.target_any(path, plan)
    actual = variant.runner.model_of(chart)
    expected = data["expected_model"]
    for key in ("categories", "series_names", "series_values", "category_extents"):
        need(variant.runner.equal(actual.get(key), expected[key]), "exact model mismatch: " + key)
    sheets = variant.runner.read_sheet(path, 1, chart["owner_name"])["sheets"]
    cells = {(x["row"], x["column"]): x["value"] for x in sheets[0]["nonempty_cells"]}
    expected_cells = {(r + 1, c + 1): v for r, row in enumerate(data["matrix"]) for c, v in enumerate(row) if v not in (None, "")}
    need(cells == expected_cells, "exact datasource cells mismatch")
    ids = chart["doc"]["ids"]
    axis = ids[chart["owner"].find("m_daxisPrimaryValue").get("idref")]
    refs = axis.findall("m_cdaxisbreak/elem")
    need(len(refs) == 2, "two native axis-break owners required")
    expected_shape_counts = expected_shape_counts or [1] * len(refs)
    need(len(expected_shape_counts) == len(refs), "shape-count expectation must cover every break")
    breaks = []
    all_names = []
    for ref in refs:
        br = ids[ref.get("idref")]
        need(br.tag == "CDataAxisBreak", "axis-break type changed")
        fraction = br.find("m_fFraction"); user = br.find("m_bUser"); ordinal = br.find("m_nOrdinal")
        need(fraction is not None and user is not None and user.get("val") == "1" and ordinal is not None, "break ownership fields incomplete")
        shapes = [ids[x.get("idref")] for x in br.findall("m_cpptbreakshp/elem")]
        need(len(shapes) >= 1 and all(shape.tag == "CPPTBreakShape" for shape in shapes), "break physical shape closure is missing")
        names = [shape.find(field + "/m_bstrShapeName").text for shape in shapes for field in variant.runner.FIELDS]
        need(len(names) == 3 * len(shapes) and len(set(names)) == len(names) and all(names), "break physical tags incomplete")
        all_names.extend(names)
        breaks.append({"ordinal": int(ordinal.get("val")), "fraction": float(fraction.get("val")), "shape_count": len(shapes)})
    need(sorted(x["ordinal"] for x in breaks) == [0, 1], "break ordinals are not ordered 0,1")
    need(len(set(all_names)) == len(all_names) and len(all_names) == 3 * sum(expected_shape_counts), "break physical tags are not unique or shape closure count differs")
    need([x["shape_count"] for x in sorted(breaks, key=lambda x: x["ordinal"])] == list(expected_shape_counts), "native break shape closure differs from expected category crossings")
    variant.runner.physical_tags(path.read_bytes(), all_names)
    vector = axis.findall("m_vecintvlOrdinal")
    need(len(vector) == 1 and len(vector[0].findall("elem")) == 12, "unexpected ordered interval partition")
    gaps = []
    for elem in vector[0].findall("elem"):
        begin, end = elem.find("begin"), elem.find("end")
        if begin is not None and end is not None:
            gaps.append((float(begin.get("val")), float(end.get("val"))))
    expected_gaps = [tuple(x) for x in expected_gaps] if expected_gaps is not None else [tuple(x["gap"]) for x in sorted(data["expected_breaks"], key=lambda x: x["ordinal"])]
    need(len(expected_gaps) == len(breaks), "expected gaps must cover every break")
    # Native think-cell associates a break owner with the interval vector by
    # m_nOrdinal. Searching for a matching pair anywhere in the partition can
    # mistake a visible interval for the omitted gap.
    for item, gap in zip(sorted(breaks, key=lambda x: x["ordinal"]), expected_gaps):
        ordinal = item["ordinal"]
        need(ordinal < len(gaps) and math.isclose(gap[0], gaps[ordinal][0], abs_tol=1e-8) and math.isclose(gap[1], gaps[ordinal][1], abs_tol=1e-8), "expected omitted gap missing at its ordinal interval slot")
    return {"status": "TWO_BREAK_READBACK_PASS", "breaks": sorted(breaks, key=lambda x: x["ordinal"]), "expected_gaps": expected_gaps, "physical_tag_count": len(all_names), "interval_count": len(vector[0].findall("elem")), "expected_shape_counts": list(expected_shape_counts)}


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--input", type=Path, required=True)
    p.add_argument("--plan", type=Path, required=True)
    p.add_argument("--expected-shape-counts", type=json.loads, help="JSON list, e.g. [1,4]")
    p.add_argument("--expected-gaps", type=json.loads, help="JSON list, e.g. [[45,90],[0,18]]")
    a = p.parse_args()
    print(json.dumps(verify(a.input, json.loads(a.plan.read_text(encoding="utf-8")), a.expected_shape_counts, a.expected_gaps), indent=2))
