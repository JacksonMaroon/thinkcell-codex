"""Lane wrapper for bounded axis-break insertion controls.

Imports the staged runner unchanged and replaces only its data/topology guards.
All Office work still runs through the staged runner's serialized operation.
"""
from __future__ import annotations

import importlib.util
import math
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
_configured_scripts = os.environ.get("THINKCELL_PLUGIN_SCRIPTS")
def _script_candidates():
    if _configured_scripts:
        yield Path(_configured_scripts).expanduser().resolve()
        return
    # A copied bundle can sit beside the plugin's scripts or under a workspace
    # lane. Prefer an installed sibling/parent, then a staged development tree.
    roots = (HERE, *HERE.parents)
    for root in roots:
        yield root / "scripts"
        yield root / "skills" / "thinkcell-edit" / "scripts"
    # Workspace staged trees are a final development fallback.
    for root in roots:
        yield root / "staged-plugin" / "skills" / "thinkcell-edit" / "scripts"


_candidates = list(_script_candidates())
BASE = next((p for p in _candidates if (p / "run_native_axis_break_insert.py").is_file()), None)
if BASE is None:
    raise RuntimeError("think-cell scripts not found; set THINKCELL_PLUGIN_SCRIPTS to the target scripts directory; searched: " + "; ".join(str(p) for p in _candidates))
spec = importlib.util.spec_from_file_location("staged_axis_break_runner", BASE / "run_native_axis_break_insert.py")
if spec is None or spec.loader is None:
    raise RuntimeError("unable to load staged axis-break runner")
runner = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = runner
spec.loader.exec_module(runner)

# The staged preparer pins the original vertical donor's length-12 interval
# and no-tickmark shape as data-independent safeguards. Keep its structural
# checks, but allow a target's own interval cardinality and existing tickmark
# collection for this lane's horizontal control.
_prepare_need = runner.prepare.need


def _prepare_need_variant(condition, message):
    if condition or message in {"unsupported ordinary interval topology", "unsupported target tickmark topology"}:
        return
    return _prepare_need(condition, message)


runner.prepare.need = _prepare_need_variant


def need(condition: bool, message: str):
    if not condition:
        raise RuntimeError(message)


def target_any(path: Path, plan: dict):
    docs, charts, _ = runner.inventory(path.read_bytes())
    need(len(docs) == len(charts) == 1, "requires exactly one logical slide, carrier, and native chart")
    wanted = plan["targets"][0]
    choices = [c for c in charts if c["exact"] and c["owner_name"] == wanted["name"]]
    need(len(choices) == 1, "target automation name did not resolve uniquely")
    chart = choices[0]
    need(chart["owner"].tag == "CSequenceChartSE", "unsupported chart subtype")
    ect = chart["owner"].find("m_ect")
    need(ect is not None and ect.get("val") in {"0", "1"}, "unsupported sequence subtype")
    need(chart["frames"][0]["shape_tag"] == wanted["selector"]["shape_tag"], "target shape tag changed")
    runner.link_contract(chart)
    cache_part = chart["frames"][0]["native_chart_part"]
    need(cache_part, "target has no native chart cache")
    with runner.zipfile.ZipFile(path) as z:
        need(len(runner.logical_slides(z)) == 1, "requires exactly one logical slide")
        need(cache_part in z.namelist(), "native chart cache is missing")
        cache = runner.chart_details(z.read(cache_part))
    subtype = cache.get("subtypes", [])
    need(len(subtype) == 1 and subtype[0].get("type") == "barChart", "requires ordinary bar chart cache")
    need(subtype[0].get("bar_direction") in {"col", "bar"}, "unsupported chart orientation")
    need(subtype[0].get("grouping") in {"clustered", "stacked"}, "unsupported chart grouping")
    return chart, wanted


def profile_any(plan: dict):
    need(set(plan) == {"targets"} and len(plan["targets"]) == 1, "one target plan is required")
    data = plan["targets"][0]["data"]
    matrix, expected = data["matrix"], data["expected_model"]
    need(isinstance(matrix, list) and len(matrix) >= 2, "matrix needs header and at least one series")
    width = len(matrix[0])
    need(width >= 2 and all(isinstance(row, list) and len(row) == width for row in matrix), "matrix is not rectangular")
    series_rows = matrix[2:] if len(matrix) >= 3 and all(value in (None, "") for value in matrix[1]) else matrix[1:]
    categories = expected["categories"]
    names = expected["series_names"]
    values = expected["series_values"]
    need(len(categories) == width - 1 and len(names) == len(series_rows), "model dimensions disagree with matrix")
    need(len(values) == len(names) and all(len(row) == len(categories) for row in values), "value dimensions disagree with labels")
    need(matrix[0][1:] == categories and [row[0] for row in series_rows] == names, "matrix labels disagree with expected model")
    for i, row in enumerate(series_rows):
        need(all(runner.equal(a, b) for a, b in zip(row[1:], values[i])), "matrix values disagree with expected model")
    for row in values:
        need(all(v is None or (isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) and v >= 0) for v in row), "values must be finite nonnegative numbers or blank")
    totals = [sum((row[i] or 0) for row in values) for i in range(len(categories))]
    need(all(runner.equal(a, b) for a, b in zip(expected["category_extents"], totals)), "category extents disagree with values")
    expected_fraction = data.get("expected_fraction")
    if expected_fraction is not None:
        need(isinstance(expected_fraction, (int, float)) and not isinstance(expected_fraction, bool) and math.isfinite(expected_fraction) and 0 < expected_fraction < 1, "expected fraction must be between zero and one")
    return data


def verify_any(path: Path, plan: dict, expected_shape_count: int):
    data = profile_any(plan)
    chart, _ = target_any(path, plan)
    actual = runner.model_of(chart)
    expected = data["expected_model"]
    for key in ("categories", "series_names", "series_values", "category_extents"):
        need(runner.equal(actual.get(key), expected[key]), "exact model mismatch: " + key)
    sheets = runner.read_sheet(path, 1, chart["owner_name"])["sheets"]
    need(len(sheets) == 1, "one datasheet required")
    cells = {(x["row"], x["column"]): x["value"] for x in sheets[0]["nonempty_cells"]}
    wanted_cells = {(r + 1, c + 1): value for r, row in enumerate(data["matrix"]) for c, value in enumerate(row) if value not in (None, "")}
    need(cells == wanted_cells, "exact datasource cells mismatch")
    ids = chart["doc"]["ids"]
    axis_ref = chart["owner"].find("m_daxisPrimaryValue")
    need(axis_ref is not None and axis_ref.get("idref") in ids, "primary value axis missing")
    axis = ids[axis_ref.get("idref")]
    refs = axis.findall("m_cdaxisbreak/elem")
    need(len(refs) == 1 and refs[0].get("idref") in ids, "one native axis break is required")
    br = ids[refs[0].get("idref")]
    fraction = br.find("m_fFraction")
    need(br.tag == "CDataAxisBreak" and br.find("m_bUser") is not None and br.find("m_bUser").get("val") == "1" and fraction is not None, "break is not user-owned native state")
    expected_fraction = data.get("expected_fraction")
    if expected_fraction is not None:
        need(math.isclose(float(fraction.get("val")), float(expected_fraction), rel_tol=0, abs_tol=1e-12), "unexpected break fraction")
    shapes = [ids[x.get("idref")] for x in br.findall("m_cpptbreakshp/elem")]
    need(shapes and all(x.tag == "CPPTBreakShape" for x in shapes), "break has no generated native shape")
    names = [x.find(field + "/m_bstrShapeName").text for x in shapes for field in runner.FIELDS]
    need(len(names) == 3 * len(shapes) and len(set(names)) == len(names) and all(names), "incomplete linked break tags")
    runner.physical_tags(path.read_bytes(), names)
    integrity = runner.specialized_integrity(path)
    if expected_shape_count == 2:
        return {"status": "PREPARED_SEED_NATIVE_REGENERATION_REQUIRED", "exact_model": True, "exact_datasheet": True, "fraction": float(fraction.get("val")), "native_break_shape_count": len(shapes), "physical_tag_count": len(names), "integrity": integrity}
    intervals = axis.findall("m_vecintvlOrdinal")
    need(intervals, "axis interval vector missing")
    scalar_values = [v for row in expected["series_values"] for v in row if v is not None]
    gap_elem = intervals[0].find("elem")
    need(gap_elem is not None and gap_elem.find("begin") is not None and gap_elem.find("end") is not None, "generated gap interval is missing")
    low, high = float(gap_elem.find("begin").get("val")), float(gap_elem.find("end").get("val"))
    need(math.isfinite(low) and math.isfinite(high) and low < high, "generated gap interval is not positive")
    # For the positive-value column/bar controls, a baseline-zero bar crosses
    # the gap exactly when its endpoint is in the gap. Keep this explicit
    # until a negative/stacked topology has its own native proof.
    crossed = [value for value in scalar_values if 0 <= high and value > low and value <= high]
    need(crossed, "generated gap does not cross requested positive data")
    expected_shape_count = data.get("expected_shape_count")
    need(len(shapes) == (expected_shape_count if expected_shape_count is not None else len(crossed)), "generated shape count disagrees with expected crossing-bar count")
    expected_crossing_values = data.get("expected_crossing_values")
    if expected_crossing_values is not None:
        need(sorted(crossed) == sorted(expected_crossing_values), "generated gap crossing values disagree with explicit control")
    return {"exact_model": True, "exact_datasheet": True, "fraction": float(fraction.get("val")), "gap": [low, high], "crossing_values": crossed, "native_break_shape_count": len(shapes), "physical_tag_count": len(names), "axis_interval_versions": [{"reqver": x.get("reqver"), "endver": x.get("endver"), "length": x.get("length")} for x in intervals], "integrity": integrity}


runner.target = target_any
runner.profile = profile_any
runner.verify = verify_any


if __name__ == "__main__":
    runner.main = None
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=("insert", "repeat"), required=True)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--expected-source-sha256", required=True)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--pre-plan", type=Path)
    parser.add_argument("--break-donor", type=Path)
    parser.add_argument("--expected-donor-sha256")
    parser.add_argument("--ppttc", type=Path)
    parser.add_argument("--execute", action="store_true")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    runner.need(args.mode == "repeat" or (args.break_donor and args.expected_donor_sha256), "insert mode requires donor")
    runner.need(not args.execute or args.ppttc, "--execute requires --ppttc")
    print(runner.json.dumps(runner.run(args), indent=2))
