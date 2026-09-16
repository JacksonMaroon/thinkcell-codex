"""Offline high-confidence legend translation plan and physical grader.

This wrapper keeps the tested v8 metric-anchor preparation and adds the missing
physical gate: the union of the tagged PowerPoint legend rectangles must move
by the requested translation from the authoritative source baseline.  No
Office or native calls are made here.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
_CORE_DIR = HERE / "scripts" / "legend_controls"
if not (_CORE_DIR / "portable_metric.py").is_file():
    # Installed layout: this entrypoint and portable_metric.py are colocated.
    _CORE_DIR = HERE
sys.path.insert(0, str(_CORE_DIR))
import portable_metric as metric  # noqa: E402


PHYSICAL_TOLERANCE_PT = 0.05


def file_sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def _guard_plan_paths(*paths: Path, fresh: tuple[Path, ...] = ()) -> tuple[Path, ...]:
    """Require resolved, pairwise-distinct inputs and fresh generated outputs."""
    resolved = tuple(Path(p).expanduser().resolve() for p in paths)
    keys = [str(p).casefold() for p in resolved]
    if len(set(keys)) != len(keys):
        raise ValueError("input, data, output and plan paths must be pairwise distinct")
    for path in fresh:
        if path.exists():
            raise FileExistsError(f"refusing to overwrite existing generated artifact: {path}")
    return resolved


def physical_bounds(path: Path, target: dict, legend_selector: dict):
    inspected = metric.inspect(path, target, legend_selector)
    bounds = inspected["physical"].get("legend_bounds")
    metric.need(bounds is not None, "Tagged physical legend union is unavailable")
    metric.need(not inspected["native_group_nesting"],
                     "Physical union is ambiguous under native group nesting")
    return inspected, [float(x) for x in bounds]


def make_plan(path: Path, translation: dict, baseline_data: Path, changed_data: Path,
              selector: dict | None = None, legend_selector: dict | None = None):
    dx = float(translation["x"])
    dy = float(translation["y"])
    core = metric.make_plan(path, selector, legend_selector, {"x": dx, "y": dy})
    inspected, baseline_physical = physical_bounds(
        path, core["target"], core["legend_selector"])
    expected_physical = [baseline_physical[0] + dx,
                         baseline_physical[1] + dy,
                         baseline_physical[2] + dx,
                         baseline_physical[3] + dy]
    core["physical_gate"] = {
        "baseline_bounds": baseline_physical,
        "expected_bounds": expected_physical,
        "tolerance_pt": PHYSICAL_TOLERANCE_PT,
        "source_method": "tagged_physical_legend_union",
    }
    core["data_requests"] = {
        "same_data": {"path": str(baseline_data), "sha256": file_sha(baseline_data)},
        "changed_data": {"path": str(changed_data), "sha256": file_sha(changed_data)},
    }
    core["baseline_model_version"] = inspected["model_version"]
    return core


def _core(plan: dict) -> dict:
    """Return the exact metric plan expected by the existing preparation gate."""
    return {k: copy.deepcopy(v) for k, v in plan.items()
            if k not in {"physical_gate", "data_requests", "baseline_model_version"}}


def _stable_target(plan: dict) -> dict:
    """Use the semantic chart tag for native readback; PowerPoint may renumber shapes."""
    target = copy.deepcopy(plan["target"])
    target["shape_id"] = None
    return target


def prepare(source: Path, output: Path, plan: dict):
    metric.need(not output.exists() and source.resolve() != output.resolve(),
                     "Distinct new output required")
    metric.need(file_sha(source) == plan["source_sha256"],
                     "Source hash changed; inspect again")
    expected = make_plan(
        source,
        plan["operation"]["translation"],
        Path(plan["data_requests"]["same_data"]["path"]),
        Path(plan["data_requests"]["changed_data"]["path"]),
        selector=plan["target"],
        legend_selector=plan["legend_selector"],
    )
    metric.need(plan == expected, "Stale or modified high legend plan")
    return metric.prepare(source, output, _core(plan))


def _close(got, want, tol):
    return len(got) == len(want) and all(
        math.isclose(float(a), float(b), abs_tol=tol) for a, b in zip(got, want)
    )


def grade(path: Path, plan: dict):
    readback_plan = _core(plan)
    readback_plan["target"] = _stable_target(plan)
    core_result = metric.grade(path, readback_plan)
    inspected, actual_physical = physical_bounds(
        path, _stable_target(plan), plan["legend_selector"])
    expected_physical = plan["physical_gate"]["expected_bounds"]
    tol = float(plan["physical_gate"]["tolerance_pt"])
    delta = [actual_physical[0] - plan["physical_gate"]["baseline_bounds"][0],
             actual_physical[1] - plan["physical_gate"]["baseline_bounds"][1],
             actual_physical[2] - plan["physical_gate"]["baseline_bounds"][2],
             actual_physical[3] - plan["physical_gate"]["baseline_bounds"][3]]
    physical_ok = _close(actual_physical, expected_physical, tol)
    structure_ok = bool(inspected["consistency"]["chart"] and
                        inspected["consistency"]["legend"] and
                        inspected["same_carrier"] and
                        not inspected["native_group_nesting"])
    return {
        "status": "PASS" if core_result["status"] == "PASS" and physical_ok and structure_ok else "FAIL",
        "model_text": core_result,
        "physical_bounds": actual_physical,
        "expected_physical_bounds": expected_physical,
        "physical_delta": delta,
        "physical_tolerance_pt": tol,
        "physical_exact": physical_ok,
        "structure": {
            "consistency": inspected["consistency"],
            "same_carrier": inspected["same_carrier"],
            "native_group_nesting": inspected["native_group_nesting"],
        },
    }


def grade_repeat(same: Path, changed: Path, plan: dict):
    same_result = grade(same, plan)
    changed_result = grade(changed, plan)
    return {
        "status": "PASS" if same_result["status"] == changed_result["status"] == "PASS" else "FAIL",
        "same_data": same_result,
        "changed_data": changed_result,
        "geometry_retained": (
            same_result["physical_bounds"] == changed_result["physical_bounds"] and
            same_result["model_text"]["model_bounds"] == changed_result["model_text"]["model_bounds"] and
            same_result["model_text"]["text_bounds"] == changed_result["model_text"]["text_bounds"]
        ),
    }


def grade_relative(baseline: Path, translated: Path, plan: dict):
    """Grade translation against an independently regenerated native baseline."""
    _, baseline_physical = physical_bounds(
        baseline, _stable_target(plan), plan["legend_selector"])
    relative_plan = copy.deepcopy(plan)
    dx = float(plan["operation"]["translation"]["x"])
    dy = float(plan["operation"]["translation"]["y"])
    relative_plan["physical_gate"]["baseline_bounds"] = baseline_physical
    relative_plan["physical_gate"]["expected_bounds"] = [
        baseline_physical[0] + dx, baseline_physical[1] + dy,
        baseline_physical[2] + dx, baseline_physical[3] + dy,
    ]
    result = grade(translated, relative_plan)
    result["relative_baseline_physical_bounds"] = baseline_physical
    result["relative_expected_delta"] = [dx, dy, dx, dy]
    return result


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("plan")
    p.add_argument("input", type=Path)
    p.add_argument("output", type=Path)
    p.add_argument("--translation", required=True)
    p.add_argument("--same-data", type=Path, required=True)
    p.add_argument("--changed-data", type=Path, required=True)
    p.add_argument("--selector-json", type=Path)
    p.add_argument("--legend-selector-json", type=Path)
    p.add_argument("--plan", type=Path, required=True)
    g = sub.add_parser("grade")
    g.add_argument("input", type=Path)
    g.add_argument("--plan", type=Path, required=True)
    r = sub.add_parser("grade-repeat")
    r.add_argument("same", type=Path)
    r.add_argument("changed", type=Path)
    r.add_argument("--plan", type=Path, required=True)
    q = sub.add_parser("grade-relative")
    q.add_argument("baseline", type=Path)
    q.add_argument("translated", type=Path)
    q.add_argument("--plan", type=Path, required=True)
    args = ap.parse_args()
    if args.cmd == "plan":
        checked = _guard_plan_paths(
            args.input, args.output, args.same_data, args.changed_data, args.plan,
            *(p for p in (args.selector_json, args.legend_selector_json) if p is not None),
            fresh=(Path(args.output).expanduser().resolve(), Path(args.plan).expanduser().resolve()))
        plan_input, plan_output, plan_same, plan_changed, plan_file = checked[:5]
        selector = json.loads(Path(args.selector_json).read_text(encoding="utf-8")) if args.selector_json else None
        legend_selector = json.loads(Path(args.legend_selector_json).read_text(encoding="utf-8")) if args.legend_selector_json else None
        plan = make_plan(plan_input, json.loads(args.translation), plan_same, plan_changed,
                         selector, legend_selector)
        plan_file.write_text(json.dumps(plan, indent=2) + "\n", encoding="utf-8")
        result = prepare(plan_input, plan_output, plan)
        print(json.dumps({"status": result["status"], "plan": str(plan_file),
                          "output": str(plan_output), "output_sha256": file_sha(plan_output)}, indent=2))
    elif args.cmd == "grade":
        result = grade(args.input, json.loads(args.plan.read_text(encoding="utf-8")))
        print(json.dumps(result, indent=2))
        raise SystemExit(0 if result["status"] == "PASS" else 1)
    elif args.cmd == "grade-repeat":
        plan = json.loads(args.plan.read_text(encoding="utf-8"))
        result = grade_repeat(args.same, args.changed, plan)
        print(json.dumps(result, indent=2))
        raise SystemExit(0 if result["status"] == "PASS" else 1)
    else:
        plan = json.loads(args.plan.read_text(encoding="utf-8"))
        result = grade_relative(args.baseline, args.translated, plan)
        print(json.dumps(result, indent=2))
        raise SystemExit(0 if result["status"] == "PASS" else 1)


if __name__ == "__main__":
    main()
