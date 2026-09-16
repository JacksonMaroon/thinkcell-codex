"""Fail-closed existing native axis-break position adapter.

This is deliberately not an insertion API. It uses the established guarded
single-field OLE writer, then supplies a hash-bound request for the installed
``broken_axis_update`` runner and verifies native fraction plus break geometry.
"""
from __future__ import annotations

import hashlib
import json
import math
import sys
from types import SimpleNamespace
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path[:0] = [str(HERE), str(HERE / "thinkcell_no_click" / "implementation")]
from prepare_thinkcell_name import need  # noqa: E402
from fraction_position_gate import snapshot, verify as verify_position  # noqa: E402

need((HERE / "axis_break.py").is_file() and (HERE / "broken_axis_update.py").is_file(), "Package local break helpers")
import axis_break as _patcher  # noqa: E402
import broken_axis_update as _runner  # noqa: E402


def _sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest().upper()


def make_plan(source, selector, fraction):
    need(isinstance(fraction, (int, float)) and not isinstance(fraction, bool) and math.isfinite(fraction),
         "Fraction must be a finite number")
    need(.05 <= fraction <= .95, "Fraction must be between .05 and .95")
    base = _patcher.make_plan(source, selector, fraction)
    profile = json.loads(json.dumps(snapshot(source, selector)))
    need(len(profile["shapes"]) == 2 and len({s["id"] for s in profile["shapes"]}) == 2,
         "Require one break bound to two distinct native shapes")
    return {"schema_version": 1, "kind": "existing_native_axis_break_position", "source_sha256": _sha(source),
            "selector": dict(selector), "requested_fraction": float(fraction), "patch_plan": base,
            "profile": profile,
            "gates": ["source_hash", "one_existing_native_break", "two_break_shapes", "only_fraction_stream_patch",
                      "official_json", "native_reopen", "fraction_readback", "break_geometry_moved"]}


def prepare(source, output, plan):
    source, output = Path(source), Path(output)
    expected = make_plan(source, plan["selector"], plan["requested_fraction"])
    need(plan == expected, "Stale or modified position plan")
    result = _patcher.prepare(source, output, plan["patch_plan"])
    need(_sha(source) == plan["source_sha256"], "Source changed during preparation")
    return {"status": "POSITION_PREPARED_FOR_OFFICIAL_JSON", "prepared": str(output),
            "prepared_sha256": _sha(output), "only_fraction_stream_patch": result["only_existing_native_axis_break_changed"],
            "profile": plan["profile"]}


def execution_request(prepared, data_plan, output, report):
    """Return the exact installed-runner arguments; caller owns Office lease."""
    return {"runner": str(HERE / "broken_axis_update.py"), "input": str(Path(prepared)),
            "expected_sha256": _sha(prepared), "plan": str(Path(data_plan)), "output": str(Path(output)),
            "report": str(Path(report)), "prepare": True, "execute": True}


def verify(before, native_output, plan):
    noop = math.isclose(plan["profile"]["fraction"], plan["requested_fraction"], abs_tol=1e-8)
    result = verify_position(before, native_output, plan["selector"], plan["requested_fraction"], allow_unchanged=noop)
    if noop:
        need(result["moved_rectangles"] == 0, "No-op fraction changed break geometry")
    else:
        need(result["moved_rectangles"] == 6, "Expected every low/high/body break rectangle to move")
    result["source_sha256"] = plan["source_sha256"]
    return result


def execute(source, data_plan, output, report, plan):
    """Stage both native operations and publish only after geometry proof."""
    source, output, report = Path(source), Path(output), Path(report)
    data_plan = Path(data_plan)
    paths = [source.resolve(), data_plan.resolve(), output.resolve(), report.resolve()]
    need(len(set(paths)) == 4 and not output.exists() and not report.exists(), "Input, data plan, output, and report must be distinct")
    original_hash = _sha(source)
    patched = output.with_name(output.stem + "_fraction_patch.pptx")
    native = output.with_name(output.stem + "_native_pending.pptx")
    prepare(source, patched, plan)
    args = SimpleNamespace(input=patched, expected_sha256=_sha(patched), plan=Path(data_plan), output=native,
                           report=report, ppttc=None, prepare=True, execute=True)
    result = _runner.run(args)
    gate = verify(source, native, plan)
    need(_sha(source) == original_hash == plan["source_sha256"], "Original source changed")
    with output.open("xb") as target: target.write(native.read_bytes())
    result["position_gate"] = gate; result["output"] = str(output); result["output_sha256"] = _sha(output); result["published_after_position_gate"] = True
    with report.open('x', encoding='utf-8') as written:
        json.dump(result, written, indent=2)
    return result


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(); parser.add_argument("--input", type=Path, required=True); parser.add_argument("--output", type=Path, required=True); parser.add_argument("--fraction", type=float, required=True); parser.add_argument("--shape-tag", required=True); parser.add_argument("--plan", type=Path, required=True); parser.add_argument("--data-plan", type=Path); parser.add_argument("--report", type=Path); parser.add_argument("--prepare", action="store_true"); parser.add_argument("--execute", action="store_true")
    args = parser.parse_args(); selector = {"shape_tag": args.shape_tag}
    need(not args.execute or args.data_plan is not None and args.report is not None, "Execute requires --data-plan and --report")
    paths = [p.resolve() for p in (args.input, args.output, args.plan, args.data_plan, args.report) if p is not None]
    need(len(paths) == len(set(paths)), "Input, output, plan, data plan and report must differ")
    plan = make_plan(args.input, selector, args.fraction)
    need(not args.plan.exists(), "Plan already exists")
    with args.plan.open('x', encoding='utf-8') as written:
        json.dump(plan, written, indent=2)
    answer = execute(args.input, args.data_plan, args.output, args.report, plan) if args.execute else (prepare(args.input, args.output, plan) if args.prepare else plan)
    print(json.dumps(answer))
