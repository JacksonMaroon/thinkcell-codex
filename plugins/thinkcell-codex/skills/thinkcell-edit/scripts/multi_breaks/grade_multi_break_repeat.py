"""Offline grade for a completed native multi-break changed-data repeat."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import portable_multi_break_adapter as adapter
import verify_two_breaks


def need(ok: bool, message: str):
    if not ok:
        raise RuntimeError(message)


def grade(input_path: Path, plan_path: Path, report_path: Path | None = None, render_path: Path | None = None):
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    data = plan["targets"][0]["data"]
    expected_counts = data["expected_generated_shape_counts"]
    expected_gaps = data["expected_generated_gaps"]
    portable = adapter.preflight(input_path.resolve(), plan, "native")
    dedicated = verify_two_breaks.verify(input_path.resolve(), plan, expected_counts, expected_gaps)
    report = None
    if report_path is not None:
        need(report_path.exists(), "native report is missing")
        report = json.loads(report_path.read_text(encoding="utf-8-sig"))
        need(report.get("native_reopen_pass") is True, "native reopen flag is not true")
        need(report.get("source_unchanged") is True, "source unchanged flag is not true")
        need(report.get("other_presentations_unchanged") is True, "scope flag is not true")
    render = None
    if render_path is not None:
        need(render_path.exists() and render_path.stat().st_size > 0, "render is missing or empty")
        render = {"path": str(render_path.resolve()), "bytes": render_path.stat().st_size}
    return {"status": "MULTI_BREAK_REPEAT_GRADE_PASS", "portable": portable, "dedicated": dedicated, "native_scope": report, "render": render}


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--render", type=Path)
    args = parser.parse_args()
    print(json.dumps(grade(args.input, args.plan, args.report, args.render), indent=2))
