"""Run one official JSON/native-reopen CAGR control from a prepared copy.

The script owns the repeatable data contract, optional semantic preparation
binding, official regeneration and scoped native save/reopen.  It is safe to
invoke only after the root agent has acquired the Office lease.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import subprocess
import sys
from pathlib import Path


def need(ok: bool, msg: str) -> None:
    if not ok:
        raise RuntimeError(msg)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def official_30_360(start: dt.date, end: dt.date) -> float:
    """Return elapsed years using think-cell's official 30/360 convention."""
    days = 360 * (end.year - start.year) + 30 * (end.month - start.month) + (min(end.day, 30) - min(start.day, 30))
    need(days > 0, "dates must have a positive 30/360 elapsed period")
    return days / 360.0


def read_plan(path: Path) -> dict:
    plan = json.loads(path.read_text(encoding="utf-8-sig"))
    need(set(plan) == {"name", "dates", "values"}, "data JSON requires exactly name, dates and values")
    need(isinstance(plan["name"], str) and plan["name"].strip(), "datasheet name is required")
    dates = [dt.date.fromisoformat(x) for x in plan["dates"]]
    values = [float(x) for x in plan["values"]]
    need(len(dates) == len(values) and len(dates) >= 2, "dates and values must have equal length >= 2")
    need(all(b > a for a, b in zip(dates, dates[1:])), "dates must be strictly increasing")
    need(all(math.isfinite(x) and x > 0 for x in values), "values must be finite and positive")
    return {"name": plan["name"], "dates": dates, "values": values}


def make_job(template: Path, plan: dict, path: Path) -> None:
    table = [[None] + [{"date": x.isoformat()} for x in plan["dates"]], [None] * (len(plan["dates"]) + 1), [{"string": "Label"}] + [{"number": x} for x in plan["values"]]]
    path.write_text(json.dumps([{"template": str(template), "data": [{"name": plan["name"], "table": table}]}], indent=2), encoding="utf-8")


def validate_prepared(input_path: Path, plan: dict, prepared_report: Path | None,
                      expected_input_sha: str | None, source_index: int,
                      sink_index: int) -> dict | None:
    """Validate the semantic preparation contract before any native write.

    The report is optional for compatibility with older prepared controls, but
    when supplied it binds the candidate bytes, chart identity and endpoint
    selection.  This keeps semantic discovery from becoming a weaker check than
    the old fixed carrier/model assertions.
    """
    actual = sha(input_path)
    if expected_input_sha:
        need(actual == expected_input_sha.upper(), "input SHA-256 mismatch")
    if prepared_report is None:
        return None
    report = json.loads(prepared_report.read_text(encoding="utf-8-sig"))
    need(report.get("status", "").startswith("PREPARED_"), "prepared report is not a preparation result")
    bound = report.get("candidate_sha256")
    need(isinstance(bound, str) and bound.upper() == actual, "input does not match prepared candidate SHA-256")
    chart = report.get("chart") or report.get("target_chart")
    need(isinstance(chart, dict) and chart.get("chart_name"), "prepared report has no semantic chart name")
    need(plan["name"] == chart["chart_name"], "data name does not match prepared semantic chart")
    endpoints = report.get("endpoints") or {}
    need(int(endpoints.get("source_index", -1)) == source_index and int(endpoints.get("sink_index", -1)) == sink_index,
         "runner endpoints do not match prepared semantic endpoints")
    need(bool(report.get("source_unchanged", report.get("target_unchanged", False))),
         "prepared source/target integrity flag is false")
    return {"report": str(prepared_report), "candidate_sha256": actual,
            "chart_name": chart["chart_name"], "carrier": chart.get("carrier"),
            "series_count": chart.get("series_count"), "category_count": chart.get("category_count")}


def run(a: argparse.Namespace) -> dict:
    scripts = a.skill_dir.resolve() / "scripts"; impl = scripts / "thinkcell_no_click" / "implementation"
    sys.path[:0] = [str(scripts), str(impl)]
    from office_operation_lock import OfficeOperationLock, run_locked_subprocess  # type: ignore
    from runtime import powershell, powershell_env  # type: ignore
    plan = read_plan(a.data); input_path = a.input.resolve()
    prepared = validate_prepared(input_path, plan, a.prepared_report.resolve() if a.prepared_report else None,
                                 a.expected_input_sha256, a.source_index, a.sink_index)
    work = a.workdir.resolve(); work.mkdir(parents=True, exist_ok=True)
    stage = work / (a.stem + "-stage"); need(not stage.exists(), "stage already exists"); stage.mkdir()
    job, generated = stage / "update.ppttc", stage / "generated.pptx"; make_job(input_path, plan, job)
    ppttc = a.ppttc.resolve(); need(ppttc.is_file(), "ppttc.exe is missing")
    with OfficeOperationLock("parameterized-cagr-control", metadata={"input": str(a.input), "data": str(a.data)}):
        with (stage / "ppttc.stdout.txt").open("w", encoding="utf-8") as stdout, (stage / "ppttc.stderr.txt").open("w", encoding="utf-8") as stderr:
            result = run_locked_subprocess([str(ppttc), str(job), "-o", str(generated)], operation="parameterized-cagr-json", timeout_seconds=a.timeout, stdout=stdout, stderr=stderr, creationflags=subprocess.CREATE_NO_WINDOW)
        need(result.returncode == 0 and generated.is_file(), "official JSON regeneration failed")
        native, native_report, render = stage / "native-reopened.pptx", stage / "native-report.json", stage / "native.png"
        command = [powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(impl / "native_verify_scoped.ps1"), "-InputFile", str(generated), "-OutputFile", str(native), "-ReportFile", str(native_report), "-RenderFile", str(render)]
        with (stage / "native.stdout.txt").open("w", encoding="utf-8") as stdout, (stage / "native.stderr.txt").open("w", encoding="utf-8") as stderr:
            result = run_locked_subprocess(command, operation="parameterized-cagr-native-reopen", timeout_seconds=a.timeout, stdout=stdout, stderr=stderr, env=powershell_env(), creationflags=subprocess.CREATE_NO_WINDOW)
        need(result.returncode == 0 and native.is_file() and native_report.is_file(), "native save/reopen failed")
    start, end = plan["dates"][a.source_index], plan["dates"][a.sink_index]
    years = official_30_360(start, end); expected = (plan["values"][a.sink_index] / plan["values"][a.source_index]) ** (1 / years) - 1
    report = {"status": "OFFICIAL_REGEN_NATIVE_REOPEN_COMPLETE", "input": str(input_path), "input_sha256": sha(input_path), "data": str(a.data), "generated": str(generated), "native": str(native), "native_report": str(native_report), "preview": str(render), "endpoints": {"source_index": a.source_index, "sink_index": a.sink_index, "start_date": start.isoformat(), "end_date": end.isoformat(), "official30_360_years": years, "expected_cagr": expected}, "prepared_binding": prepared}
    (work / (a.stem + "-run.json")).write_text(json.dumps(report, indent=2), encoding="utf-8"); print(json.dumps(report, indent=2)); return report


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__); p.add_argument("--input", type=Path, required=True); p.add_argument("--data", type=Path, required=True); p.add_argument("--skill-dir", type=Path, required=True); p.add_argument("--ppttc", type=Path, required=True); p.add_argument("--workdir", type=Path, required=True); p.add_argument("--stem", required=True); p.add_argument("--source-index", type=int, required=True); p.add_argument("--sink-index", type=int, required=True); p.add_argument("--prepared-report", type=Path); p.add_argument("--expected-input-sha256"); p.add_argument("--timeout", type=int, default=240); run(p.parse_args())
