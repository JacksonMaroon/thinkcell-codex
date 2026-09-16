"""One-command scalar endpoint-expansion production candidate.

This wrapper is deliberately limited to the native-proven scalar route. It
prepares a semantic endpoint insertion, runs official JSON regeneration and
native save/reopen, then requires the independent model/field/datasheet/source
gate before writing one final report. The caller supplies the existing
think-cell skill directory and ppttc path; no workspace or cache path is
embedded here.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import subprocess
import sys
from pathlib import Path


def need(ok: bool, msg: str) -> None:
    if not ok:
        raise RuntimeError(msg)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def call(command: list[str], log: Path) -> None:
    with log.open("w", encoding="utf-8") as handle:
        result = subprocess.run(command, stdout=handle, stderr=subprocess.STDOUT, text=True)
    need(result.returncode == 0, f"child command failed ({result.returncode}); see {log}")


def expected_graphs(prep: dict) -> int:
    chart = prep.get("chart") or {}
    count = chart.get("cagr_count")
    need(isinstance(count, int) and count >= 1, "promotable route requires an existing native CAGR donor")
    return count + 1


def run(a: argparse.Namespace) -> dict:
    # The promotable bundle carries its core scripts beside this wrapper.  Only
    # the caller-supplied skill directory supplies external implementation
    # files; no plugin/cache path is inferred here.
    root = Path(__file__).resolve().parent
    adapter = root / "parameterized_cagr.py"
    runner = root / "run_parameterized_control.py"
    verifier = root / "verify_cagr_result.py"
    for path in (adapter, runner, verifier, a.input, a.data, a.skill_dir, a.ppttc):
        need(path.exists(), f"missing required path: {path}")
    work = a.workdir.resolve(); work.mkdir(parents=True, exist_ok=True)
    candidate = work / "candidate.pptx"; prep_path = work / "prepare.json"
    need(not candidate.exists() and not prep_path.exists(), "promotable output names already exist")
    prep_cmd = [sys.executable, str(adapter), "--input", str(a.input.resolve()), "--expected-sha256", a.expected_sha256, "--output", str(candidate), "--report", str(prep_path), "--workdir", str(work), "--skill-dir", str(a.skill_dir.resolve()), "--source-index", str(a.source_index), "--sink-index", str(a.sink_index)]
    if a.chart_name:
        prep_cmd.extend(["--chart-name", a.chart_name])
    call(prep_cmd, work / "prepare.log")
    prep = json.loads(prep_path.read_text(encoding="utf-8-sig")); need(prep.get("candidate_sha256")==sha(candidate), "prepared candidate hash does not match report"); need(bool(prep.get("source_unchanged")), "source integrity gate failed during preparation")
    native_stem = a.stem or "native"
    run_cmd = [sys.executable, str(runner), "--input", str(candidate), "--data", str(a.data.resolve()), "--skill-dir", str(a.skill_dir.resolve()), "--ppttc", str(a.ppttc.resolve()), "--workdir", str(work), "--stem", native_stem, "--source-index", str(a.source_index), "--sink-index", str(a.sink_index), "--prepared-report", str(prep_path)]
    call(run_cmd, work / "native-run.log")
    run_report = work / f"{native_stem}-run.json"; need(run_report.is_file(), "runner report missing")
    run_info = json.loads(run_report.read_text(encoding="utf-8-sig")); native = Path(run_info["native"]); need(native.is_file(), "native reopened artifact missing")
    readback = work / "independent-readback.json"
    verify_cmd = [sys.executable, str(verifier), "--input", str(native), "--prep-report", str(prep_path), "--data", str(a.data.resolve()), "--skill-dir", str(a.skill_dir.resolve()), "--report", str(readback), "--source-index", str(a.source_index), "--sink-index", str(a.sink_index), "--expected-graphs", str(expected_graphs(prep))]
    call(verify_cmd, work / "verify.log")
    gate = json.loads(readback.read_text(encoding="utf-8-sig")); need(gate.get("status")=="NATIVE_CAGR_INDEPENDENT_READBACK_PASS", "independent native gate did not pass")
    report = {"status": "PROMOTABLE_SCALAR_ENDPOINT_EXPANSION_NATIVE_PASS", "input": str(a.input.resolve()), "input_sha256": a.expected_sha256.upper(), "candidate": str(candidate), "candidate_sha256": sha(candidate), "prepare_report": str(prep_path), "native_run_report": str(run_report), "native_reopened": str(native), "independent_readback": str(readback), "chart": prep.get("chart"), "endpoints": prep.get("endpoints"), "native_gate": {"status": gate.get("status"), "graph_count": gate.get("graph_count"), "actual_cagr": gate.get("actual_cagr"), "visible_cache": gate.get("visible_cache"), "datasheet_values": gate.get("datasheet_values")}, "source_unchanged": bool(prep.get("source_unchanged"))}
    final = work / "promotable-report.json"; final.write_text(json.dumps(report, indent=2), encoding="utf-8"); print(json.dumps(report, indent=2)); return report


if __name__ == "__main__":
    p=argparse.ArgumentParser(description=__doc__); p.add_argument("--input",type=Path,required=True); p.add_argument("--expected-sha256",required=True); p.add_argument("--data",type=Path,required=True); p.add_argument("--skill-dir",type=Path,required=True); p.add_argument("--ppttc",type=Path,required=True); p.add_argument("--workdir",type=Path,required=True); p.add_argument("--source-index",type=int,required=True); p.add_argument("--sink-index",type=int,required=True); p.add_argument("--chart-name"); p.add_argument("--stem",default="native"); run(p.parse_args())
