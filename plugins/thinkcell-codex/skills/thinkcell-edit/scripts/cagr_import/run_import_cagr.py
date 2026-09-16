"""Prepare and fully audit a semantic CAGR donor import package.

This wrapper performs offline preparation and package gates only. Native
regeneration/reopen remains a separate root-owned step using the candidate and
the standard scalar runner. Inputs and outputs are hash-bound and all output
names must be fresh.
"""
from __future__ import annotations

import argparse
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


def run_child(command: list[str], log: Path) -> None:
    with log.open("w", encoding="utf-8") as handle:
        result = subprocess.run(command, stdout=handle, stderr=subprocess.STDOUT, text=True)
    need(result.returncode == 0, f"child command failed ({result.returncode}); see {log}")


def audit(path: Path, output: Path) -> dict:
    command = [sys.executable, str(Path(__file__).with_name("audit_thinkcell_integrity.py")), str(path), "--strict-parity", "--output", str(output)]
    run_child(command, output.with_suffix(".log"))
    report = json.loads(output.read_text(encoding="utf-8-sig")); checks = report.get("checks") or {}
    need(bool(report.get("pass")) and all(checks.values()), "full relationship/package audit failed")
    return {"pass": True, "checks": checks, "presentation_sha256": report.get("presentation", {}).get("sha256")}


def run(a: argparse.Namespace) -> dict:
    bundle = Path(__file__).resolve().parent
    importer, allocator_test = bundle / "import_cagr_closure.py", bundle / "test_relationship_allocator.py"
    for path in (importer, allocator_test, a.target, a.donor, a.skill_dir):
        need(path.exists(), f"missing required path: {path}")
    target, donor = a.target.resolve(), a.donor.resolve()
    need(sha(target) == a.target_sha256.upper(), "target SHA-256 mismatch")
    need(sha(donor) == a.donor_sha256.upper(), "donor SHA-256 mismatch")
    work = a.workdir.resolve(); work.mkdir(parents=True, exist_ok=True)
    candidate, prep = work / "candidate.pptx", work / "prepare.json"
    final = work / "import-bundle-report.json"
    need(not candidate.exists() and not prep.exists() and not final.exists(), "bundle outputs already exist")
    run_child([sys.executable, str(allocator_test)], work / "allocator-test.log")
    command = [sys.executable, str(importer), "--target", str(target), "--donor", str(donor), "--target-sha256", a.target_sha256, "--output", str(candidate), "--report", str(prep), "--workdir", str(work), "--skill-dir", str(a.skill_dir.resolve()), "--target-name", a.target_name, "--donor-name", a.donor_name, "--source-index", str(a.source_index), "--sink-index", str(a.sink_index)]
    run_child(command, work / "prepare.log")
    prepared = json.loads(prep.read_text(encoding="utf-8-sig")); need(prepared.get("candidate_sha256") == sha(candidate), "prepared candidate SHA does not match report"); need(bool(prepared.get("target_unchanged")), "target integrity flag is false")
    candidate_audit = audit(candidate, work / "candidate-audit.json")
    report = {"status": "PREPARED_CAGR_IMPORT_FULL_AUDIT_PASS", "target": str(target), "target_sha256": a.target_sha256.upper(), "donor": str(donor), "donor_sha256": a.donor_sha256.upper(), "candidate": str(candidate), "candidate_sha256": sha(candidate), "prepare_report": str(prep), "allocator_test": "RELATIONSHIP_ALLOCATOR_PASS", "full_relationship_audit": candidate_audit, "native_execution": "not run", "prepared": prepared}
    final.write_text(json.dumps(report, indent=2), encoding="utf-8"); print(json.dumps(report, indent=2)); return report


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__); p.add_argument("--target", type=Path, required=True); p.add_argument("--donor", type=Path, required=True); p.add_argument("--target-sha256", required=True); p.add_argument("--donor-sha256", required=True); p.add_argument("--skill-dir", type=Path, required=True); p.add_argument("--target-name", required=True); p.add_argument("--donor-name", required=True); p.add_argument("--source-index", type=int, required=True); p.add_argument("--sink-index", type=int, required=True); p.add_argument("--workdir", type=Path, required=True); run(p.parse_args())
