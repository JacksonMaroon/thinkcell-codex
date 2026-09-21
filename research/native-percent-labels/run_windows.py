"""Reproduce scoped native reopen and independent JSON regeneration on copies.

Requires Windows PowerPoint with think-cell enabled. Never calls Quit.
Inspect emitted PNGs manually; successful automation is not visual approval.
"""
from pathlib import Path
import argparse
import hashlib
import json
import subprocess
import sys

LAB = Path(__file__).resolve().parent
SCRIPTS = LAB.parents[1] / "plugins/thinkcell-codex/skills/thinkcell-edit/scripts"
sys.path.insert(0, str(SCRIPTS))
from office_operation_lock import run_locked_subprocess
from runtime import powershell, powershell_env
import thinkcell
from prepare_thinkcell_name import inventory


def run(out):
    out.mkdir(parents=True, exist_ok=False)
    for stem in ["02-controlled-percent", "03-numerator-change",
                 "04-denominator-change", "05-reopened-saved"]:
        source = LAB / (stem + ".pptx")
        report = out / (stem + "-native.json")
        command = [powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned",
                   "-File", str(SCRIPTS / "thinkcell_no_click/implementation/native_verify_scoped.ps1"),
                   "-InputFile", str(source), "-OutputFile", str(out / (stem + "-windows.pptx")),
                   "-ReportFile", str(report), "-RenderFile", str(out / (stem + ".png"))]
        result = run_locked_subprocess(command, operation="percent-fixture-native-reopen",
                                       timeout_seconds=180, env=powershell_env(),
                                       creationflags=subprocess.CREATE_NO_WINDOW)
        if result.returncode:
            raise RuntimeError(f"Native verification failed: {stem}")
        data = json.loads(report.read_text(encoding="utf-8-sig"))
        if not all(data.get(k) for k in ["native_reopen_pass", "source_unchanged",
                                        "other_presentations_unchanged"]):
            raise RuntimeError(f"Native verification contract failed: {stem}")

    source = out / "02-controlled-percent-windows.pptx"
    for step in ["same", "numerator", "denominator"]:
        _, charts, _ = inventory(source.read_bytes())
        if len(charts) != 1:
            raise RuntimeError("Expected exactly one synthetic chart")
        chart = charts[0]
        request = thinkcell.baseline_request(source, chart)
        if step != "same":
            series, value = ("Series 3", 200) if step == "numerator" else ("Series 1", 125)
            next(row for row in request["matrix"] if row[0] == series)[1] = value
            index = request["expected_model"]["series_names"].index(series)
            request["expected_model"]["series_values"][index][0] = value
        request_path = out / ("json-" + step + "-request.json")
        request_path.write_text(json.dumps(request, indent=2), encoding="utf-8")
        destination = out / ("json-" + step + ".pptx")
        subprocess.run([sys.executable, str(SCRIPTS / "thinkcell.py"), "update",
                        "--input", str(source), "--expected-sha256",
                        hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                        "--slide-number", "1", "--shape-tag", chart["frames"][0]["shape_tag"],
                        "--data-json", str(request_path), "--output", str(destination),
                        "--execute"], check=True)
        source = destination


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True,
                        help="New local directory; raw reports may contain open-deck identities")
    run(parser.parse_args().output_dir.resolve())
