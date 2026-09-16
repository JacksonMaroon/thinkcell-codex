"""Portable changed-data repeat for an existing native multi-break chart."""
from __future__ import annotations
import argparse, json, subprocess
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
import portable_multi_break_adapter as adapter
import run_variant_control as variant


def need(ok, msg):
    if not ok:
        raise RuntimeError(msg)


@variant.runner.serialized_office
def run(args):
    source, input_plan, plan_path, output, report, render = [x.resolve() for x in (args.input, args.input_plan, args.plan, args.output, args.report, args.render)]
    need(source.exists() and not output.exists() and not report.exists() and not render.exists(), "input must exist and outputs must be new")
    need(variant.runner.sha(source) == args.expected_source_sha256.upper(), "source hash mismatch")
    old_plan = json.loads(input_plan.read_text(encoding="utf-8")); plan = json.loads(plan_path.read_text(encoding="utf-8"))
    old_readback = adapter.preflight(source, old_plan, "native")
    stage = output.parent / (output.stem + "_multi_break_repeat_work"); need(not stage.exists(), "stage path exists"); stage.mkdir()
    matrix = plan["targets"][0]["data"]["matrix"]
    cell = lambda value: None if value is None else ({"number": value} if isinstance(value, (int, float)) else {"string": value})
    job = stage / "update.ppttc"; job.write_text(json.dumps([{"template": str(source), "data": [{"name": plan["targets"][0]["name"], "table": [[cell(v) for v in row] for row in matrix]}]}]), encoding="utf-8")
    generated = stage / "generated.pptx"
    with (stage / "ppttc.stdout.txt").open("w", encoding="utf-8") as stdout, (stage / "ppttc.stderr.txt").open("w", encoding="utf-8") as stderr:
        result = variant.runner.run_locked_subprocess([str(args.ppttc), str(job), "-o", str(generated)], operation="portable-multi-break-repeat-json", timeout_seconds=240, stdout=stdout, stderr=stderr, creationflags=subprocess.CREATE_NO_WINDOW)
    need(result.returncode == 0 and generated.exists(), "official JSON generation failed")
    generated_readback = adapter.preflight(generated, plan, "native")
    native, native_report = stage / "native-reopened.pptx", stage / "native-report.json"
    command = [variant.runner.powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(variant.runner.HERE / "thinkcell_no_click" / "implementation" / "native_verify_scoped.ps1"), "-InputFile", str(generated), "-OutputFile", str(native), "-ReportFile", str(native_report), "-RenderFile", str(render)]
    with (stage / "native.stdout.txt").open("w", encoding="utf-8") as stdout, (stage / "native.stderr.txt").open("w", encoding="utf-8") as stderr:
        result = variant.runner.run_locked_subprocess(command, operation="portable-multi-break-repeat-native-reopen", timeout_seconds=240, stdout=stdout, stderr=stderr, env=variant.runner.powershell_env(), creationflags=subprocess.CREATE_NO_WINDOW)
    need(result.returncode == 0 and native.exists() and native_report.exists() and render.exists(), "native save/reopen or render failed")
    scope = json.loads(native_report.read_text(encoding="utf-8-sig")); need(scope.get("native_reopen_pass") and scope.get("source_unchanged") and scope.get("other_presentations_unchanged"), "native scope verification failed")
    native_readback = adapter.preflight(native, plan, "native")
    output.write_bytes(native.read_bytes())
    payload = {"status": "PORTABLE_MULTI_BREAK_REPEAT_PASS", "source_sha256": variant.runner.sha(source), "output_sha256": variant.runner.sha(output), "input": old_readback, "generated": generated_readback, "native_reopened": native_readback, "native_scope": scope, "render": str(render)}
    report.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return payload


if __name__ == "__main__":
    parser = argparse.ArgumentParser(); parser.add_argument("--input", type=Path, required=True); parser.add_argument("--expected-source-sha256", required=True); parser.add_argument("--input-plan", type=Path, required=True); parser.add_argument("--plan", type=Path, required=True); parser.add_argument("--ppttc", type=Path, required=True); parser.add_argument("--output", type=Path, required=True); parser.add_argument("--report", type=Path, required=True); parser.add_argument("--render", type=Path, required=True); args = parser.parse_args(); print(json.dumps(run(args), indent=2))
