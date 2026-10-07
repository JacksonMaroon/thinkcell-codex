"""Run public tests and embedded regressions without launching Office.

Optional --fixture-root contains caller-owned files with the names in
FIXTURE_CHECKS. They are read in place and never copied into the package.
The exact historical profiles are required; arbitrary donor files are unsuitable.
These checks do not renew native save/reopen or visual certifications.
"""
from pathlib import Path
import argparse
import json
import os
import re
import subprocess
import sys

HERE = Path(__file__).resolve().parent
PLUGIN = HERE.parents[2]
PORTABLE_CHECKS = (
    "test_multi_chart_appearance.py",
    "experimental_composition/test_render_tag_remap.py",
    "cagr_import/test_relationship_allocator.py",
)
# None denotes a standalone script accepting --fixture; named globals belong
# to unittest modules whose source fixtures are normally private skill assets.
FIXTURE_CHECKS = (
    ("test_axis_break_insert.py", {"SOURCE": "axis-break-ordinary-38764.pptx", "DONOR": "axis-break-seed-38764.pptx"}),
    ("test_percent_conversion_guards.py", {"FIXTURE": "percent-dual-38764.pptx"}),
    ("experimental_reference_labels/test_prepare_multiline_scalar_format.py", {None: "multiline-scalar-format.pptx"}),
    ("experimental_label_suppression/test_prepare_existing_labels.py", {None: "existing-label-suppression.pptx"}),
    ("experimental_label_suppression/test_direct_precision_profile.py", {None: "direct-precision-profile.pptx"}),
)


def _environment():
    env = os.environ.copy()
    paths = [str(HERE), str(HERE / "thinkcell_no_click/implementation")]
    if env.get("PYTHONPATH"):
        paths.append(env["PYTHONPATH"])
    env["PYTHONPATH"] = os.pathsep.join(paths)
    # Standalone historical regressions use assert statements.
    env.pop("PYTHONOPTIMIZE", None)
    return env


def _execute(name, command, timeout):
    try:
        result = subprocess.run(command, cwd=PLUGIN, env=_environment(),
                                capture_output=True, text=True, timeout=timeout)
        output = result.stdout + result.stderr
        counts = re.findall(r"Ran (\d+) tests? in", output)
        skipped = re.findall(r"skipped=(\d+)", output)
        return {"name": name, "status": "PASS" if result.returncode == 0 else "FAIL",
                "exit_code": result.returncode,
                "unittest_count": sum(map(int, counts)),
                "unittest_skipped": sum(map(int, skipped)), "output": output}
    except (subprocess.TimeoutExpired, OSError) as error:
        return {"name": name, "status": "FAIL", "error": str(error),
                "unittest_count": 0, "unittest_skipped": 0}


def _fixture_command(script, fixtures):
    if None in fixtures:
        return [sys.executable, str(HERE / script), "--fixture", str(fixtures[None])]
    # Import only unittest modules, then substitute the caller's fixture paths.
    # Historical standalone scripts parsing argv are executed separately above.
    code = ("import importlib.util, pathlib, unittest; "
            f"s=importlib.util.spec_from_file_location('embedded_fixture_check', {str(HERE / script)!r}); "
            "m=importlib.util.module_from_spec(s); s.loader.exec_module(m); "
            f"fixtures={ {key: str(value) for key, value in fixtures.items()}!r}; "
            "[setattr(m,k,pathlib.Path(v)) for k,v in fixtures.items()]; "
            "r=unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromModule(m)); "
            "raise SystemExit(0 if r.wasSuccessful() and r.testsRun else 1)")
    return [sys.executable, "-c", code]


def run(fixture_root=None, timeout=180):
    results = [_execute("public unittest discovery", [sys.executable, "-m", "unittest",
                        "discover", "-s", str(PLUGIN / "tests"), "-v"], timeout)]
    for script in PORTABLE_CHECKS:
        results.append(_execute(script, [sys.executable, str(HERE / script)], timeout))
    fixture_root = Path(fixture_root).resolve() if fixture_root else None
    for script, names in FIXTURE_CHECKS:
        fixtures = {key: fixture_root / name for key, name in names.items()} if fixture_root else {}
        missing = [str(path) for path in fixtures.values() if not path.is_file()]
        if fixture_root is None or missing:
            results.append({"name": script, "status": "SKIP", "reason":
                            "Caller-owned historical fixture root was not supplied" if fixture_root is None
                            else "Required historical fixtures missing: " + ", ".join(missing)})
            continue
        results.append(_execute(script, _fixture_command(script, fixtures), timeout))
    return {"status": "FAIL" if any(x["status"] == "FAIL" for x in results) else "PASS",
            "scope": "portable and caller-fixture offline regressions; no native certification",
            "native_execution": False, "checks": results,
            "unittest_count": sum(x.get("unittest_count", 0) for x in results),
            "unittest_skipped": sum(x.get("unittest_skipped", 0) for x in results),
            "fixture_checks_skipped": sum(x["status"] == "SKIP" for x in results)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--fixture-root", type=Path)
    parser.add_argument("--timeout", type=int, default=180, help="Seconds per subprocess")
    parser.add_argument("--report", type=Path, help="Optional fresh JSON report outside package")
    args = parser.parse_args()
    if args.timeout <= 0:
        parser.error("timeout must be positive")
    if args.fixture_root and not args.fixture_root.is_dir():
        parser.error("fixture-root must be an existing directory")
    report = args.report.resolve() if args.report else None
    if report and (report.exists() or report.is_relative_to(PLUGIN) or
                   not report.parent.is_dir() or report.suffix.lower() != ".json"):
        parser.error("report must be a fresh JSON file outside the package in an existing directory")
    result = run(args.fixture_root, args.timeout)
    if report:
        with report.open("x", encoding="utf-8") as stream:
            json.dump(result, stream, indent=2)
    print(json.dumps(result, indent=2))
    return 1 if result["status"] == "FAIL" else 0


if __name__ == "__main__":
    raise SystemExit(main())
