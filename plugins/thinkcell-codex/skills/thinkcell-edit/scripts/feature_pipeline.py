"""One hash-bound, guarded think-cell feature pipeline.

Install beside multi_chart_update.py, enable_datasheet_fill.py, their shared
helpers, and office_operation_lock.py. Pure package preparation happens before
one official JSON/native update. Requested research-only features are rejected
before any write until their visual verification gates exist.
"""
from __future__ import annotations

import argparse, hashlib, json, os, sys
from pathlib import Path
from types import SimpleNamespace

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

multi = fill_enable = None

def need(condition, message):
    if not condition:
        raise ValueError(message)

def _dependencies():
    """Load native package adapters only for an actual pipeline run.

    Keeping plan-schema validation independent makes it usable in lightweight
    environments that do not have the bundled CFB parser installed.
    """
    global multi, fill_enable, inventory, _select_chart
    if multi is None:
        import multi_chart_update as multi_module
        import enable_datasheet_fill as fill_module
        from chart_geometry import inventory as inventory_fn, _select_chart as select_fn
        multi, fill_enable = multi_module, fill_module
        inventory, _select_chart = inventory_fn, select_fn
    return multi

try:
    from office_operation_lock import OfficeOperationLock
except ImportError:  # Allows pure schema checks before lock installation.
    OfficeOperationLock = None

SCHEMA = "tc.feature-pipeline.v1"
KINDS = {"datasheet_fill_enable", "total_label_precision", "label_precision", "axis_break", "relative_label_content"}


def sha(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest().upper()


def load(path: Path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def dump(path: Path, value) -> None:
    Path(path).write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def _selection(feature):
    selection = feature.get("selector")
    need(isinstance(selection, dict) and set(selection) in ({"slide_number", "shape_tag"}, {"slide_id", "shape_tag"}),
         "Feature selector requires exactly slide_number/slide_id and shape_tag")
    return selection


def validate_plan(plan):
    """Validate the pure JSON contract without opening Office or a deck."""
    need(isinstance(plan, dict) and set(plan) == {"schema", "source_sha256", "data_plan", "features"},
         "Feature pipeline plan must contain schema, source_sha256, data_plan, features")
    need(plan["schema"] == SCHEMA and isinstance(plan["source_sha256"], str) and len(plan["source_sha256"]) == 64,
         "Unsupported or incomplete feature pipeline plan")
    need(isinstance(plan["data_plan"], dict) and isinstance(plan["features"], list) and plan["features"],
         "data_plan and non-empty features are required")
    seen = set()
    for feature in plan["features"]:
        need(isinstance(feature, dict) and isinstance(feature.get("kind"), str) and feature["kind"] in KINDS,
             "Unknown feature kind")
        selection = _selection(feature); key = (feature["kind"], tuple(sorted(selection.items())))
        need(key not in seen, "Duplicate feature target")
        seen.add(key)
        allowed = {
            "datasheet_fill_enable": {"kind", "selector"},
            "label_precision": {"kind", "selector", "decimal_digits"},
            "total_label_precision": {"kind", "selector", "decimal_digits"},
            "axis_break": {"kind", "selector", "fraction"},
            "relative_label_content": {"kind", "relative_field_id"},
        }[feature["kind"]]
        need(set(feature) == allowed, "Unexpected or missing feature fields")
        if feature["kind"] in {"label_precision", "total_label_precision"}:
            need(isinstance(feature["decimal_digits"], int) and not isinstance(feature["decimal_digits"], bool)
                 and 0 <= feature["decimal_digits"] <= 3, "decimal_digits must be an integer from 0 through 3")
        if feature["kind"] == "axis_break":
            need(isinstance(feature["fraction"], (int, float)) and not isinstance(feature["fraction"], bool)
                 and .05 <= feature["fraction"] <= .95, "Axis-break fraction must be 0.05 through 0.95")
        if feature["kind"] == "relative_label_content":
            need(isinstance(feature["relative_field_id"], str) and feature["relative_field_id"],
                 "relative_field_id is required")
    return plan


def _make_preparation(current: Path, feature: dict):
    _dependencies()
    kind = feature["kind"]
    if kind == "relative_label_content":
        raise ValueError("relative_label_content remains experimental because no post-regeneration verifier exists")
    if kind == "label_precision":
        raise ValueError("General label_precision is unverified; use total_label_precision for its bounded one-category total")
    if kind == "total_label_precision":
        import total_label_precision
        return total_label_precision, total_label_precision.make_plan(current, {"shape_tag": _selection(feature)["shape_tag"]}, feature["decimal_digits"])
    if kind == "axis_break":
        raise ValueError("Use axis_break_position.py for the verified dedicated native break-position route")
    if kind == "datasheet_fill_enable":
        return fill_enable, fill_enable.make_plan(current, _selection(feature))
    raise AssertionError(kind)


def _verify_fill(path: Path, prepared_plan):
    _dependencies()
    _, charts, _ = inventory(path.read_bytes())
    chart = _select_chart(charts, {"shape_tag": prepared_plan["target"]["shape_tag"]})
    node = chart["table"].find(fill_enable.FIELD)
    need(node is not None and node.get("val") == "1", "Datasheet-fill setting did not survive regeneration")
    return {"status": "DATASHEET_FILL_SETTING_RETAINED", "shape_tag": prepared_plan["target"]["shape_tag"]}


def run(input_path: Path, output_path: Path, report_path: Path, plan: dict, *, execute: bool, ppttc=None):
    validate_plan(plan)
    _dependencies()
    src, out, report = Path(input_path).resolve(), Path(output_path).resolve(), Path(report_path).resolve()
    need(src.is_file() and sha(src) == plan["source_sha256"].upper(), "Input hash mismatch")
    need(len({src, out, report}) == 3 and out.suffix.lower() == ".pptx", "Input, output, report must be distinct and output PPTX")
    need(not out.exists() and not report.exists(), "Output/report already exists")
    # Validate data coverage before any guarded preparation. This also rejects a
    # partial multi-chart update which could otherwise silently regenerate a sibling.
    multi.validate_plan(src.read_bytes(), plan["data_plan"])
    if not execute:
        planned = []
        # Dry-run deliberately does not call prepare: all feature plans are
        # independently hash-bound to the unchanged source and prove that the
        # intended guarded chain can start, but no PPTX/staging path is made.
        for feature in plan["features"]:
            _, feature_plan = _make_preparation(src, feature)
            planned.append({"kind": feature["kind"], "source_sha256": feature_plan["source_sha256"],
                            "required_followup": feature_plan.get("required_followup", ["official regeneration", "selected dynamic total field readback", "native reopen"])})
        return {"status": "DRY_RUN_PASS", "source_sha256": sha(src), "feature_plans": planned,
                "pending_chain_validation": "Preparations, one official JSON regeneration, and final feature readback run only with --execute.",
                "no_presentation_or_staging_output_created": True}
    stage = out.parent / (out.stem + "_feature_work")
    need(not stage.exists(), "Feature staging directory already exists")
    stage.mkdir(); current = src; prepared = []
    try:
        for index, feature in enumerate(plan["features"], 1):
            module, feature_plan = _make_preparation(current, feature)
            candidate = stage / f"{index:02d}-{feature['kind']}.pptx"
            result = module.prepare(current, candidate, feature_plan)
            need(sha(candidate) == result["output_sha256"], "Preparation output hash mismatch")
            prepared.append({"feature": feature, "plan": feature_plan, "result": result, "path": str(candidate)})
            current = candidate
        multi.validate_plan(current.read_bytes(), plan["data_plan"])
        generated = stage / "official-regenerated.pptx"
        gate_report = stage / "multi-chart-gates.json"
        args = SimpleNamespace(input=current, expected_sha256=sha(current), plan=stage / "data-plan.json", output=generated,
                               report=gate_report, prepare=True, execute=True, ppttc=ppttc)
        dump(args.plan, plan["data_plan"])
        if OfficeOperationLock is None:
            raise RuntimeError("office_operation_lock.py is required for execution")
        with OfficeOperationLock("feature-pipeline", metadata={"input": str(src), "features": [x["feature"]["kind"] for x in prepared]}):
            gates = multi.run(args)
        dump(gate_report, gates)
        checks = []
        for item in prepared:
            kind = item["feature"]["kind"]
            if kind == "datasheet_fill_enable": checks.append(_verify_fill(generated, item["plan"]))
            elif kind == "total_label_precision":
                import total_label_precision
                tag = item["plan"]["target"]["shape_tag"]
                targets = [target for target in plan["data_plan"]["targets"] if target["selector"].get("shape_tag") == tag]
                need(len(targets) == 1, "Total precision needs one exact data target")
                expected = targets[0]["data"]["expected_model"]
                need(len(expected["categories"]) == 1, "Total precision needs one category")
                total = sum(row[0] or 0 for row in expected["series_values"])
                checks.append(total_label_precision.verify(generated, item["plan"], total))
        need(sha(src) == plan["source_sha256"].upper(), "Source changed during pipeline")
        with out.open("xb") as handle: handle.write(generated.read_bytes())
        return {"status": "ALL_GATES_PASS", "source_sha256": sha(src), "output": str(out), "output_sha256": sha(out),
                "preparations": prepared, "official_json_gates": gates, "feature_verifiers": checks,
                "source_unchanged": True}
    except Exception:
        # Preserve staging evidence for diagnosis; never retry a native failure blindly.
        raise


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True); parser.add_argument("--expected-sha256", required=True)
    parser.add_argument("--plan", type=Path, required=True); parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True); parser.add_argument("--execute", action="store_true"); parser.add_argument("--ppttc")
    args = parser.parse_args(); plan = load(args.plan); need(args.expected_sha256.upper() == plan.get("source_sha256", "").upper(), "CLI hash differs from plan hash")
    try:
        result = run(args.input, args.output, args.report, plan, execute=args.execute, ppttc=args.ppttc)
        dump(args.report, result); print(json.dumps({"status": result["status"], "report": str(args.report.resolve())}))
    except Exception as error:
        print(json.dumps({"status": "REJECTED", "error": str(error)}), file=sys.stderr); raise SystemExit(1)


if __name__ == "__main__": main()
