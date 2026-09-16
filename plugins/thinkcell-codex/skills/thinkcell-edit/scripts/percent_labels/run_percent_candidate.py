"""Run official/native regeneration on a prepared percent candidate.

This helper is intentionally separate from the v4 COM conversion. It accepts
the candidate's existing one-field physical label and lets official think-cell
regeneration decide whether the coherent model/field pair is supported.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from portable_scripts import resolve_plugin_scripts  # noqa: E402
SCRIPTS = resolve_plugin_scripts(HERE)
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "thinkcell_no_click" / "implementation"))
from chart_geometry import inventory  # noqa: E402
from multi_chart_update import model_of, run as official_run  # noqa: E402
from discover_percent_semantics import discover, select  # noqa: E402
from audit_format_consistency import audit as audit_format  # noqa: E402


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def need(value, message):
    if not value:
        raise ValueError(message)


def expected_display(numerator, denominator, digits):
    value = (Decimal(str(numerator)) * 100 / Decimal(str(denominator))).quantize(
        Decimal("1." + "0" * digits), rounding=ROUND_HALF_UP)
    return f"{value}%"


def verify(path, plan, source_row, wrapper):
    target = plan["targets"][0]
    matrix = target["data"]["matrix"]
    ci, si = source_row["category_index"], source_row["series_index"]
    numerator, denominator = matrix[si + 2][ci + 1], matrix[1][ci + 1]
    digits = int(source_row["relative_decimal_digits"] or 0)
    expected = expected_display(numerator, denominator, digits)
    row = select(discover(path), category=source_row["category"], series=source_row["series"],
                 chart_part=source_row["chart_part"], chart_name=source_row["chart_name"])
    need(len(row["physical_shapes"]) == 1, "final semantic label missing or ambiguous")
    shape = row["physical_shapes"][0]
    expected_literals = [expected] if wrapper == "bare" else ["(", expected, ")"]
    need(shape["literal_texts"] == expected_literals, "final wrapper/value mismatch")
    need(len(shape["fields"]) == 1 and shape["fields"][0]["text"] == expected,
         "final label is not one native field")
    chart = next(c for c in inventory(path.read_bytes())[1] if c["doc"]["part"] == source_row["chart_part"])
    model = model_of(chart)
    need(model["series_values"][si][ci] == numerator, "saved numerator mismatch")
    need(model["category_extents"][ci] == denominator, "saved denominator mismatch")
    need(row["relative_decimal_digits"] == str(digits), "saved precision mismatch")
    return {"expected_display": expected, "wrapper": wrapper, "selected": row,
            "saved_ratio": {"numerator": numerator, "denominator": denominator}}


def run(args):
    source = Path(args.input).resolve()
    output = Path(args.output).resolve()
    report = Path(args.report).resolve()
    evidence = Path(args.evidence).resolve()
    plan = Path(args.plan).resolve()
    need(source.exists() and sha(source) == args.expected_sha256.upper(), "source hash mismatch")
    need(all(not p.exists() for p in (output, report, evidence)), "outputs must be new")
    source_row = select(discover(source), category=args.category, series=args.series,
                        chart_part=args.chart_part, chart_name=args.chart_name)
    expected_chart_tag = next(c["frames"][0]["shape_tag"] for c in inventory(source.read_bytes())[1]
                              if c["doc"]["part"] == source_row["chart_part"])
    plan_obj = json.loads(plan.read_text(encoding="utf-8-sig"))
    need(plan_obj["targets"][0]["selector"].get("shape_tag") == expected_chart_tag,
         "plan chart selector does not match source chart")
    stage_report = output.parent / (output.stem + "-official.json")
    result = official_run(argparse.Namespace(input=source, expected_sha256=args.expected_sha256,
                                             plan=plan, output=output, report=stage_report,
                                             prepare=True, execute=True, ppttc=None))
    final = verify(output, plan_obj, source_row, args.wrapper)
    final["format_consistency"] = audit_format(
        output, category=source_row["category"], series=source_row["series"], wrapper=args.wrapper,
        chart_name=source_row["chart_name"])
    result["percent_candidate"] = final
    result["source_unchanged"] = sha(source) == args.expected_sha256.upper()
    with report.open("xb") as handle:
        handle.write(json.dumps(result, indent=2).encode())
    evidence_obj = {"status": "PERCENT_CANDIDATE_NATIVE_PASS", "input_sha256": args.expected_sha256.upper(),
                    "output_sha256": sha(output), "selector": {"category": args.category,
                    "series": args.series, "chart_part": source_row["chart_part"]},
                    "wrapper": args.wrapper, "result": final, "official_report": str(report)}
    with evidence.open("xb") as handle:
        handle.write(json.dumps(evidence_obj, indent=2).encode())
    return evidence_obj


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--input", required=True)
    p.add_argument("--expected-sha256", required=True)
    p.add_argument("--plan", required=True)
    p.add_argument("--category", required=True)
    p.add_argument("--series", required=True)
    p.add_argument("--chart-part")
    p.add_argument("--chart-name")
    p.add_argument("--wrapper", choices=["bare", "parentheses"], default="bare")
    p.add_argument("--output", required=True)
    p.add_argument("--report", required=True)
    p.add_argument("--evidence", required=True)
    a = p.parse_args()
    try:
        print(json.dumps(run(a), indent=2))
    except Exception as exc:
        print(json.dumps({"status": "REJECTED", "error": str(exc)}), file=sys.stderr)
        raise SystemExit(1)
