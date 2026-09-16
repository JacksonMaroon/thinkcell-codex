"""Recover and audit an already-produced native percent output offline."""
from __future__ import annotations

import argparse
import hashlib
import json
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from audit_format_consistency import audit
from discover_percent_semantics import discover, select


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def display(numerator, denominator, digits):
    value = (Decimal(str(numerator)) * 100 / Decimal(str(denominator))).quantize(
        Decimal("1." + "0" * digits), rounding=ROUND_HALF_UP)
    return f"{value}%"


def recover(path: Path, plan_path: Path, *, category: str, series: str,
            wrapper: str, input_sha256: str | None = None) -> dict:
    plan = json.loads(plan_path.read_text(encoding="utf-8-sig"))
    target = plan["targets"][0]
    matrix = target["data"]["matrix"]
    row = select(discover(path), category=category, series=series)
    ci, si = row["category_index"], row["series_index"]
    numerator, denominator = matrix[si + 2][ci + 1], matrix[1][ci + 1]
    expected_model = target["data"]["expected_model"]
    audit_result = audit(path, category=category, series=series, wrapper=wrapper)
    expected = display(numerator, denominator, int(row["relative_decimal_digits"]))
    if row["physical_shapes"][0]["fields"][0]["text"] != expected:
        raise ValueError("saved field text does not match plan ratio")
    if row["numerator"] != numerator or row["denominator"] != denominator:
        raise ValueError("saved semantic ratio does not match plan")
    result = {
        "status": "RECOVERED_NATIVE_READBACK_PASS",
        "output": str(path),
        "output_sha256": sha(path),
        "input_sha256": input_sha256,
        "selector": {"category": category, "series": series, "chart_part": row["chart_part"]},
        "expected_display": expected,
        "saved_model": {"numerator": row["numerator"], "denominator": row["denominator"],
                        "plan_expected_model": expected_model},
        "format_audit": audit_result,
        "native_execution": "completed_before_offline_recovery",
    }
    return result


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--input", type=Path, required=True)
    p.add_argument("--plan", type=Path, required=True)
    p.add_argument("--category", required=True)
    p.add_argument("--series", required=True)
    p.add_argument("--wrapper", choices=["bare", "parentheses"], default="parentheses")
    p.add_argument("--input-sha256")
    p.add_argument("--output", type=Path, required=True)
    a = p.parse_args()
    value = recover(a.input, a.plan, category=a.category, series=a.series,
                    wrapper=a.wrapper, input_sha256=a.input_sha256)
    a.output.write_text(json.dumps(value, indent=2), encoding="utf-8")
    print(json.dumps(value, indent=2))
