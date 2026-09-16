"""Audit the model/physical format contract for one semantic percent label.

The native field's ``type`` is the persisted physical format key while the
model owns the matching ``m_bstrFormat`` and decimal precision.  This audit
keeps those identities, the visible text, and the requested wrapper aligned
before a centrally leased native regeneration.
"""
from __future__ import annotations

import argparse
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
sys.path.insert(0, str(HERE))
from chart_geometry import inventory, streams, xml  # noqa: E402
from multi_chart_update import model_of  # noqa: E402
from discover_percent_semantics import discover, select  # noqa: E402


def _expected(numerator, denominator, digits):
    value = (Decimal(str(numerator)) * Decimal("100") / Decimal(str(denominator))).quantize(
        Decimal("1." + "0" * digits), rounding=ROUND_HALF_UP)
    return f"{value}%"


def _format_key(display: str) -> str:
    return "".join("'" + char + "'" for char in display)


def audit(path: str | Path, *, category: str, series: str, wrapper: str | None = None,
          chart_name: str | None = None, require_canonical_key: bool = False) -> dict:
    path = Path(path)
    raw = path.read_bytes()
    row = select(discover(path), category=category, series=series, chart_name=chart_name)
    chart = next(c for c in inventory(raw)[1] if c["doc"]["part"] == row["chart_part"])
    model = model_of(chart)
    root = xml(streams(chart["doc"]["ole"])[("think-cellXML",)])
    ids = {n.get("id"): n for n in root if n.get("id")}
    textvar = ids.get(row["relative_text_variable"])
    if textvar is None:
        raise ValueError("relative text variable is not present in the chart model")
    format_key = textvar.findtext("m_bstrFormat")
    digits_node = textvar.find("m_prec17834/m_nDecimalDigits17909")
    suffix = textvar.findtext("m_prec17834/m_strSuffix17909")
    if not format_key or digits_node is None:
        raise ValueError("relative model format or precision is missing")
    digits = int(digits_node.get("val"))
    expected = _expected(row["numerator"], row["denominator"], digits)
    shapes = row["physical_shapes"]
    fields = shapes[0]["fields"] if len(shapes) == 1 else []
    field_types = [f.get("type") for f in fields]
    field_texts = [f.get("text") for f in fields]
    expected_literals = [expected] if wrapper == "bare" else ["(", expected, ")"]
    canonical_key = format_key == _format_key(expected)
    native_opaque_key = (not canonical_key and field_types == ["datetime" + format_key]
                         and field_texts == [expected])
    checks = {
        "one_physical_shape": len(shapes) == 1,
        "one_native_field": len(fields) == 1,
        "physical_type_matches_model_format": field_types == ["datetime" + format_key],
        # Prepared candidates must carry the generated key. Native outputs may
        # retain an opaque donor key after official regeneration, provided the
        # physical/model identity and visible precision remain coherent.
        "format_key_matches_requested_precision": (canonical_key or native_opaque_key)
        and (not require_canonical_key or canonical_key),
        "model_suffix_percent": suffix == "%",
        "visible_text_matches_model_ratio": field_texts == [expected],
        "wrapper_matches_requested_mode": wrapper is None or shapes[0]["literal_texts"] == expected_literals,
        "semantic_binding_matches_model": model["series_values"][row["series_index"]][row["category_index"]] == row["numerator"]
        and model["category_extents"][row["category_index"]] == row["denominator"],
    }
    if not all(checks.values()):
        raise ValueError("format consistency audit failed: " + json.dumps(checks, sort_keys=True))
    return {
        "status": "PERCENT_FORMAT_FIELD_CONSISTENCY_PASS",
        "input": str(path),
        "selector": {"category": category, "series": series, "chart_part": row["chart_part"]},
        "model": {"relative_text_variable": row["relative_text_variable"], "format_key": format_key,
                  "decimal_digits": digits, "suffix": suffix},
        "physical": {"shape_tag": row["shape_tag"], "field_id": fields[0]["id"],
                     "field_type": fields[0]["type"], "field_text": fields[0]["text"],
                     "literal_texts": shapes[0]["literal_texts"]},
        "expected_display": expected,
        "format_key_mode": "canonical" if canonical_key else ("native_opaque" if native_opaque_key else "invalid"),
        "require_canonical_key": require_canonical_key,
        "checks": checks,
    }


def audit_candidate(source: str | Path, candidate: str | Path, *, category: str,
                    series: str, wrapper: str, chart_name: str | None = None) -> dict:
    """Prove the candidate transition removed only the absolute field."""
    source_row = select(discover(source), category=category, series=series, chart_name=chart_name)
    candidate_row = select(discover(candidate), category=category, series=series, chart_name=chart_name)
    source_shape = source_row["physical_shapes"][0]
    candidate_shape = candidate_row["physical_shapes"][0]
    source_fields = source_shape["fields"]
    candidate_fields = candidate_shape["fields"]
    checks = {
        "source_has_absolute_and_relative_fields": len(source_fields) == 2,
        "candidate_has_one_native_field": len(candidate_fields) == 1,
        "relative_field_identity_retained": candidate_fields[0]["id"] == source_fields[1]["id"],
        "absolute_field_removed": source_fields[0]["id"] not in {f["id"] for f in candidate_fields},
        "absolute_cached_text_removed": source_fields[0]["text"] not in candidate_shape["literal_texts"],
        "shape_tag_retained": candidate_row["shape_tag"] == source_row["shape_tag"],
        "relative_source_retained": candidate_row["relative_source_id"] == source_row["relative_source_id"],
        "format_contract_pass": audit(candidate, category=category, series=series, wrapper=wrapper,
                                       chart_name=chart_name, require_canonical_key=True)["status"] == "PERCENT_FORMAT_FIELD_CONSISTENCY_PASS",
    }
    if not all(checks.values()):
        raise ValueError("candidate stripping audit failed: " + json.dumps(checks, sort_keys=True))
    return {"status": "PERCENT_CANDIDATE_STRIPPING_AND_FORMAT_PASS",
            "source": str(Path(source)), "candidate": str(Path(candidate)),
            "selector": {"category": category, "series": series}, "checks": checks}


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--category", required=True)
    parser.add_argument("--series", required=True)
    parser.add_argument("--chart-name")
    parser.add_argument("--wrapper", choices=["parentheses", "bare"])
    parser.add_argument("--source")
    parser.add_argument("--out")
    args = parser.parse_args()
    value = (audit_candidate(args.source, args.input, category=args.category, series=args.series,
                              wrapper=args.wrapper, chart_name=args.chart_name)
             if args.source else audit(args.input, category=args.category, series=args.series,
                                       wrapper=args.wrapper, chart_name=args.chart_name))
    payload = json.dumps(value, indent=2)
    if args.out:
        Path(args.out).write_text(payload, encoding="utf-8")
    print(payload)
