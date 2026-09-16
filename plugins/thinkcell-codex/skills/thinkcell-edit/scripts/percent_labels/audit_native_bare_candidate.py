"""Independent post-native audit for the v8 bare percentage experiment.

This grader deliberately does not use ``select`` from the production audit:
that selector fails closed when the native field has disappeared, which is
the negative result this lane must distinguish from a data failure.  It
records all gates in JSON and exits non-zero on any failed gate.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

HERE = Path(__file__).resolve().parent


def _load_production_modules():
    """Resolve the installed script rail beside this package, independent of CWD."""
    scripts = next((p.resolve() for p in (HERE, HERE.parent, *HERE.parents)
                    if (p / "chart_geometry.py").is_file()), None)
    if scripts is None:
        raise FileNotFoundError("installed think-cell scripts were not found beside this package")
    sys.path.insert(0, str(scripts / "percent_labels"))
    sys.path.insert(0, str(scripts))
    sys.path.insert(0, str(scripts / "thinkcell_no_click" / "implementation"))
    from chart_geometry import inventory  # type: ignore
    from multi_chart_update import model_of  # type: ignore
    from discover_percent_semantics import discover  # type: ignore
    return inventory, model_of, discover


inventory, model_of, discover = _load_production_modules()


def sha256(path: str | Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest().upper()


def package_hashes(path: str | Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {info.filename: hashlib.sha256(package.read(info.filename)).hexdigest().upper()
                for info in package.infolist()}


def expected_percent(numerator: float, denominator: float, digits: int) -> str:
    value = (Decimal(str(numerator)) * Decimal("100") / Decimal(str(denominator))).quantize(
        Decimal("1." + "0" * digits), rounding=ROUND_HALF_UP)
    return f"{value}%"


def normalize_zero_width(value: str) -> str:
    """Remove only the experiment's invisible U+200B wrapper marker."""
    return str(value or "").replace("\u200b", "")


def visible_literal_runs(values: list[str]) -> list[str]:
    """Return visible literal runs after removing U+200B and empty runs."""
    return [normalized for value in values
            if (normalized := normalize_zero_width(value)) != ""]


def has_zero_width_wrapper(values: list[str], expected_display: str) -> bool:
    """Require raw U+200B prefix/suffix while allowing empty native runs."""
    raw = [str(value or "") for value in values]
    return (
        len(raw) >= 3
        and raw[0] == "\u200b"
        and raw[-1] == "\u200b"
        and visible_literal_runs(raw) == [expected_display]
    )


def _one_row(rows: list[dict], *, part: str, category: str, series: str) -> dict | None:
    matches = [r for r in rows if r["chart_part"] == part and r["category"] == category
               and r["series"] == series]
    return matches[0] if len(matches) == 1 else None


def _chart(path: str | Path, part: str):
    raw = Path(path).read_bytes()
    _, charts, _ = inventory(raw)
    matches = [c for c in charts if c["doc"]["part"] == part]
    return matches[0] if len(matches) == 1 else None


def _model_matches(actual: dict, expected: dict) -> bool:
    return (
        actual.get("categories") == expected.get("categories")
        and actual.get("series_names") == expected.get("series_names")
        and actual.get("series_values") == expected.get("series_values")
        and actual.get("category_extents") == expected.get("category_extents")
    )


def _relative_binding(chart: dict | None, row: dict | None) -> dict:
    """Read the authoritative relative source -> text-variable binding.

    Native regeneration may renumber model IDs. The source GUID and the
    actual ``m_ctextvar/elem`` reference are the stable ownership evidence.
    """
    if not chart or not row:
        return {"active": False, "guid": None, "text_variable": None}
    ids = chart.get("doc", {}).get("ids", {})
    source = ids.get(row.get("relative_source_id"))
    if source is None or source.tag != "CVariableSource":
        return {"active": False, "guid": None, "text_variable": None}
    guid_node = source.find("m_guid")
    guid = guid_node.get("val") if guid_node is not None else None
    ref = source.find("m_ctextvar/elem")
    text_id = ref.get("idref") if ref is not None else None
    text_variable = ids.get(text_id)
    active = (bool(text_id) and text_id == row.get("relative_text_variable")
              and text_variable is not None and text_variable.tag == "CTextVariable")
    return {"active": active, "guid": guid, "text_variable": text_id}


def audit(args: argparse.Namespace) -> tuple[dict, bool]:
    source_sha = sha256(args.source)
    candidate_sha = sha256(args.candidate)
    native_sha = sha256(args.input)
    source_rows = discover(args.source)
    candidate_rows = discover(args.candidate)
    native_rows = discover(args.input)
    source_row = _one_row(source_rows, part=args.chart_part, category=args.category, series=args.series)
    candidate_row = _one_row(candidate_rows, part=args.chart_part, category=args.category, series=args.series)
    native_row = _one_row(native_rows, part=args.chart_part, category=args.category, series=args.series)

    with Path(args.plan).open(encoding="utf-8") as handle:
        plan = json.load(handle)
    targets = [t for t in plan.get("targets", []) if t.get("name")]
    plan_target = next((t for t in targets if t["name"] == (native_row or {}).get("chart_name")),
                       targets[0] if len(targets) == 1 else None)
    expected_model = (plan_target or {}).get("data", {}).get("expected_model", {})
    native_chart = _chart(args.input, args.chart_part)
    source_chart = _chart(args.source, args.chart_part)
    native_model = model_of(native_chart) if native_chart else {}
    source_model = model_of(source_chart) if source_chart else {}
    source_binding = _relative_binding(source_chart, source_row)
    candidate_binding = _relative_binding(_chart(args.candidate, args.chart_part), candidate_row)
    native_binding = _relative_binding(native_chart, native_row)

    expected_display = expected_percent(args.numerator, args.denominator, args.precision)
    source_packages = package_hashes(args.source)
    candidate_packages = package_hashes(args.candidate)
    allowed_candidate_diffs = {args.chart_part, "ppt/slides/slide1.xml"}
    candidate_diffs = {name for name in set(source_packages) | set(candidate_packages)
                       if source_packages.get(name) != candidate_packages.get(name)}

    checks: dict[str, bool] = {
        "source_sha256_guard": source_sha == args.source_sha256.upper(),
        "candidate_sha256_guard": candidate_sha == args.candidate_sha256.upper(),
        "candidate_changed_only_selected_chart_and_slide": candidate_diffs == allowed_candidate_diffs,
        "source_selector_resolved": source_row is not None,
        "candidate_selector_resolved": candidate_row is not None,
        "native_selector_resolved": native_row is not None,
        "native_chart_resolved": native_chart is not None,
        "plan_selector_resolved": plan_target is not None,
        "native_model_matches_plan_datasheet": bool(expected_model) and _model_matches(native_model, expected_model),
        "changed_numerator_saved": bool(source_row and native_row)
        and source_row["numerator"] != native_row["numerator"]
        and native_row["numerator"] == args.numerator,
        "changed_denominator_saved": bool(source_row and native_row)
        and source_row["denominator"] != native_row["denominator"]
        and native_row["denominator"] == args.denominator,
        "native_ratio_semantics_exact": bool(native_row)
        and native_row["numerator"] == args.numerator
        and native_row["denominator"] == args.denominator
        and expected_display == expected_percent(native_row["numerator"], native_row["denominator"], args.precision),
        "native_relative_text_variable_preserved": native_binding["active"],
        "native_physical_shape_preserved": bool(native_row and len(native_row.get("physical_shapes", [])) == 1),
        "native_physical_percent_field_preserved": bool(native_row
            and len(native_row.get("physical_shapes", [])) == 1
            and len(native_row["physical_shapes"][0].get("fields", [])) == 1),
        "native_physical_label_exact": bool(native_row
            and len(native_row.get("physical_shapes", [])) == 1
            and len(native_row["physical_shapes"][0].get("fields", [])) == 1
            and native_row["physical_shapes"][0]["fields"][0].get("text", "") == expected_display
            and visible_literal_runs(native_row["physical_shapes"][0].get("literal_texts", []))
                 == [expected_display]),
        "native_relative_suffix_percent": bool(native_row)
        and native_row.get("relative_suffix") == "%",
        "native_precision_exact": bool(native_row)
        and native_row.get("relative_decimal_digits") == str(args.precision),
        "native_relative_source_identity_preserved": source_binding["active"]
        and candidate_binding["active"] and native_binding["active"]
        and source_binding["guid"] == candidate_binding["guid"] == native_binding["guid"],
    }
    if args.allow_zero_width_wrapper:
        candidate_literals = [t for s in (candidate_row or {}).get("physical_shapes", [])
                              for t in s.get("literal_texts", [])]
        checks["candidate_zero_width_wrapper_present"] = candidate_literals.count("\u200b") >= 2
        native_literals = [t for s in (native_row or {}).get("physical_shapes", [])
                           for t in s.get("literal_texts", [])]
        checks["native_raw_zero_width_wrapper_structure"] = has_zero_width_wrapper(
            native_literals, expected_display)

    # Verify that every other semantic label still has a physical native
    # shape and an active percent field.  Their values are expected to change
    # with the plan, so identity/presence is the sibling contract here.
    sibling_failures = []
    for sibling in source_rows:
        key = (sibling["chart_part"], sibling["category"], sibling["series"])
        if key == (args.chart_part, args.category, args.series):
            continue
        other = _one_row(native_rows, part=key[0], category=key[1], series=key[2])
        if (other is None or sibling.get("shape_tag") != other.get("shape_tag")
                or len(other.get("physical_shapes", [])) != 1
                or not other.get("relative_text_variable")
                or not any(str(f.get("text", "")).endswith("%")
                           for f in other["physical_shapes"][0].get("fields", []))):
            sibling_failures.append({"category": key[1], "series": key[2]})
    checks["sibling_native_labels_preserved"] = not sibling_failures

    observed = {
        "source_sha256": source_sha,
        "candidate_sha256": candidate_sha,
        "native_sha256": native_sha,
        "candidate_diffs": sorted(candidate_diffs),
        "expected": {"numerator": args.numerator, "denominator": args.denominator,
                      "display": expected_display, "precision": args.precision,
                      "wrapper": args.wrapper},
        "source_row": source_row,
        "candidate_row": candidate_row,
        "native_row": native_row,
        "source_model": source_model,
        "native_model": native_model,
        "relative_bindings": {"source": source_binding, "candidate": candidate_binding,
                              "native": native_binding},
        "plan_expected_model": expected_model,
        "sibling_failures": sibling_failures,
        "checks": checks,
    }
    if not all(checks.values()):
        if not checks["native_relative_text_variable_preserved"] or not checks["native_physical_shape_preserved"]:
            diagnosis = "NATIVE_BARE_FIELD_DELETED"
        elif not checks["native_ratio_semantics_exact"] or not checks["native_model_matches_plan_datasheet"]:
            diagnosis = "NATIVE_DATA_OR_RATIO_MISMATCH"
        else:
            diagnosis = "NATIVE_PERCENT_FIELD_CONTRACT_FAILED"
        observed["status"] = diagnosis
        return observed, False
    observed["status"] = "NATIVE_BARE_PERCENT_RATIO_PASS"
    return observed, True


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--candidate", required=True)
    parser.add_argument("--candidate-sha256", required=True)
    parser.add_argument("--source-sha256", required=True)
    parser.add_argument("--input", required=True)
    parser.add_argument("--plan", required=True)
    parser.add_argument("--chart-part", required=True)
    parser.add_argument("--category", required=True)
    parser.add_argument("--series", required=True)
    parser.add_argument("--wrapper", choices=["bare"], default="bare")
    parser.add_argument("--allow-zero-width-wrapper", action="store_true",
                        help="normalize U+200B wrapper markers for the bounded closure experiment")
    parser.add_argument("--numerator", type=float, required=True)
    parser.add_argument("--denominator", type=float, required=True)
    parser.add_argument("--precision", type=int, default=1)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()
    result, passed = audit(args)
    Path(args.out).write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps(result, indent=2))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
