"""Build a plan-only structural request from a canonical JSON request."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from generic_matrix_builder import (
    ALLOWED_OPERATIONS,
    MatrixBuilderError,
    add_mekko_category,
    add_waterfall_series,
    add_waterfall_subtotal,
    change_mekko_widths,
)


def _json_arg(raw: str, label: str):
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise MatrixBuilderError(f"{label} must be valid JSON") from exc


def main() -> None:
    ap = argparse.ArgumentParser(description="Prepare a plan-only native structural request")
    ap.add_argument("--request", required=True, type=Path)
    ap.add_argument("--operation", required=True, choices=ALLOWED_OPERATIONS)
    ap.add_argument("--output-plan", required=True, type=Path)
    ap.add_argument("--name")
    ap.add_argument("--values-json", help="JSON array of values, one per named series")
    ap.add_argument("--width", type=float)
    ap.add_argument("--widths-json")
    ap.add_argument("--subtotal-range-json", help="Half-open category range, for example [0, 3]")
    ap.add_argument("--insert-at", type=int)
    ap.add_argument("--equals-series")
    args = ap.parse_args()
    request_path = args.request.resolve()
    output_path = args.output_plan.resolve()
    if not request_path.is_file() or request_path.suffix.lower() != ".json":
        raise MatrixBuilderError(f"Request must be an existing .json: {request_path}")
    if output_path == request_path or output_path.exists():
        raise MatrixBuilderError("Output plan must be a fresh file distinct from the request")
    request = json.loads(request_path.read_text(encoding="utf-8-sig"))
    if args.operation == "waterfall_add_subtotal":
        if not args.name or not args.subtotal_range_json:
            raise MatrixBuilderError("subtotal operation requires --name and --subtotal-range-json")
        value = _json_arg(args.subtotal_range_json, "subtotal range")
        if not isinstance(value, list) or len(value) != 2:
            raise MatrixBuilderError("subtotal range must be a two-item JSON array")
        plan = add_waterfall_subtotal(request, args.name, (value[0], value[1]), equals_series=args.equals_series)
    elif args.operation == "waterfall_add_series":
        if not args.name or args.values_json is None:
            raise MatrixBuilderError("series operation requires --name and --values-json")
        plan = add_waterfall_series(request, args.name, _json_arg(args.values_json, "series values"))
    elif args.operation in {"mekko_percent_add_category", "mekko_units_add_category"}:
        if not args.name or args.values_json is None:
            raise MatrixBuilderError("Mekko category operation requires --name and --values-json")
        chart_kind = "mekko-percent" if args.operation.startswith("mekko_percent") else "mekko-units"
        plan = add_mekko_category(request, chart_kind, args.name, _json_arg(args.values_json, "Mekko values"), width=args.width, insert_at=args.insert_at)
    else:
        if args.widths_json is None:
            raise MatrixBuilderError("width operation requires --widths-json")
        plan = change_mekko_widths(request, _json_arg(args.widths_json, "Mekko widths"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("x", encoding="utf-8") as handle:
        json.dump(plan, handle, indent=2, ensure_ascii=False, allow_nan=False)
    print(json.dumps({"status": plan["status"], "operation": plan["operation"], "plan": str(output_path)}, ensure_ascii=False))


if __name__ == "__main__":
    try:
        main()
    except (MatrixBuilderError, OSError, json.JSONDecodeError) as exc:
        raise SystemExit(str(exc))
