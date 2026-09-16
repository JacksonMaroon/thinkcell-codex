"""Plan-only structural builders for canonical waterfall and Mekko requests.

The builders operate on a request containing ``matrix`` and ``expected_model``
and return a cloned request plus explicit semantic guards. They do not open
Office, call ``ppttc.exe``, or write a presentation. A plan always declares
that official regeneration is required before native readback is accepted.
"""
from __future__ import annotations

import copy
import math
from typing import Any, Iterable


class MatrixBuilderError(ValueError):
    pass


ALLOWED_OPERATIONS = (
    "waterfall_add_subtotal",
    "waterfall_add_series",
    "mekko_percent_add_category",
    "mekko_units_add_category",
    "mekko_units_width_change",
)


def _finite(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def _canonical(request: dict, expected_kind: str | None = None) -> tuple[dict, list, int, dict[str, int]]:
    if not isinstance(request, dict):
        raise MatrixBuilderError("request must be an object")
    model = request.get("expected_model")
    matrix = request.get("matrix")
    if not isinstance(model, dict) or not isinstance(matrix, list) or not matrix or any(not isinstance(row, list) for row in matrix):
        raise MatrixBuilderError("request must contain matrix and expected_model")
    categories = model.get("categories")
    names = model.get("series_names")
    values = model.get("series_values")
    if not isinstance(categories, list) or not categories or not isinstance(names, list) or not names:
        raise MatrixBuilderError("expected_model needs nonempty categories and series_names")
    if not isinstance(values, list) or len(values) != len(names) or any(not isinstance(row, list) or len(row) != len(categories) for row in values):
        raise MatrixBuilderError("series_values dimensions do not match categories and series_names")
    width = len(matrix[0])
    if width != len(categories) + 1 or width < 2 or any(len(row) != width for row in matrix):
        raise MatrixBuilderError("matrix must be rectangular with one row-name column")
    if expected_kind and request.get("kind", expected_kind) != expected_kind:
        raise MatrixBuilderError(f"request kind must be {expected_kind}")
    header_rows = [i for i, row in enumerate(matrix) if row[1:] == categories]
    if len(header_rows) != 1:
        raise MatrixBuilderError("could not uniquely identify the category header row")
    header = header_rows[0]
    row_by_name: dict[str, int] = {}
    for i, row in enumerate(matrix[header + 1 :], header + 1):
        if row and row[0] in names:
            if row[0] in row_by_name:
                raise MatrixBuilderError(f"series row {row[0]!r} occurs more than once")
            row_by_name[row[0]] = i
    if set(row_by_name) != set(names):
        raise MatrixBuilderError("matrix does not contain every named series")
    return model, matrix, header, row_by_name


def _recompute_waterfall(request: dict) -> None:
    model, matrix, header, row_by_name = _canonical(request, "waterfall")
    running = 0.0
    for ci in range(len(model["categories"])):
        total = 0.0
        equals: list[int] = []
        for si, name in enumerate(model["series_names"]):
            cell = matrix[row_by_name[name]][ci + 1]
            if cell == "e":
                equals.append(si)
                model["series_values"][si][ci] = None
            elif _finite(cell):
                value = float(cell)
                model["series_values"][si][ci] = value
                total += value
            else:
                model["series_values"][si][ci] = None
        if equals:
            for si in equals:
                model["series_values"][si][ci] = running
        else:
            running += total
    if "category_extents" in model:
        model["category_extents"] = [
            None if any(matrix[row_by_name[name]][ci + 1] == "e" for name in model["series_names"])
            else sum((model["series_values"][si][ci] or 0.0) for si in range(len(model["series_names"])))
            for ci in range(len(model["categories"]))
        ]


def _recompute_mekko(request: dict, chart_kind: str) -> None:
    model, matrix, header, row_by_name = _canonical(request, chart_kind)
    width_row = header + 1
    for si, name in enumerate(model["series_names"]):
        model["series_values"][si] = [
            float(matrix[row_by_name[name]][ci + 1]) if _finite(matrix[row_by_name[name]][ci + 1]) else None
            for ci in range(len(model["categories"]))
        ]
    if chart_kind == "mekko-percent":
        widths = []
        for ci in range(len(model["categories"])):
            cells = [model["series_values"][si][ci] or 0.0 for si in range(len(model["series_names"]))]
            if any(value < 0 for value in cells):
                raise MatrixBuilderError("percent-Mekko values must be nonnegative")
            widths.append(sum(cells))
        model["column_widths"] = widths
    else:
        widths = matrix[width_row][1:]
        if any(not _finite(value) or value <= 0 for value in widths):
            raise MatrixBuilderError("units-Mekko width row must contain positive finite values")
        model["column_widths"] = [float(value) for value in widths]


def _plan(operation: str, request: dict, guards: dict[str, Any]) -> dict:
    if operation not in ALLOWED_OPERATIONS:
        raise MatrixBuilderError(f"operation {operation!r} is not in the installed-ready allowlist")
    return {
        "status": "STRUCTURAL_PLAN_READY",
        "operation": operation,
        "official_regeneration_required": True,
        "native_readback_required": True,
        "semantic_guards": guards,
        "request": request,
    }


def add_waterfall_subtotal(
    request: dict,
    subtotal_name: str,
    subtotal_range: tuple[int, int],
    *,
    equals_series: str | None = None,
) -> dict:
    """Insert one prefix subtotal using an exact half-open category range.

    ``subtotal_range`` is ``(start, end)`` and denotes categories
    ``start:end``. A think-cell equals cell is a running prefix total, so the
    exact guard requires ``start == 0`` and insertion at ``end``.
    """
    model, matrix, header, row_by_name = _canonical(request, "waterfall")
    n = len(model["categories"])
    if not isinstance(subtotal_name, str) or not subtotal_name.strip():
        raise MatrixBuilderError("subtotal_name must be nonempty text")
    if subtotal_name in model["categories"]:
        raise MatrixBuilderError("subtotal_name must be new")
    if len(subtotal_range) != 2 or any(not isinstance(value, int) or isinstance(value, bool) for value in subtotal_range):
        raise MatrixBuilderError("subtotal_range must be a half-open (start, end) pair")
    start, end = subtotal_range
    if start != 0 or not (0 < end <= n):
        raise MatrixBuilderError("subtotal_range must be a nonempty prefix within categories")
    if any(matrix[row][ci + 1] == "e" for row in row_by_name.values() for ci in range(start, end)):
        raise MatrixBuilderError("subtotal_range cannot contain an existing equals cell")
    if equals_series is None:
        terminal = [name for name, row in row_by_name.items() if matrix[row][-1] == "e"]
        if len(terminal) != 1:
            raise MatrixBuilderError("equals_series is required unless exactly one terminal equals series exists")
        equals_series = terminal[0]
    if equals_series not in row_by_name:
        raise MatrixBuilderError("equals_series is not a named matrix series")
    result = copy.deepcopy(request)
    result_model = result["expected_model"]
    result_matrix = result["matrix"]
    result_model["categories"].insert(end, subtotal_name)
    for row in result_model["series_values"]:
        row.insert(end, None)
    for ri, row in enumerate(result_matrix):
        cell = subtotal_name if ri == header else "e" if ri == row_by_name[equals_series] else None
        row.insert(end + 1, cell)
    _recompute_waterfall(result)
    expected_total = sum(
        float(result_matrix[row_by_name[name]][ci + 1])
        for name in row_by_name
        for ci in range(start, end)
        if _finite(result_matrix[row_by_name[name]][ci + 1])
    )
    actual_total = result["expected_model"]["series_values"][list(result["expected_model"]["series_names"]).index(equals_series)][end]
    if actual_total != expected_total:
        raise MatrixBuilderError("subtotal equals value does not match the requested subtotal range")
    return _plan("waterfall_add_subtotal", result, {"equals_subtotal_range": [start, end], "equals_series": equals_series, "insert_at": end})


def add_waterfall_series(request: dict, series_name: str, values: Iterable[Any]) -> dict:
    model, matrix, header, row_by_name = _canonical(request, "waterfall")
    values = list(values)
    if not isinstance(series_name, str) or not series_name.strip() or series_name in model["series_names"]:
        raise MatrixBuilderError("series_name must be new nonempty text")
    if len(values) != len(model["categories"]) or any(value is not None and not _finite(value) for value in values):
        raise MatrixBuilderError("series values must match category count and be finite numbers or None")
    result = copy.deepcopy(request)
    result_model = result["expected_model"]
    result_matrix = result["matrix"]
    result_model["series_names"].append(series_name)
    result_model["series_values"].append([float(value) if value is not None else None for value in values])
    # Keep leading reserved rows in place. If the source has trailing
    # reserved rows, insert immediately before them; otherwise append after
    # the last named series.
    last_series_row = max(row_by_name.values())
    insert_row = next(
        (i for i, row in enumerate(result_matrix[last_series_row + 1 :], last_series_row + 1) if not row or row[0] not in row_by_name),
        last_series_row + 1,
    )
    result_matrix.insert(insert_row, [series_name, *values])
    _recompute_waterfall(result)
    return _plan("waterfall_add_series", result, {"series_count": len(result_model["series_names"]), "insert_before_reserved_rows": insert_row})


def add_mekko_category(
    request: dict,
    chart_kind: str,
    category_name: str,
    values: Iterable[Any],
    *,
    width: float | None = None,
    insert_at: int | None = None,
) -> dict:
    operation = "mekko_percent_add_category" if chart_kind == "mekko-percent" else "mekko_units_add_category" if chart_kind == "mekko-units" else None
    if operation is None:
        raise MatrixBuilderError("chart_kind must be mekko-percent or mekko-units")
    model, matrix, header, row_by_name = _canonical(request, chart_kind)
    values = list(values)
    if not isinstance(category_name, str) or not category_name.strip():
        raise MatrixBuilderError("category_name must be nonempty text")
    if len(values) != len(model["series_names"]) or any(not _finite(value) or value < 0 for value in values):
        raise MatrixBuilderError("Mekko values must be nonnegative finite numbers for every named series")
    if category_name in model["categories"]:
        raise MatrixBuilderError("category_name must be new")
    insert_at = len(model["categories"]) if insert_at is None else insert_at
    if not isinstance(insert_at, int) or isinstance(insert_at, bool) or not 0 <= insert_at <= len(model["categories"]):
        raise MatrixBuilderError("insert_at is outside the category range")
    derived_width = float(sum(values)) if chart_kind == "mekko-percent" else width
    if derived_width is None or not _finite(derived_width) or derived_width <= 0:
        raise MatrixBuilderError("units-Mekko requires a positive finite width")
    result = copy.deepcopy(request)
    result_model = result["expected_model"]
    result_matrix = result["matrix"]
    result_model["categories"].insert(insert_at, category_name)
    for si, row in enumerate(result_model["series_values"]):
        row.insert(insert_at, float(values[si]))
    for ri, row in enumerate(result_matrix):
        if ri == header:
            cell = category_name
        elif ri == header + 1:
            cell = float(derived_width) if chart_kind == "mekko-units" else None
        elif row and row[0] in row_by_name:
            cell = float(values[model["series_names"].index(row[0])])
        else:
            cell = None
        row.insert(insert_at + 1, cell)
    _recompute_mekko(result, chart_kind)
    if chart_kind == "mekko-percent" and result["expected_model"]["column_widths"][insert_at] != derived_width:
        raise MatrixBuilderError("percent-Mekko width must equal the absolute segment total")
    return _plan(operation, result, {"category_count": len(result_model["categories"]), "width_semantics": "column_total" if chart_kind == "mekko-percent" else "independent_width", "insert_at": insert_at})


def change_mekko_widths(request: dict, widths: Iterable[Any]) -> dict:
    model, matrix, header, _ = _canonical(request, "mekko-units")
    widths = list(widths)
    if len(widths) != len(model["categories"]) or any(not _finite(value) or value <= 0 for value in widths):
        raise MatrixBuilderError("widths must contain one positive finite value per category")
    result = copy.deepcopy(request)
    result["matrix"][header + 1][1:] = [float(value) for value in widths]
    result["expected_model"]["column_widths"] = [float(value) for value in widths]
    _recompute_mekko(result, "mekko-units")
    return _plan("mekko_units_width_change", result, {"category_count": len(widths), "width_semantics": "independent_width"})
