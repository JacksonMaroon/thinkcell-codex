"""Generalized official-JSON adapter for adding one waterfall step.

The adapter operates on a canonical think-cell request consisting of
``matrix`` and ``expected_model``. It discovers the terminal equals category
from the request, inserts a category immediately before it, preserves all
reserved rows, and recalculates expected equals values from the resulting
matrix. No donor name, automation ID, category label, or hard-coded value is
embedded in this module.
"""
from __future__ import annotations

import copy
import math
from typing import Any, Iterable


class WaterfallAdapterError(ValueError):
    pass


def _number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def _model(request: dict) -> dict:
    try:
        model = request["expected_model"]
        categories = model["categories"]
        names = model["series_names"]
        values = model["series_values"]
    except (KeyError, TypeError) as exc:
        raise WaterfallAdapterError("request lacks expected_model categories/series fields") from exc
    if not isinstance(categories, list) or not categories or not isinstance(names, list) or not names:
        raise WaterfallAdapterError("waterfall model needs nonempty categories and series_names")
    if len(values) != len(names) or any(len(row) != len(categories) for row in values):
        raise WaterfallAdapterError("series_values dimensions do not match the model")
    return model


def _header_and_terminal_equals(request: dict) -> tuple[int, int]:
    model = _model(request)
    matrix = request.get("matrix")
    categories = model["categories"]
    if not isinstance(matrix, list) or not matrix or any(not isinstance(row, list) for row in matrix):
        raise WaterfallAdapterError("request matrix must be a nonempty rectangular list")
    width = len(matrix[0])
    if width < 2 or any(len(row) != width for row in matrix):
        raise WaterfallAdapterError("request matrix must be rectangular")
    header_rows = [i for i, row in enumerate(matrix) if row[1 : 1 + len(categories)] == categories]
    if len(header_rows) != 1:
        raise WaterfallAdapterError("could not uniquely identify the category header row")
    header = header_rows[0]
    equals_columns = []
    for row in matrix[header + 1 :]:
        for col, cell in enumerate(row[1:], 1):
            if cell == "e":
                equals_columns.append(col - 1)
    if not equals_columns:
        raise WaterfallAdapterError("waterfall request has no equals cell")
    terminal = max(equals_columns)
    if terminal != len(categories) - 1:
        raise WaterfallAdapterError("terminal equals cell is not at the final category")
    return header, terminal


def _recalculate_model(request: dict) -> None:
    model = _model(request)
    matrix = request["matrix"]
    names = model["series_names"]
    rows = {row[0]: row for row in matrix[1:] if row and row[0] in names}
    if set(rows) != set(names):
        raise WaterfallAdapterError("matrix does not contain every named waterfall series")
    running = 0.0
    for ci in range(len(model["categories"])):
        total = 0.0
        equals: list[int] = []
        for si, name in enumerate(names):
            cell = rows[name][ci + 1]
            if cell == "e":
                equals.append(si)
                model["series_values"][si][ci] = None
            elif _number(cell):
                total += float(cell)
                model["series_values"][si][ci] = float(cell)
            else:
                model["series_values"][si][ci] = None
        if equals:
            for si in equals:
                model["series_values"][si][ci] = running
        else:
            running += total
    if "category_extents" in model:
        model["category_extents"] = [
            None if any(row[ci + 1] == "e" for row in rows.values())
            else sum((model["series_values"][si][ci] or 0.0) for si in range(len(names)))
            for ci in range(len(model["categories"]))
        ]


def add_step(request: dict, category_name: str, contributions: Iterable[Any]) -> dict:
    """Return a cloned request with one contribution category before final ``e``."""
    if not isinstance(category_name, str) or not category_name.strip():
        raise WaterfallAdapterError("category_name must be nonempty text")
    model = _model(request)
    contributions = list(contributions)
    if len(contributions) != len(model["series_names"]):
        raise WaterfallAdapterError("contributions length must equal series count")
    if any(value is not None and not _number(value) for value in contributions):
        raise WaterfallAdapterError("contributions must be finite numbers or None")
    header, terminal = _header_and_terminal_equals(request)
    result = copy.deepcopy(request)
    result_model = result["expected_model"]
    result_model["categories"].insert(terminal, category_name)
    for si, value in enumerate(contributions):
        result_model["series_values"][si].insert(terminal, float(value) if value is not None else None)
    matrix = result["matrix"]
    series_names = result_model["series_names"]
    by_name = {name: value for name, value in zip(series_names, contributions)}
    for ri, row in enumerate(matrix):
        if ri == header:
            row.insert(terminal + 1, category_name)
        else:
            row.insert(terminal + 1, by_name.get(row[0]) if row and row[0] in by_name else None)
    _recalculate_model(result)
    return result


def build_ppttc_job(template: str, request: dict) -> list[dict]:
    """Build the official generator job without depending on lane fixtures."""
    name = request.get("name") or request.get("expected_model", {}).get("automation_name")
    if not isinstance(name, str) or not name:
        raise WaterfallAdapterError("request name or expected_model automation_name is required")

    def cell(value: Any) -> Any:
        if value is None:
            return None
        if isinstance(value, dict):
            return copy.deepcopy(value)
        if isinstance(value, bool):
            return {"boolean": value}
        if _number(value):
            return {"number": value}
        return {"string": value}

    return [{"template": str(template), "data": [{"name": name, "table": [[cell(v) for v in row] for row in request["matrix"]]}]}]


def grade_waterfall(expected: dict, actual_model: dict, actual_semantics: dict) -> dict:
    """Grade stable model and visible waterfall invariants after native reopen."""
    expected_model = expected["expected_model"]
    for key in ("categories", "series_names", "series_values"):
        if actual_model.get(key) != expected_model.get(key):
            raise WaterfallAdapterError(f"native model {key} differs from expected")
    if "category_extents" in expected_model and actual_model.get("category_extents") != expected_model["category_extents"]:
        raise WaterfallAdapterError("native waterfall category extents differ")
    categories = expected_model["categories"]
    equals = actual_semantics.get("equals_slots", [])
    if len(equals) != 1 or equals[0][1] != len(categories) - 1:
        raise WaterfallAdapterError("native waterfall must retain one final equals category")
    connectors = actual_semantics.get("connectors", [])
    if len(connectors) != len(categories) - 1:
        raise WaterfallAdapterError("native waterfall connector count does not match adjacent categories")
    covered = sorted({endpoint[0] for connector in connectors for endpoint in connector})
    if covered != list(range(len(categories))):
        raise WaterfallAdapterError("native connectors do not cover every category")
    grounds = actual_semantics.get("grounds", [])
    if len(grounds) != len(categories) or any(entry[2] != "0" for entry in grounds):
        raise WaterfallAdapterError("native waterfall grounding is incomplete")
    return {"model_match": True, "category_count": len(categories), "equals_count": 1, "connector_count": len(connectors), "ground_count": len(grounds)}
