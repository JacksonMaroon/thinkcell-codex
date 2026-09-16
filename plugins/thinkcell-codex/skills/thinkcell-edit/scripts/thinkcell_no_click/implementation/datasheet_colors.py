"""Strict, public-JSON datasheet fill decorator.

The caller supplies the normal, scalar think-cell request unchanged.  This
module adds only documented ``fill`` attributes to the JSON table that is sent
to ppttc; it never changes the underlying values, categories, or labels.
"""
from __future__ import annotations

import copy
import re

HEX = re.compile(r"#[0-9A-Fa-f]{6}$")


def need(condition, message):
    if not condition:
        raise ValueError(message)


def _color(value):
    need(isinstance(value, str) and HEX.fullmatch(value), "Colors must be #RRGGBB")
    return value.upper()


def validate(appearance, request, contract):
    """Validate a donor-aware color plan without mutating the input request."""
    need(isinstance(appearance, dict) and set(appearance) == {"schema", "series_fills", "point_fills"},
         "Appearance must contain exactly schema, series_fills, and point_fills")
    need(appearance["schema"] == "tc.datasheet-fill.v1", "Unsupported appearance schema")
    need(contract.get("family") in {"sequence", "pie", "scatter", "bubble", "mekko-percent", "mekko-units"},
         "Datasheet fills require a supported chart family")
    series = appearance["series_fills"]
    points = appearance["point_fills"]
    need(isinstance(series, dict) and isinstance(points, dict), "Fill mappings must be objects")
    if "series_names" not in request["expected_model"]:
        need(not series and not points, "This chart has no named series for a fill mapping")
        return
    names = request["expected_model"]["series_names"]
    need(set(series) <= set(names) and set(points) <= set(names), "Unknown series in fill mapping")
    count = len(request["expected_model"]["categories"])
    for name, color in series.items():
        _color(color)
    for name, colors in points.items():
        need(isinstance(colors, list) and len(colors) == count, "Point fills must match the category count")
        for color in colors:
            if color is not None:
                _color(color)


def decorate_table(request, contract, appearance):
    """Return a typed JSON table with public fill attributes.

    Series fills are applied to all value cells.  Point fills override the
    matching value cell.  For line/area donors, callers must use a donor whose
    Color Scheme uses datasheet fills and whose applicable series-label cells
    are configured to inherit them; JSON has no switch for that setting.
    """
    validate(appearance, request, contract)
    table = copy.deepcopy(request["matrix"])
    names = request["expected_model"].get("series_names", [])
    for row in table:
        if not row or row[0] not in names:
            continue
        name = row[0]
        for idx in range(1, len(row)):
            value = row[idx]
            color = appearance["point_fills"].get(name, [None] * (len(row) - 1))[idx - 1]
            color = color or appearance["series_fills"].get(name)
            if color is not None and value is not None:
                row[idx] = {"number": value, "fill": _color(color)}
    return table
