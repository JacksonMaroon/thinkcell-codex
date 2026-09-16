"""Fail-closed mapping of ordinary native cache series to think-cell model rows.

Unlike the certified 100%-column path, ordinary bar/column caches must not be
assumed to use a reverse model order. This helper identifies each visible cache
series by its complete numeric vector and accepts a mapping only when that
vector matches exactly one model series.
"""
from __future__ import annotations

import math
from lxml import etree as E

NS = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}

def _need(ok, message):
    if not ok: raise ValueError(message)

def _equal(left, right):
    return len(left) == len(right) and all(
        isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool)
        and math.isclose(float(a), float(b), rel_tol=1e-10, abs_tol=1e-9)
        for a, b in zip(left, right)
    )

def cache_value_vectors(chart_xml):
    """Read ordered visible-series vectors from ordinary c:barChart cache XML."""
    root = E.fromstring(chart_xml) if isinstance(chart_xml, (bytes, str)) else chart_xml
    chart = root.find(".//c:barChart", NS); _need(chart is not None, "Expected native bar/column chart cache")
    values = []
    for series in chart.findall("c:ser", NS):
        points = series.findall("./c:val/c:numRef/c:numCache/c:pt", NS)
        _need(points, "Visible series has no numeric cache")
        indexed = {int(point.get("idx")): float(point.find("c:v", NS).text) for point in points}
        _need(set(indexed) == set(range(len(indexed))), "Visible series cache indices are sparse or duplicated")
        values.append([indexed[index] for index in range(len(indexed))])
    return values

def match_unique_series(cache_vectors, model_names, model_values):
    """Return visible-order names only when every numeric vector is unique."""
    _need(len(model_names) == len(model_values) and len(set(model_names)) == len(model_names), "Model series names are invalid")
    mapping = []
    used = set()
    for vector in cache_vectors:
        candidates = [index for index, values in enumerate(model_values) if _equal(vector, values)]
        _need(len(candidates) == 1, "Visible cache vector does not uniquely identify one model series")
        index = candidates[0]; _need(index not in used, "Two visible cache series map to one model series")
        used.add(index); mapping.append(model_names[index])
    _need(len(used) == len(model_names), "Model/visible series counts differ")
    return mapping
