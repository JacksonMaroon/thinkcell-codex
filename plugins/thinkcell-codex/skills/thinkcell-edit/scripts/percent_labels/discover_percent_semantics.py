"""Discover native percentage labels by chart/category/series semantics.

This is a read-only scanner for research copies.  It intentionally resolves a
label from the owning data vector and model names instead of relying on the
old donor's scalar id or shape tag.
"""
from __future__ import annotations

import argparse
import itertools
import json
import math
import sys
import zipfile
from pathlib import Path

from lxml import etree as E

sys.path.insert(0, str(Path(__file__).resolve().parent))
from portable_scripts import resolve_plugin_scripts  # noqa: E402
SCRIPTS = resolve_plugin_scripts(Path(__file__).resolve().parent)
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "thinkcell_no_click" / "implementation"))
from chart_geometry import inventory, streams, xml  # noqa: E402
from multi_chart_update import cells_matrix, model_of  # noqa: E402
from audit_thinkcell_integrity import NS, logical_slides, relationship_map  # noqa: E402


def _field_shapes(raw: bytes, shape_tag: str) -> list[dict]:
    found: list[dict] = []
    with zipfile.ZipFile(__import__("io").BytesIO(raw)) as package:
        for slide in logical_slides(package):
            root = xml(package.read(slide["part"]))
            rels = relationship_map(package, slide["part"])
            for shape in root.findall(".//p:sp", NS):
                tags = []
                for ref in shape.findall(".//p:tags", NS):
                    rel = rels.get(ref.get("{" + NS["r"] + "}id"))
                    if not rel:
                        continue
                    try:
                        tags.extend(n.get("val") for n in xml(package.read(rel["resolved"]))
                                    if n.get("name", "").upper() == "THINKCELLSHAPEDONOTDELETE")
                    except KeyError:
                        continue
                if shape_tag not in tags:
                    continue
                found.append({
                    "slide_id": slide["id"],
                    "shape_id": shape.find(".//p:cNvPr", NS).get("id"),
                    "shape_name": shape.find(".//p:cNvPr", NS).get("name"),
                    "fields": [{"id": f.get("id"), "type": f.get("type"),
                                "text": f.findtext("a:t", namespaces=NS)}
                               for f in shape.findall(".//a:fld", NS)],
                    "literal_texts": [n.text for n in shape.findall(".//a:t", NS)],
                })
    return found


def _scalar_locations(root, table_id=None):
    """Yield (scalar, category index, series index) from actual vector order."""
    vector_ids = None
    if table_id is not None:
        table = next((node for node in root.iter("CSequenceChartDataTable")
                      if node.get("id") == str(table_id)), None)
        if table is None:
            return
        vector_ids = {ref.get("idref") for ref in table.findall("./ocol/elem")}
    for vector in root.xpath(".//CSequenceChartDataVector"):
        if vector_ids is not None and vector.get("id") not in vector_ids:
            continue
        index = vector.find("m_nIndexInDataSheet")
        if index is None:
            continue
        category_index = int(index.get("val"))
        if category_index < 0:
            continue
        for series_index, elem in enumerate(vector.xpath("./ocol/elem")):
            scalar = next((n for n in root.iter() if n.get("id") == elem.get("idref")), None)
            if scalar is not None:
                yield scalar, category_index, series_index


def _node_text(node, path):
    found = node.find(path) if node is not None else None
    value = (found.text if found is not None else None)
    if value is None and found is not None:
        value = found.get("val")
    return value if value is None else value.strip()


def _number(node, path):
    value = _node_text(node, path)
    try:
        return float(value) if value is not None else None
    except (TypeError, ValueError):
        return None


def _direct_precision(label):
    """Return the label's own native precision, separate from field precision."""
    precision = label.find("m_prec") if label is not None else None
    digits = precision.find("m_nDecimalDigits17909") if precision is not None else None
    return {
        "source": "m_prec" if precision is not None else None,
        "prefix": _node_text(precision, "m_strPrefix"),
        "suffix": _node_text(precision, "m_strSuffix17909"),
        "decimal_digits": digits.get("val") if digits is not None else None,
        "present": precision is not None,
    }


def _native_percent_axis(chart, ids):
    """Recognize the native percentage axis without relying on cache formatting."""
    reference = chart["owner"].find("m_daxisPrimaryValue")
    axis = ids.get(reference.get("idref")) if reference is not None else None
    if axis is None or axis.tag != "CSequenceChartDataAxis":
        return {"verified": False, "reason": "missing_native_value_axis"}
    explicit = axis.find("m_bPercentage")
    if explicit is not None and explicit.get("val") is not None:
        value = explicit.get("val")
        return {"verified": value in {"1", "true", "True"},
                "source": "m_bPercentage", "axis_id": axis.get("id")}
    min_value = _number(axis, "m_fMinValue")
    max_value = _number(axis, "m_fMaxValue")
    axis_type = _node_text(axis, "m_edaxistype")
    user_precision = axis.find("m_precUser")
    suffix = _node_text(user_precision, "m_strSuffix17909")
    verified = (min_value == 0.0 and max_value == 1.0 and axis_type == "1"
                and suffix == "%")
    return {
        "verified": verified,
        "source": "normalized_native_axis_signature",
        "axis_id": axis.get("id"),
        "min": min_value,
        "max": max_value,
        "axis_type": axis_type,
        "suffix": suffix,
    }


def _chart_cache(package, chart):
    """Read the chart cache and stable c:idx/c:order/formula identities."""
    if not chart.get("frames") or len(chart["frames"]) != 1:
        return {"verified": False, "reason": "ambiguous_chart_frame", "series": []}
    native_part = chart["frames"][0].get("native_chart_part")
    if not native_part or native_part not in package.namelist():
        return {"verified": False, "reason": "missing_native_chart_part", "series": []}
    root = xml(package.read(native_part))
    series = []
    for index, item in enumerate(root.xpath(".//c:ser", namespaces={"c": NS["c"]})):
        reference = item.find("c:val/c:numRef", {"c": NS["c"]})
        cache = reference.find("c:numCache", {"c": NS["c"]}) if reference is not None else None
        if cache is None:
            series.append({"chart_index": index, "idx": item.xpath("string(./c:idx/@val)", namespaces={"c": NS["c"]}) or None,
                           "order": item.xpath("string(./c:order/@val)", namespaces={"c": NS["c"]}) or None,
                           "formula": reference.findtext("c:f", namespaces={"c": NS["c"]}) if reference is not None else None,
                           "values": {}})
            continue
        values = {}
        for point in cache.findall("c:pt", {"c": NS["c"]}):
            try:
                values[int(point.get("idx"))] = float(point.findtext("c:v", namespaces={"c": NS["c"]}))
            except (TypeError, ValueError):
                continue
        series.append({
            "chart_index": index,
            "idx": item.xpath("string(./c:idx/@val)", namespaces={"c": NS["c"]}) or None,
            "order": item.xpath("string(./c:order/@val)", namespaces={"c": NS["c"]}) or None,
            "formula": reference.findtext("c:f", namespaces={"c": NS["c"]}) if reference is not None else None,
            "values": values,
        })
    return {"verified": bool(series), "native_part": native_part, "series": series}


def _cache_matrix_agreement(model, cache):
    """Resolve cache row/column identity by a unique numerical bijection."""
    values = model.get("series_values") or []
    extents = model.get("category_extents") or []
    nseries, ncats = len(values), len(model.get("categories") or [])
    if not nseries or not ncats or len(cache.get("series", [])) != nseries:
        return {"verified": False, "reason": "cache_dimensions_ambiguous"}
    expected = []
    for row in values:
        expected.append([
            (100.0 * row[ci] / extents[ci])
            if ci < len(row) and ci < len(extents) and row[ci] is not None
            and extents[ci] not in (None, 0) else None
            for ci in range(ncats)
        ])
    observed = [item.get("values", {}) for item in cache["series"]]
    if any(set(row) != set(range(ncats)) for row in observed):
        return {"verified": False, "reason": "cache_points_incomplete", "expected": expected}

    def matches(permutation):
        # permutation maps chart cache row/column to native model row/column.
        for chart_row, model_row in enumerate(permutation[0]):
            for chart_col, model_col in enumerate(permutation[1]):
                actual = observed[chart_row].get(chart_col)
                wanted = expected[model_row][model_col]
                if wanted is None or actual is None or not math.isclose(actual, wanted, rel_tol=1e-6, abs_tol=1e-6):
                    return False
        return True

    # The corpus is small, but keep this explicitly bounded for arbitrary decks.
    mapping_budget = 100_000
    if nseries > 8 or ncats > 8 or (math.factorial(nseries) * math.factorial(ncats)) > mapping_budget:
        return {"verified": False, "reason": "mapping_bound_exceeded", "expected": expected}
    candidates = []
    for series_permutation in itertools.permutations(range(nseries)):
        for category_permutation in itertools.permutations(range(ncats)):
            pair = (series_permutation, category_permutation)
            if matches(pair):
                candidates.append(pair)
                if len(candidates) > 1:
                    return {"verified": False, "reason": "mapping_ambiguous", "expected": expected,
                            "candidate_count": len(candidates)}
    if len(candidates) != 1:
        return {"verified": False, "reason": "mapping_not_found", "expected": expected,
                "candidate_count": len(candidates)}
    series_map, category_map = candidates[0]
    return {
        "verified": True,
        "series_cache_to_model": list(series_map),
        "category_cache_to_model": list(category_map),
        "expected": expected,
        "observed": [[observed[r][c] for c in range(ncats)] for r in range(nseries)],
        "candidate_count": 1,
    }


def discover(path: str | Path) -> list[dict]:
    path = Path(path)
    raw = path.read_bytes()
    _, charts, _ = inventory(raw)
    result: list[dict] = []
    with zipfile.ZipFile(__import__("io").BytesIO(raw)) as package:
      for chart in charts:
        doc = chart["doc"]
        model = model_of(chart)
        root = xml(streams(doc["ole"])[("think-cellXML",)])
        ids = {n.get("id"): n for n in root if n.get("id")}
        # Resolve the chart-local graph.  Iterating every vector in a shared
        # model would silently attribute a sibling chart's labels here.
        locations = list(_scalar_locations(root, chart["table"].get("id")))
        axis = _native_percent_axis(chart, ids)
        cache = _chart_cache(package, chart)
        agreement = _cache_matrix_agreement(model, cache) if cache.get("verified") else {
            "verified": False, "reason": cache.get("reason", "cache_unavailable")
        }
        owner_id = chart["owner"].get("id")
        table_id = chart["table"].get("id")
        frame = chart["frames"][0] if len(chart.get("frames", [])) == 1 else None
        chart_identity = {
            "carrier": doc["part"],
            "owner_id": owner_id,
            "owner_type": chart["owner"].tag,
            "table_id": table_id,
            "shape_id": frame.get("shape_id") if frame else None,
            "shape_tag": frame.get("shape_tag") if frame else None,
            "native_chart_part": frame.get("native_chart_part") if frame else None,
            "automation_name": chart.get("owner_name") or None,
            "exact": bool(chart.get("exact")),
        }
        for scalar, category_index, series_index in locations:
            label_ref = scalar.find("m_scdlabel")
            if label_ref is None or label_ref.get("idref") in (None, "0"):
                continue
            if category_index >= len(model["categories"]) or series_index >= len(model["series_names"]):
                continue
            label = ids.get(label_ref.get("idref"))
            if label is None:
                continue
            tag = label.findtext("m_ppttb/m_bstrShapeName")
            if not tag:
                continue
            abs_ref = scalar.find("m_varsrcAbsolute")
            rel_ref = scalar.find("m_varsrcRelative")
            absolute = ids.get(abs_ref.get("idref")) if abs_ref is not None else None
            relative = ids.get(rel_ref.get("idref")) if rel_ref is not None else None
            abs_text = ids.get((absolute.find("m_ctextvar/elem").get("idref")
                                if absolute is not None and absolute.find("m_ctextvar/elem") is not None else ""))
            rel_text = ids.get((relative.find("m_ctextvar/elem").get("idref")
                                if relative is not None and relative.find("m_ctextvar/elem") is not None else ""))
            fields = _field_shapes(raw, tag)
            precision = _direct_precision(label)
            rendering = _node_text(label, "m_bMSGraphRendering")
            direct_rendered = bool(precision["present"] and rendering in {"1", "true", "True"})
            direct_percent_shape = bool(direct_rendered and precision["prefix"] in (None, "")
                                       and precision["suffix"] == "%"
                                       and precision["decimal_digits"] is not None)
            mapping_verified = bool(agreement.get("verified"))
            cache_series_index = None
            cache_category_index = None
            cache_value = None
            if mapping_verified:
                try:
                    cache_series_index = agreement["series_cache_to_model"].index(series_index)
                    cache_category_index = agreement["category_cache_to_model"].index(category_index)
                    cache_value = agreement["observed"][cache_series_index][cache_category_index]
                except (ValueError, IndexError):
                    mapping_verified = False
            numerical = None
            if cache_value is not None and model["category_extents"][category_index] not in (None, 0):
                numerical = 100.0 * model["series_values"][series_index][category_index] / model["category_extents"][category_index]
            verified = bool(
                direct_percent_shape
                and chart_identity["exact"]
                and chart["owner"].tag == "CSequenceChartSE"
                and axis.get("verified")
                and mapping_verified
                and numerical is not None
                and math.isclose(cache_value, numerical, rel_tol=1e-6, abs_tol=1e-6)
            )
            result.append({
                "chart_part": doc["part"],
                "chart_model_id": model.get("id"),
                "chart_name": chart.get("owner_name") or model.get("automation_name"),
                "slide_number": doc.get("slide_number"),
                "category": model["categories"][category_index],
                "category_index": category_index,
                "series": model["series_names"][series_index],
                "series_index": series_index,
                "numerator": model["series_values"][series_index][category_index],
                "denominator": model["category_extents"][category_index],
                "scalar_id": scalar.get("id"),
                "label_id": label_ref.get("idref"),
                "shape_tag": tag,
                "absolute_source_id": abs_ref.get("idref") if abs_ref is not None else None,
                "relative_source_id": rel_ref.get("idref") if rel_ref is not None else None,
                "absolute_text_variable": abs_text.get("id") if abs_text is not None else None,
                "relative_text_variable": rel_text.get("id") if rel_text is not None else None,
                "relative_suffix": rel_text.findtext("m_prec17834/m_strSuffix17909") if rel_text is not None else None,
                "relative_decimal_digits": (rel_text.find("m_prec17834/m_nDecimalDigits17909").get("val")
                                             if rel_text is not None and rel_text.find("m_prec17834/m_nDecimalDigits17909") is not None else None),
                "physical_shapes": fields,
                "owning_chart": chart_identity,
                "direct_precision": precision,
                "direct_rendered": direct_rendered,
                "native_semantics": axis,
                "cache": {
                    "native_chart_part": cache.get("native_part"),
                    "series_index": cache_series_index,
                    "category_index": cache_category_index,
                    "value": cache_value,
                    "series_identifier": (cache.get("series", [])[cache_series_index]
                                           if cache_series_index is not None and cache_series_index < len(cache.get("series", [])) else None),
                },
                "cache_agreement": {
                    "verified": bool(agreement.get("verified") and mapping_verified),
                    "mapping_reason": agreement.get("reason"),
                    "series_cache_to_model": agreement.get("series_cache_to_model"),
                    "category_cache_to_model": agreement.get("category_cache_to_model"),
                    "expected": numerical,
                    "observed": cache_value,
                },
                "classification": "verified_direct_percentage" if verified else "unverified",
                "verified_direct_percentage": verified,
            })
    return result


def select(rows: list[dict], *, category: str, series: str, chart_part: str | None = None,
           chart_name: str | None = None) -> dict:
    matches = [r for r in rows if r["category"] == category and r["series"] == series
               and (chart_part is None or r["chart_part"] == chart_part)
               and (chart_name is None or r["chart_name"] == chart_name)]
    if chart_part is None and chart_name is None:
        charts = {(r["chart_part"], r["chart_name"]) for r in matches}
        if len(charts) > 1:
            raise ValueError("semantic selector spans multiple charts; provide chart_part or chart_name")
    if len(matches) != 1:
        raise ValueError(f"semantic selector resolved {len(matches)} labels; expected one")
    selected = matches[0]
    if len(selected["physical_shapes"]) != 1:
        raise ValueError("selected label tag does not resolve exactly one physical shape")
    if selected["relative_source_id"] is None or selected["relative_text_variable"] is None:
        raise ValueError("selected semantic label has no active relative field")
    if selected["relative_suffix"] != "%" or selected["relative_decimal_digits"] is None:
        raise ValueError("selected semantic label is not a percent field with active precision")
    return selected


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--input", required=True)
    p.add_argument("--category")
    p.add_argument("--series")
    p.add_argument("--chart-part")
    p.add_argument("--chart-name")
    p.add_argument("--out")
    a = p.parse_args()
    rows = discover(a.input)
    value = (select(rows, category=a.category, series=a.series, chart_part=a.chart_part,
                    chart_name=a.chart_name)
             if a.category and a.series else rows)
    payload = json.dumps(value, indent=2)
    if a.out:
        Path(a.out).write_text(payload, encoding="utf-8")
    print(payload)
