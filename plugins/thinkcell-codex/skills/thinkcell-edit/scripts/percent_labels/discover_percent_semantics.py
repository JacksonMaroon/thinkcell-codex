"""Discover native percentage labels by chart/category/series semantics.

This is a read-only scanner for research copies.  It intentionally resolves a
label from the owning data vector and model names instead of relying on the
old donor's scalar id or shape tag.
"""
from __future__ import annotations

import argparse
import json
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


def _scalar_locations(root):
    """Yield (scalar, category index, series index) from actual vector order."""
    for vector in root.xpath(".//CSequenceChartDataVector"):
        index = vector.find("m_nIndexInDataSheet")
        if index is None:
            continue
        category_index = int(index.get("val"))
        for series_index, elem in enumerate(vector.xpath("./ocol/elem")):
            scalar = next((n for n in root.iter() if n.get("id") == elem.get("idref")), None)
            if scalar is not None:
                yield scalar, category_index, series_index


def discover(path: str | Path) -> list[dict]:
    path = Path(path)
    raw = path.read_bytes()
    _, charts, _ = inventory(raw)
    result: list[dict] = []
    for chart in charts:
        doc = chart["doc"]
        model = model_of(chart)
        root = xml(streams(doc["ole"])[("think-cellXML",)])
        ids = {n.get("id"): n for n in root if n.get("id")}
        for scalar, category_index, series_index in _scalar_locations(root):
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
            result.append({
                "chart_part": doc["part"],
                "chart_model_id": model.get("id"),
                "chart_name": chart.get("name") or model.get("automation_name"),
                "slide_number": chart.get("slide_number"),
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
