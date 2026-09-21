"""Prepare a bounded native relative-label candidate on a local PPTX copy.

This module is deliberately a package-level experiment.  It performs no
Office automation: the only binary operation is replacing the embedded
think-cell model stream in a new local package.  Native regeneration and
visual review remain required after this preparer returns.
"""
from __future__ import annotations

import copy
import hashlib
import io
import math
import posixpath
import subprocess
import tempfile
import uuid
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from lxml import etree as E

HERE = Path(__file__).resolve().parent
try:
    from .portable_scripts import resolve_plugin_scripts
    from .discover_percent_semantics import discover, select
except ImportError:  # Direct script import from this folder.
    from portable_scripts import resolve_plugin_scripts  # type: ignore
    from discover_percent_semantics import discover, select  # type: ignore

SCRIPTS = resolve_plugin_scripts(HERE)
import sys
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "thinkcell_no_click" / "implementation"))
try:
    from chart_geometry import inventory, streams, xml
    from prepare_thinkcell_name import link_contract
    from runtime import powershell, powershell_env
except ImportError:  # pragma: no cover - the configured runtime supplies these.
    raise

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"p": P, "a": A, "c": C, "r": R}

SUPPORTED_MODEL_VERSIONS = {"38775", "38772"}
EXPECTED_DIMENSIONS = (3, 3)


def need(value: Any, message: str) -> None:
    if not value:
        raise ValueError(message)


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def format_percent(numerator: Any, denominator: Any, digits: int) -> str:
    value = (Decimal(str(numerator)) * Decimal("100") / Decimal(str(denominator))).quantize(
        Decimal("1." + "0" * digits), rounding=ROUND_HALF_UP)
    return f"{value}%"


def format_key(display: str) -> str:
    return "".join("'" + char + "'" for char in display)


def _new_id(ids: dict[str, E._Element]) -> str:
    numeric = [int(value) for value in ids if str(value).isdigit()]
    return str(max(numeric, default=0) + 1)


def _find_shape(package: zipfile.ZipFile, tag: str):
    for slide_name in package.namelist():
        if not (slide_name.startswith("ppt/slides/slide") and slide_name.endswith(".xml")):
            continue
        rel_name = posixpath.join(posixpath.dirname(slide_name), "_rels",
                                  posixpath.basename(slide_name) + ".rels")
        if rel_name not in package.namelist():
            continue
        slide = E.fromstring(package.read(slide_name))
        rels = E.fromstring(package.read(rel_name))
        by_id = {r.get("Id"): r for r in rels}
        for shape in slide.xpath(".//p:sp", namespaces=NS):
            for ref in shape.xpath(".//p:tags/@r:id", namespaces=NS):
                rel = by_id.get(ref)
                if rel is None:
                    continue
                target = posixpath.normpath(posixpath.join(posixpath.dirname(slide_name),
                                                          rel.get("Target")))
                if target.startswith("../"):
                    target = posixpath.normpath(posixpath.join(posixpath.dirname(slide_name), target))
                if target in package.namelist() and tag in package.read(target).decode("utf-8", errors="strict"):
                    return slide_name, slide, rel_name, rels, shape
    raise ValueError("series-label physical shape donor is unresolved")


def _model_version(root: E._Element) -> str | None:
    version = root.find("version")
    return version.get("val") if version is not None else None


def _select_row(rows: list[dict], *, category_index: int | None,
                series_index: int | None, category: str | None,
                series: str | None) -> dict:
    by_index = category_index is not None or series_index is not None
    by_name = category is not None or series is not None
    need(by_index != by_name, "Provide both category/series indices or both category/series names")
    if by_index:
        need(category_index is not None and series_index is not None,
             "category_index and series_index must be provided together")
        matches = [r for r in rows if r.get("category_index") == category_index
                   and r.get("series_index") == series_index]
    else:
        need(category is not None and series is not None,
             "category and series must be provided together")
        matches = [r for r in rows if str(r.get("category")) == str(category)
                   and str(r.get("series")) == str(series)]
    need(len(matches) == 1, "selector must resolve exactly one scalar label")
    return matches[0]


def _cached_label(chart_root: E._Element, row: dict) -> tuple[E._Element, str]:
    """Resolve one cached c:dLbl by series/category and verify its color."""
    bar_dir=chart_root.find('.//c:barChart/c:barDir',NS)
    grouping=chart_root.find('.//c:barChart/c:grouping',NS)
    need(bar_dir is not None and bar_dir.get('val')=='col' and grouping is not None
         and grouping.get('val')=='stacked','candidate requires the tested stacked-column cache profile')
    selected_value = float(row["numerator"])
    category_index = int(row["category_index"])
    series_index = int(row["series_index"])
    series_nodes = chart_root.xpath(".//c:barChart/c:ser", namespaces=NS)
    if len(series_nodes) != 3:
        raise ValueError("supported profile requires exactly three cached bar series")
    matches = []
    for ser in series_nodes:
        points = ser.xpath("./c:val//c:numCache/c:pt", namespaces=NS)
        for point in points:
            if int(point.get("idx", "-1")) != category_index:
                continue
            try:
                value = float(point.findtext("c:v", namespaces=NS))
            except (TypeError, ValueError):
                continue
            if math.isclose(value, selected_value, rel_tol=1e-8, abs_tol=1e-8):
                matches.append(ser)
    need(len(matches) == 1, "selected scalar does not resolve one cached value vector/category")
    ser = matches[0]
    # think-cell may reverse the native cache series order.  The unique value
    # vector/category match above is the identity bridge; do not assume cache
    # series index equals the model series index.
    labels = [label for label in ser.xpath("./c:dLbls/c:dLbl", namespaces=NS)
              if label.find("c:idx", NS).get("val") == str(category_index)]
    need(len(labels) == 1, "selected cached scalar does not resolve one dLbl")
    tx_pr = labels[0].find("c:txPr", NS)
    need(tx_pr is not None, "selected cached dLbl has no txPr font")
    colors = [node.get("val") for node in tx_pr.xpath(".//a:schemeClr", namespaces=NS)]
    need(colors and all(color == "bg1" for color in colors),
         "selected cached dLbl font color is outside the tested bg1 profile")
    return tx_pr, "bg1"


def _guard_tested_profile(model_root: E._Element, chart: dict, rows: list[dict]) -> None:
    """Reject model/axis/data profiles outside the tested absolute 3x3 case."""
    owner = model_root.find("CSequenceChartSE")
    need(owner is not None, "sequence chart owner is missing from model")
    consistent = owner.find("m_bConsistent")
    ect = owner.find("m_ect")
    default_mode = owner.find("m_estDefault")
    orientation=owner.find('m_eorient')
    secondary=owner.find('m_daxisSecondaryValue')
    need(orientation is not None and orientation.get('val')=='0'
         and secondary is not None and secondary.get('idref')=='0',
         'candidate requires a vertical chart without a secondary axis')
    need(consistent is not None and consistent.get("val") == "1",
         "candidate requires a consistent native chart model")
    need(ect is not None and default_mode is not None and ect.get("val") == "0"
         and default_mode.get("val") == "0",
         "candidate requires the ordinary tested chart mode")
    axis_ref = owner.find("m_daxisPrimaryValue")
    need(axis_ref is not None and axis_ref.get("idref") not in (None, "0"),
         "candidate requires an ordinary primary value axis")
    ids = {node.get("id"): node for node in model_root if node.get("id")}
    axis = ids.get(axis_ref.get("idref"))
    axis_type = axis.find("m_euseraxistype") if axis is not None else None
    need(axis is not None and axis_type is not None and axis_type.get("val") == "0",
         "candidate requires an ordinary value-axis type")
    min_node = axis.find("m_fMinValue")
    max_node = axis.find("m_fMaxValue")
    need(min_node is not None and max_node is not None,
         "candidate requires explicit primary-axis bounds")
    minimum = float(min_node.get("val"))
    maximum = float(max_node.get("val"))
    need(math.isfinite(minimum) and math.isfinite(maximum) and minimum >= 0 and maximum > minimum,
         "candidate requires a finite nonnegative absolute axis")
    breaks = axis.find("m_cdaxisbreak")
    need(breaks is not None and breaks.get("length") == "0" and not list(breaks),
         "candidate does not support an axis break")
    values: dict[int, dict[int, float]] = {}
    for row in rows:
        numerator = float(row["numerator"]); denominator = float(row["denominator"])
        need(math.isfinite(numerator) and numerator >= 0 and math.isfinite(denominator) and denominator > 0,
             "candidate requires finite nonnegative scalars and positive denominators")
        values.setdefault(int(row["category_index"]), {})[int(row["series_index"])] = numerator
    for category_index, series_values in values.items():
        need(len(series_values) == 3 and math.isclose(sum(series_values.values()),
                                                      float(next(r["denominator"] for r in rows
                                                                 if r["category_index"] == category_index)),
                                                      rel_tol=1e-8, abs_tol=1e-8),
             "candidate requires category extents equal the tested 3-series sums")


def _rewrite_colors(node: E._Element, color: str) -> None:
    for scheme in node.xpath(".//a:schemeClr", namespaces=NS):
        scheme.set("val", color)


def _prepare_shape(shape: E._Element, *, rect: E._Element, display: str,
                   format_code: str, cached_tx_pr: E._Element, shape_id: int) -> E._Element:
    clone = copy.deepcopy(shape)
    cnv = clone.find(".//p:cNvPr", NS)
    need(cnv is not None, "series-label donor has no cNvPr")
    cnv.set("id", str(shape_id))
    cnv.set("name", "Native percentage label candidate")
    for ext in cnv.findall("a:extLst", NS):
        cnv.remove(ext)
    xfrm = clone.find("p:spPr/a:xfrm", NS)
    need(xfrm is not None, "series-label donor has no transform")
    values = {key: round(float(rect.get(key)) * 1587.5)
              for key in ("left", "top", "right", "bottom")}
    xfrm.find("a:off", NS).attrib.update({"x": str(values["left"]), "y": str(values["top"])})
    xfrm.find("a:ext", NS).attrib.update({"cx": str(values["right"] - values["left"]),
                                           "cy": str(values["bottom"] - values["top"])})
    tx = clone.find("p:txBody", NS)
    need(tx is not None, "series-label donor has no text body")
    para = tx.find("a:p", NS)
    need(para is not None, "series-label donor has no paragraph")
    ppr = para.find("a:pPr", NS)
    for child in list(para):
        if child is not ppr:
            para.remove(child)
    if ppr is not None:
        ppr.set("algn", "ctr")
    cached_rpr = cached_tx_pr.find("a:p/a:pPr/a:defRPr", NS)
    need(cached_rpr is not None, "selected cached dLbl txPr has no default run font")
    rpr = copy.deepcopy(cached_rpr)
    rpr.tag = "{" + A + "}rPr"
    _rewrite_colors(rpr, "bg1")
    field = E.SubElement(para, "{" + A + "}fld",
                         id="{" + str(uuid.uuid4()).upper() + "}",
                         type="datetime" + format_code)
    field.append(rpr)
    E.SubElement(field, "{" + A + "}t").text = display
    _rewrite_colors(tx, "bg1")
    style = clone.find("p:style", NS)
    if style is not None:
        _rewrite_colors(style, "bg1")
    return clone


def _new_tag_part(entries: dict[str, bytes], rels: E._Element, shape_tag: str) -> tuple[str, str]:
    token = uuid.uuid4().hex
    part = f"ppt/tags/nativePercent-{token}.xml"
    rid_base = "rIdNativePercent"
    existing = {node.get("Id") for node in rels}
    rid = rid_base
    suffix = 2
    while rid in existing:
        rid = rid_base + str(suffix)
        suffix += 1
    tag_root = E.Element("{" + P + "}tagLst", nsmap={"p": P})
    E.SubElement(tag_root, "{" + P + "}tag", name="THINKCELLSHAPEDONOTDELETE", val=shape_tag)
    entries[part] = E.tostring(tag_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    rel_ns = rels.tag.split("}")[0] + "}" if rels.tag.startswith("{") else ""
    rel = E.SubElement(rels, rel_ns + "Relationship")
    rel.set("Id", rid)
    rel.set("Type", R + "/tags")
    rel.set("Target", "../tags/" + posixpath.basename(part))
    return part, rid


def prepare(source: Path, output: Path, report: Path, *, expected_sha: str,
            category_index: int | None = None, series_index: int | None = None,
            category: str | None = None, series: str | None = None,
            digits: int = 0, wrapper: str = "bare") -> dict:
    source = Path(source); output = Path(output); report = Path(report)
    need(digits == 0, "tested native-relative candidate supports zero decimal places only")
    need(wrapper == "bare", "tested native-relative candidate supports bare labels only")
    need(len({source.resolve(), output.resolve(), report.resolve()}) == 3,
         "source, output and report paths must differ")
    need(output.parent.is_dir() and report.parent.is_dir(),
         "output and report parent folders must already exist")
    need(not output.exists() and not report.exists(), "output and report must be new")
    raw = source.read_bytes()
    need(sha(raw) == expected_sha.upper(), "source SHA-256 mismatch")
    with zipfile.ZipFile(io.BytesIO(raw)) as zin:
        _, charts, _ = inventory(raw)
        need(len(charts) == 1, "candidate requires exactly one chart")
        chart = charts[0]
        need(chart["owner"].tag == "CSequenceChartSE" and chart.get("exact"),
             "candidate requires one exact native CSequenceChartSE chart")
        link = link_contract(chart)
        need(link.get("state") == "INTERNAL_DATASHEET_OBSERVED",
             "candidate requires an internal datasheet")
        model_part = chart["doc"]["streams"][("think-cellXML",)]
        model_root = xml(model_part)
        need(_model_version(model_root) in SUPPORTED_MODEL_VERSIONS,
             "candidate requires tested native model version 38775 or 38772")
        rows = discover(source)
        chart_rows = [r for r in rows if r.get("chart_part") == chart["doc"]["part"]]
        pairs = {(r.get("category_index"), r.get("series_index")) for r in chart_rows}
        need(pairs == {(cat, ser) for cat in range(3) for ser in range(3)},
             "candidate requires the tested fixed 3x3 scalar profile")
        _guard_tested_profile(model_root, chart, chart_rows)
        selected = _select_row(chart_rows, category_index=category_index,
                               series_index=series_index, category=category, series=series)
        need(selected.get("direct_rendered") and (selected.get("direct_precision") or {}).get("suffix") in (None, ""),
             "selected label must be an absolute directly rendered label")
        need(not selected.get("physical_shapes"), "selected label already has a physical field")
        need(selected.get("absolute_source_id") and selected.get("relative_source_id"),
             "selected scalar has no absolute and relative source bindings")
        need(selected.get("absolute_text_variable") is None and selected.get("relative_text_variable") is None,
             "selected scalar already has a text-variable binding")
        denominator = float(selected["denominator"])
        need(math.isfinite(denominator) and denominator != 0, "selected label denominator must be finite and nonzero")
        ratio = float(selected["numerator"]) / denominator
        need(math.isfinite(ratio), "selected label ratio must be finite")
        display = format_percent(selected["numerator"], selected["denominator"], digits)
        format_code = format_key(display)
        ids = {n.get("id"): n for n in model_root if n.get("id")}
        relative = ids.get(selected["relative_source_id"])
        label = ids.get(selected["label_id"])
        need(relative is not None and label is not None, "selected model bindings are unresolved")
        text_ref_parent = relative.find("m_ctextvar")
        need(text_ref_parent is not None and text_ref_parent.find("elem") is None,
             "selected relative source already has a text-variable binding")
        owner = model_root.find("CSequenceChartSE")
        need(owner is not None, "sequence chart owner is missing from model")
        donor_labels = model_root.findall("CSequenceChartDataSeriesLabel")
        need(len(donor_labels) == 3, "candidate requires three native series-label donors")
        donor = donor_labels[int(selected["series_index"])]
        donor_tag = donor.findtext("m_ppttb/m_bstrShapeName")
        need(donor_tag, "native series-label donor has no physical shape tag")
        rect = label.find("m_rectPosition[@reqver='36208']")
        if rect is None:
            rect = label.find("m_rectPosition")
        need(rect is not None, "selected scalar label has no geometry")
        frame = chart["frames"][0]
        chart_part = frame["native_chart_part"]
        chart_root = xml(zin.read(chart_part))
        cached_tx_pr, source_color = _cached_label(chart_root, selected)
        slide_name, slide_root, rel_name, rels_root, donor_shape = _find_shape(zin, donor_tag)
        entries = {info.filename: zin.read(info.filename) for info in zin.infolist()}
        infos = {info.filename: info for info in zin.infolist()}
    # Bind the model's existing relative source to a new native percentage text variable.
    text_id = _new_id(ids)
    textvar = E.Element("CTextVariable", id=text_id)
    E.SubElement(textvar, "m_bstrFormat").text = format_code
    relative_precision = copy.deepcopy(owner.find("m_precRelativeScalar"))
    need(relative_precision is not None, "relative scalar precision is missing")
    relative_precision.tag = "m_prec17834"
    textvar.append(relative_precision)
    E.SubElement(textvar, "m_bUseExcelFont", val="0")
    E.SubElement(textvar, "m_bUseExcelFontColor", val="0")
    model_root.append(textvar)
    E.SubElement(text_ref_parent, "elem", idref=text_id)
    rendering = label.find("m_bMSGraphRendering")
    need(rendering is not None, "selected scalar label rendering flag is missing")
    rendering.set("val", "0")
    old_tb = label.find("m_ppttb")
    need(old_tb is not None, "selected scalar label text box is missing")
    donor_tb = copy.deepcopy(donor.find("m_ppttb"))
    need(donor_tb is not None, "series-label donor text box is missing")
    donor_tb.find("m_bstrShapeName").text = selected["shape_tag"]
    for key, value in rect.attrib.items():
        if key != "reqver":
            donor_tb.find("m_rectPPTShape").set(key, str(round(float(value))))
    donor_tb.find("m_bPlaced").set("val", "1")
    donor_tb.find("m_ppparaalign").set("val", "2")
    donor_tb.find("m_msoverticalanchor").set("val", "3")
    donor_tb.replace(donor_tb.find("m_colFont"), copy.deepcopy(old_tb.find("m_colFont")))
    label.replace(old_tb, donor_tb)
    candidate_shape = _prepare_shape(donor_shape, rect=rect, display=display,
                                     format_code=format_code, cached_tx_pr=cached_tx_pr,
                                     shape_id=max(int(node.get("id")) for node in slide_root.findall(".//p:cNvPr", NS)) + 1)
    tags_ref = candidate_shape.find(".//p:tags", NS)
    need(tags_ref is not None, "series-label donor has no tags relationship")
    tag_part, rel_id = _new_tag_part(entries, rels_root, selected["shape_tag"])
    tags_ref.set("{" + R + "}id", rel_id)
    slide_root.find("p:cSld/p:spTree", NS).append(candidate_shape)
    entries[slide_name] = E.tostring(slide_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    entries[rel_name] = E.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    types_root = E.fromstring(entries["[Content_Types].xml"])
    override = E.SubElement(types_root, "{" + CT + "}Override")
    override.set("PartName", "/" + tag_part)
    override.set("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.tags+xml")
    entries["[Content_Types].xml"] = E.tostring(types_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    model_after = E.tostring(model_root, encoding="utf-8")
    with tempfile.TemporaryDirectory(prefix="native_relative_candidate_", dir=str(output.parent)) as tmp:
        tmp_path = Path(tmp)
        carrier = tmp_path / "carrier.bin"
        payload = tmp_path / "model.xml"
        carrier.write_bytes(chart["doc"]["ole"])
        payload.write_bytes(model_after)
        command = [powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File",
                   str(SCRIPTS / "thinkcell_no_click" / "implementation" / "replace_ole_stream.ps1"),
                   "-StoragePath", str(carrier), "-StreamBytesPath", str(payload)]
        proc = subprocess.run(command, capture_output=True, text=True,
                              env=powershell_env(), timeout=60)
        need(proc.returncode == 0, "OLE model patch failed: " + proc.stderr[-400:])
        entries[chart["doc"]["part"]] = carrier.read_bytes()
    with zipfile.ZipFile(output, "x") as zout:
        for name, data in entries.items():
            if name in infos:
                zout.writestr(copy.copy(infos[name]), data)
            else:
                zout.writestr(name, data)
    with zipfile.ZipFile(io.BytesIO(raw)) as before_zip, zipfile.ZipFile(output) as after_zip:
        before_names = set(before_zip.namelist())
        after_names = set(after_zip.namelist())
        changed = [name for name in sorted(before_names | after_names)
                   if name not in before_names or name not in after_names
                   or before_zip.read(name) != after_zip.read(name)]
    expected_changed = {chart["doc"]["part"], slide_name, rel_name, "[Content_Types].xml", tag_part}
    need(set(changed) == expected_changed, "unexpected package changes: " + str(changed))
    result = {
        "status": "PREPARED_NATIVE_RELATIVE_CANDIDATE_NATIVE_PROOF_REQUIRED",
        "source_sha256": expected_sha.upper(), "output_sha256": sha(output.read_bytes()),
        "selector": {"category_index": selected["category_index"], "series_index": selected["series_index"],
                     "category": selected["category"], "series": selected["series"],
                     "chart_part": selected["chart_part"]},
        "selected": selected, "model_version": _model_version(model_root),
        "requested_precision": digits, "expected_display": display, "ratio": ratio,
        "format_key": format_code, "relative_source_id": selected["relative_source_id"],
        "physical_field_id": candidate_shape.find(".//a:fld", NS).get("id"),
        "physical_field_id_retained": None, "physical_shape_tag": selected["shape_tag"],
        "source_cached_label_color": source_color, "geometry_unchanged": True,
        "changed_entries": changed,
        "native_gates": ["official regeneration", "native save/reopen", "changed numerator",
                         "changed denominator", "visual wrapper review"],
    }
    report.write_text(__import__("json").dumps(result, indent=2), encoding="utf-8")
    return result


if __name__ == "__main__":
    import argparse
    import json
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--expected-sha256", required=True)
    parser.add_argument("--category-index", type=int)
    parser.add_argument("--series-index", type=int)
    parser.add_argument("--category")
    parser.add_argument("--series")
    parser.add_argument("--digits", type=int, default=0)
    parser.add_argument("--wrapper", choices=["bare"], default="bare")
    args = parser.parse_args()
    print(json.dumps(prepare(args.input, args.output, args.report,
                             expected_sha=args.expected_sha256,
                             category_index=args.category_index, series_index=args.series_index,
                             category=args.category, series=args.series,
                             digits=args.digits, wrapper=args.wrapper), indent=2))
