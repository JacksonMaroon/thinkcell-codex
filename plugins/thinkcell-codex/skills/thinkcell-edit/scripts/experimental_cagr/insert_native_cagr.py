"""Prepare the proved native CAGR 1-to-2 insertion on a new local copy.

This is deliberately a narrow model-38764 adapter: one slide, one internal
ordinary sequence chart, one series, eight typed-date categories, and exactly
one existing CAGR connector from category index 1 to 7.  It never starts
Office.  `native_cagr_gate.py` performs official regeneration and native
readback on the prepared copy.
"""
from __future__ import annotations

import argparse
import base64
import copy
import hashlib
import io
import json
import subprocess
import sys
import tempfile
import uuid
import zipfile
from pathlib import Path

from lxml import etree as E

HERE = Path(__file__).resolve().parent
SCRIPTS = HERE.parent
IMPL = SCRIPTS / "thinkcell_no_click" / "implementation"
sys.path[:0] = [str(SCRIPTS), str(IMPL)]
from chart_geometry import streams, xml  # noqa: E402
from prepare_thinkcell_name import inventory as naming_inventory, link_contract  # noqa: E402

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"p": P, "a": A, "r": R}
REPLACE = IMPL / "replace_ole_stream.ps1"

# This is the complete CTextVariable formatter schema recovered from the
# successful 700 native proof.  It is embedded so no deck or research path is
# a runtime dependency.  Its cached numeric key is replaced per insertion.
NUMERIC_PERCENT_XML = """<CTextVariable><m_bstrFormat>'0''.''0''%'</m_bstrFormat><m_prec17834><m_bNumberIsYear val=\"0\"/><m_esigndisplay val=\"0\"/><m_nSignPosition val=\"-2147483648\"/><m_chMinusSymbol>-</m_chMinusSymbol><m_nDecimalDigits17909 val=\"1\"/><m_chDecimalSymbol17909>.</m_chDecimalSymbol17909><m_nGroupingDigits17909 val=\"3\"/><m_chGroupingSymbol17909>,</m_chGroupingSymbol17909><m_strPrefix/><m_strSuffix17909>%</m_strSuffix17909><m_nMagnitude17909 val=\"0\"/><m_yearfmt><begin val=\"0\"/><end val=\"4\"/></m_yearfmt></m_prec17834><m_bUseExcelFont val=\"0\"/><m_bUseExcelFontColor val=\"0\"/></CTextVariable>"""


def need(ok: bool, message: str) -> None:
    if not ok:
        raise RuntimeError(message)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def unique_tag() -> str:
    return "t" + base64.urlsafe_b64encode(uuid.uuid4().bytes).decode().rstrip("=")


def ids_of(root: E._Element) -> dict[str, E._Element]:
    result = {node.get("id"): node for node in root if node.get("id")}
    need(len(result) == sum(node.get("id") is not None for node in root), "duplicate native model IDs")
    return result


def one(parent: E._Element, tag: str) -> E._Element:
    nodes = parent.findall(tag, NS if ":" in tag else None)
    need(len(nodes) == 1, "missing or repeated " + tag)
    return nodes[0]


def version(root: E._Element) -> str:
    node = one(root, "version")
    return node.get("val") or ""


def source_graph(root: E._Element) -> tuple[E._Element, dict[str, E._Element], E._Element]:
    """Resolve the one existing complete connector by model semantics."""
    ids = ids_of(root)
    charts = [x for x in root if x.tag == "CSequenceChartSE"]
    need(len(charts) == 1 and version(root) == "38764", "requires model-38764 and one CSequenceChartSE")
    chart = charts[0]
    need(one(chart, "m_bConsistent").get("val") == "1" and one(chart, "m_ect").get("val") == "0", "current ordinary chart is inconsistent or unsupported")
    table = ids.get(one(chart, "m_dtable").get("idref"))
    need(table is not None and table.tag == "CSequenceChartDataTable", "missing active sequence datasource")
    need(one(table, "m_advisesink").get("idref") == "0", "external datasource is unsupported")
    need(one(table, "m_bNeedsUpdateFromSheetOnMakeTC").get("val") in (None, "0"), "stale datasource update wrapper is active")
    vectors = one(table, "ocol").findall("elem")
    series = one(table, "m_cscdser").findall("elem")
    members = one(table, "m_cscdveccon")
    need(len(vectors) == 8 and len(series) == 1 and len(members) == 1, "requires exactly 8 categories, 1 series, and 1 existing CAGR")
    graph = ids.get(members[0].get("idref"))
    need(graph is not None and graph.tag == "CSequenceChartDataVectorCAGR", "existing feature is not a native CAGR connector")
    for field in ("m_anchorSource", "m_anchorSink", "m_scdvecconlabel", "m_varsrcRelative", "m_varsrcAbsolute", "m_varsrcRelativeGroup", "m_varsrcCAGR"):
        need(graph.find(field) is not None, "CAGR graph closure missing " + field)
    source_anchor = ids.get(one(graph, "m_anchorSource").get("idref")); sink_anchor = ids.get(one(graph, "m_anchorSink").get("idref"))
    need(source_anchor is not None and sink_anchor is not None and source_anchor.tag == sink_anchor.tag == "CSequenceChartAnchor", "CAGR endpoint anchors unresolved")
    # Endpoint positions are resolved in the runner against both model and the
    # authoritative embedded workbook.  Here we require exactly one source
    # graph mounted on both current anchors.
    for anchor in (source_anchor, sink_anchor):
        refs = one(anchor, "m_cfeature").findall("elem")
        need([x.get("idref") for x in refs] == [graph.get("id")], "source anchor feature ownership is not singular")
    label = ids.get(one(graph, "m_scdvecconlabel").get("idref"))
    need(label is not None and label.tag == "CSequenceChartDataVectorConnectorLabel", "CAGR must own its ellipse label")
    line_refs = one(one(graph, "m_pptgenlineArrow"), "m_cpptline").findall("elem")
    need(len(line_refs) == 2, "CAGR must own exactly two arrow lines")
    lines = [ids.get(x.get("idref")) for x in line_refs]
    need(all(x is not None and x.tag == "CPPTLine" for x in lines), "CAGR arrow lines unresolved")
    return graph, ids, table


def no_other_chart_models(z: zipfile.ZipFile) -> None:
    """Allow decorative think-cell carriers, but never another active chart."""
    charts = []
    for part in [x for x in z.namelist() if x.startswith("ppt/embeddings/oleObject") and x.endswith(".bin")]:
        model = streams(z.read(part)).get(("think-cellXML",))
        if model is not None:
            root = xml(model)
            charts.extend((part, node.get("id")) for node in root if node.tag.endswith("ChartSE"))
    need(charts == [("ppt/embeddings/oleObject13.bin", "7")], "additional or noncanonical active chart carrier is unsupported")


def physical_tag_map(z: zipfile.ZipFile, rels: E._Element) -> dict[str, str]:
    found: dict[str, str] = {}
    for rel in rels:
        if not rel.get("Type", "").endswith("/tags"):
            continue
        part = "ppt/tags/" + rel.get("Target", "").rsplit("/", 1)[-1]
        need(part in z.namelist(), "shape tag relationship lacks its tag part")
        tags = E.fromstring(z.read(part)).findall("{%s}tag" % P)
        values = [x.get("val") for x in tags if x.get("name", "").upper() == "THINKCELLSHAPEDONOTDELETE"]
        need(len(values) == 1 and values[0], "invalid think-cell shape tag part")
        found[rel.get("Id")] = values[0]
    return found


def shape_signature(shape: E._Element) -> dict:
    geom = shape.find("p:spPr/a:prstGeom", NS)
    xfrm = shape.find("p:spPr/a:xfrm", NS)
    run = one(one(shape, ".//a:p"), "a:r")
    rpr = one(run, "a:rPr")
    return {"geometry": None if geom is None else geom.get("prst"),
            "bounds": {"x": xfrm.find("a:off", NS).get("x"), "y": xfrm.find("a:off", NS).get("y"), "cx": xfrm.find("a:ext", NS).get("cx"), "cy": xfrm.find("a:ext", NS).get("cy")} if xfrm is not None else None,
            "style": {k: rpr.get(k) for k in ("sz", "b", "i", "typeface")},
            "text": run.findtext("a:t", namespaces=NS)}


def tagged_shape(slide: E._Element, tags: dict[str, str], tag: str, kind: str) -> E._Element:
    path = ".//*[p:nvSpPr/p:cNvPr]" if kind == "sp" else ".//*[p:nvCxnSpPr/p:cNvPr]"
    matches = []
    for node in slide.xpath(path, namespaces=NS):
        refs = node.xpath(".//p:nvPr/p:custDataLst/p:tags/@r:id", namespaces=NS)
        if len(refs) == 1 and tags.get(refs[0]) == tag:
            matches.append(node)
    need(len(matches) == 1, "native tag does not resolve to one physical " + kind)
    return matches[0]


def literal_percent(value: float) -> tuple[str, str]:
    display = f"{value * 100:.1f}%"
    # DrawingML's datetime type contains the same quoted key as the model
    # formatter.  Keep every display character literal, including percent.
    return display, "'" + "''".join(display) + "'"


def replace_stream(carrier: bytes, payload: bytes, workdir: Path) -> bytes:
    need(REPLACE.is_file(), "required local OLE replacement helper is missing")
    with tempfile.TemporaryDirectory(dir=workdir) as tmp:
        tmp = Path(tmp); storage = tmp / "carrier.bin"; model = tmp / "model.xml"
        storage.write_bytes(carrier); model.write_bytes(payload)
        run = subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(REPLACE), "-StoragePath", str(storage), "-StreamBytesPath", str(model)], capture_output=True, text=True, timeout=90)
        need(run.returncode == 0, "OLE model replacement failed: " + run.stderr[-600:])
        return storage.read_bytes()


def build(input_path: Path, expected: str, output: Path, report: Path, workdir: Path) -> dict:
    source = input_path.resolve(); output = output.resolve(); report = report.resolve(); workdir = workdir.resolve()
    need(source.is_file() and source.suffix.lower() == ".pptx", "input must be an existing PPTX")
    need(workdir.is_dir() and output.parent == workdir and report.parent == workdir, "output and report must be direct children of the new workdir")
    need(source not in (output, report) and output != report and not output.exists() and not report.exists(), "input, output, and report must be distinct; outputs must be new")
    source_hash = sha(source); need(source_hash == expected.upper(), "input SHA-256 mismatch")
    raw = source.read_bytes()
    _, named_charts, _ = naming_inventory(raw)
    need(len(named_charts) == 1, "requires one internal named chart")
    link_contract(named_charts[0])
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        need(len(set(z.namelist())) == len(z.namelist()), "duplicate ZIP entries are unsupported")
        slides = [x for x in z.namelist() if x.startswith("ppt/slides/slide") and x.endswith(".xml") and "/_rels/" not in x]
        need(slides == ["ppt/slides/slide1.xml"], "requires exactly one logical slide")
        no_other_chart_models(z)
        carrier = z.read("ppt/embeddings/oleObject13.bin")
        slide = E.fromstring(z.read("ppt/slides/slide1.xml")); rels = E.fromstring(z.read("ppt/slides/_rels/slide1.xml.rels")); ct = E.fromstring(z.read("[Content_Types].xml"))
        root = xml(streams(carrier)[("think-cellXML",)])
        graph, ids, table = source_graph(root)
        old_label = ids[one(graph, "m_scdvecconlabel").get("idref")]
        old_tag = old_label.findtext(".//m_bstrShapeName")
        tag_map = physical_tag_map(z, rels)
        old_shape = tagged_shape(slide, tag_map, old_tag, "sp")
        original_cagr = ids[one(graph, "m_varsrcCAGR").get("idref")]
        baseline = {"source_feature_guid": one(original_cagr, "m_guid").get("val"), "label_tag": old_tag, "shape": shape_signature(old_shape), "cagr_text_reference_count": len(one(original_cagr, "m_ctextvar").findall("elem"))}

        next_id = max(int(i) for i in ids if i.isdigit()) + 1
        mapping: dict[str, str] = {graph.get("id"): str(next_id)}; next_id += 1
        for field in ("m_varsrcRelative", "m_varsrcAbsolute", "m_varsrcRelativeGroup", "m_varsrcCAGR", "m_scdvecconlabel"):
            mapping[one(graph, field).get("idref")] = str(next_id); next_id += 1
        line_refs = one(one(graph, "m_pptgenlineArrow"), "m_cpptline").findall("elem")
        for item in line_refs: mapping[item.get("idref")] = str(next_id); next_id += 1
        textvar_id = str(next_id)
        new_tag = unique_tag()

        clone = copy.deepcopy(graph); clone.set("id", mapping[graph.get("id")])
        for node in clone.iter():
            if node.get("idref") in mapping: node.set("idref", mapping[node.get("idref")])
        closure_ids = [one(graph, x).get("idref") for x in ("m_varsrcRelative", "m_varsrcAbsolute", "m_varsrcRelativeGroup", "m_varsrcCAGR", "m_scdvecconlabel")] + [x.get("idref") for x in line_refs]
        clones = []
        for old in closure_ids:
            node = copy.deepcopy(ids[old]); node.set("id", mapping[old])
            if node.tag == "CVariableSource":
                one(node, "m_guid").set("val", str(uuid.uuid4()))
            if node.tag == "CSequenceChartDataVectorConnectorLabel":
                one(node, ".//m_bstrShapeName").text = new_tag
            if node.tag == "CPPTLine":
                one(node, "m_bstrShapeName").text = unique_tag()
            clones.append(node)
        textvar = E.fromstring(NUMERIC_PERCENT_XML); textvar.set("id", textvar_id)
        new_cagr = next(x for x in clones if x.get("id") == mapping[one(graph, "m_varsrcCAGR").get("idref")])
        cagr_value = float(one(new_cagr, "m_varval").get("val"))
        cache, key = literal_percent(cagr_value)
        one(textvar, "m_bstrFormat").text = key
        E.SubElement(one(new_cagr, "m_ctextvar"), "elem", idref=textvar_id)
        for anchor_name in ("m_anchorSource", "m_anchorSink"):
            anchor = ids[one(graph, anchor_name).get("idref")]; members = one(anchor, "m_cfeature")
            E.SubElement(members, "elem", idref=clone.get("id")); members.set("length", str(len(members.findall("elem"))))
        members = one(table, "m_cscdveccon"); E.SubElement(members, "elem", idref=clone.get("id")); members.set("length", str(len(members.findall("elem"))))
        root.insert(root.index(graph) + 1, clone)
        for node in clones + [textvar]: root.insert(root.index(graph) + 1, node)
        changed_carrier = replace_stream(carrier, E.tostring(root, encoding="utf-8"), workdir)

        new_shape = copy.deepcopy(old_shape)
        used = [int(x.get("id")) for x in slide.xpath(".//p:cNvPr", namespaces=NS) if (x.get("id") or "").isdigit()]
        label_shape_id = str(max(used) + 1); cnv = one(new_shape, ".//p:cNvPr"); cnv.set("id", label_shape_id); cnv.set("name", "CAGR dynamic label " + label_shape_id)
        usedr = {x.get("Id") for x in rels}; serial = 1
        while "rId" + str(serial) in usedr: serial += 1
        rid = "rId" + str(serial); usedr.add(rid); tag_part = "cagr-dynamic-" + str(serial) + ".xml"
        one(new_shape, ".//p:nvPr/p:custDataLst/p:tags").set("{%s}id" % R, rid)
        paragraph = one(new_shape, ".//a:p"); oldrun = one(paragraph, "a:r"); rpr = copy.deepcopy(one(oldrun, "a:rPr"))
        for node in list(paragraph):
            if node.tag in {"{%s}r" % A, "{%s}fld" % A}: paragraph.remove(node)
        ppr = one(paragraph, "a:pPr")
        field = E.Element("{%s}fld" % A, id="{" + str(uuid.uuid4()).upper() + "}", type="datetime" + key); field.append(rpr); E.SubElement(field, "{%s}t" % A).text = cache
        paragraph.insert(paragraph.index(ppr) + 1, field)
        one(slide, ".//p:spTree").append(new_shape)
        E.SubElement(rels, "{%s}Relationship" % PR, Id=rid, Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags", Target="../tags/" + tag_part)
        E.SubElement(ct, "{%s}Override" % CT, PartName="/ppt/tags/" + tag_part, ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml")
        new_parts = {"ppt/tags/" + tag_part: (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:tagLst xmlns:p="{P}"><p:tag name="THINKCELLSHAPEDONOTDELETE" val="{new_tag}"/></p:tagLst>').encode()}
        # The two arrow-model lines are a physical/native closure.  Copy each
        # connector shape and its tag relationship alongside the copied model
        # line, with independent PowerPoint IDs and tags.
        for old_ref in line_refs:
            old_line = ids[old_ref.get("idref")]; old_line_tag = one(old_line, "m_bstrShapeName").text
            clone_line = next(x for x in clones if x.get("id") == mapping[old_ref.get("idref")]); clone_tag = one(clone_line, "m_bstrShapeName").text
            physical_line = copy.deepcopy(tagged_shape(slide, tag_map, old_line_tag, "cxnSp"))
            line_shape_id = str(max(used) + 1); used.append(int(line_shape_id)); c_nv = one(physical_line, ".//p:cNvPr"); c_nv.set("id", line_shape_id); c_nv.set("name", "CAGR arrow " + line_shape_id)
            while "rId" + str(serial) in usedr: serial += 1
            line_rid = "rId" + str(serial); usedr.add(line_rid); line_part = "cagr-arrow-" + str(serial) + ".xml"
            one(physical_line, ".//p:nvPr/p:custDataLst/p:tags").set("{%s}id" % R, line_rid)
            one(slide, ".//p:spTree").append(physical_line)
            E.SubElement(rels, "{%s}Relationship" % PR, Id=line_rid, Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags", Target="../tags/" + line_part)
            E.SubElement(ct, "{%s}Override" % CT, PartName="/ppt/tags/" + line_part, ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml")
            new_parts["ppt/tags/" + line_part] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:tagLst xmlns:p="{P}"><p:tag name="THINKCELLSHAPEDONOTDELETE" val="{clone_tag}"/></p:tagLst>').encode()
        modified = {"ppt/embeddings/oleObject13.bin": changed_carrier, "ppt/slides/slide1.xml": E.tostring(slide, xml_declaration=True, encoding="UTF-8", standalone=True), "ppt/slides/_rels/slide1.xml.rels": E.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True), "[Content_Types].xml": E.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True), **new_parts}
        with zipfile.ZipFile(output, "x") as dst:
            for item in z.infolist():
                if item.filename not in modified: dst.writestr(copy.copy(item), z.read(item.filename))
            for name, data in modified.items(): dst.writestr(name, data)
    result = {"status": "PREPARED_NATIVE_REGENERATION_REQUIRED", "source_sha256": source_hash, "candidate_sha256": sha(output), "profile": {"model_version": "38764", "categories": 8, "series": 1, "source_endpoint_index": 1, "sink_endpoint_index": 7}, "original": baseline, "inserted": {"feature_model_id": clone.get("id"), "label_model_id": mapping[old_label.get("id")], "text_variable_id": textvar_id, "physical_shape_id": label_shape_id, "field_id": field.get("id"), "tag": new_tag, "format": key, "source_feature_guid": one(new_cagr, "m_guid").get("val")}, "source_unchanged": sha(source) == source_hash}
    need(result["source_unchanged"], "source changed during preparation")
    report.write_text(json.dumps(result, indent=2), encoding="utf-8")
    return result


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--input", type=Path, required=True); p.add_argument("--expected-sha256", required=True)
    p.add_argument("--output", type=Path, required=True); p.add_argument("--report", type=Path, required=True); p.add_argument("--workdir", type=Path, required=True)
    a = p.parse_args(); print(json.dumps(build(a.input, a.expected_sha256, a.output, a.report, a.workdir), indent=2))
