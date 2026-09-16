"""Prepare a semantic, endpoint-selectable native CAGR insertion candidate.

This is an offline model/package adapter.  It does not start Office or ppttc.
The companion runner supplied by the parent agent performs official JSON
regeneration and native reopen/readback under the Office lease.

The original v5 adapter was intentionally bounded to model 38764, carrier
oleObject13, one series, eight categories, and anchors 1 -> 7.  This adapter
discovers a CSequenceChartSE by its native name, resolves its active carrier
and table, and creates fresh endpoint anchors for any two scalar category
indices.  It preserves the chart's own label/arrow geometry and model
closures so think-cell remains the owner of the inserted annotation.
"""
from __future__ import annotations

import argparse
import base64
import copy
import hashlib
import io
import json
import sys
import subprocess
import tempfile
import uuid
import zipfile
from pathlib import Path

from lxml import etree as E

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"p": P, "a": A, "r": R}


def need(ok: bool, message: str) -> None:
    if not ok:
        raise RuntimeError(message)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def unique_tag(prefix: str = "t") -> str:
    return prefix + base64.urlsafe_b64encode(uuid.uuid4().bytes).decode().rstrip("=")


def init_helpers(skill_dir: Path):
    scripts = skill_dir / "scripts"
    impl = scripts / "thinkcell_no_click" / "implementation"
    sys.path[:0] = [str(scripts), str(impl)]
    from chart_geometry import streams as read_streams, xml  # type: ignore
    from prepare_thinkcell_name import inventory as naming_inventory, link_contract  # type: ignore
    helper = impl / "replace_ole_stream.ps1"
    def replace_ole(carrier: bytes, payload: bytes, workdir: Path) -> bytes:
        with tempfile.TemporaryDirectory(dir=workdir) as tmp:
            tmp_path = Path(tmp); storage = tmp_path / "carrier.bin"; model = tmp_path / "model.xml"
            storage.write_bytes(carrier); model.write_bytes(payload)
            run = subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(helper), "-StoragePath", str(storage), "-StreamBytesPath", str(model)], capture_output=True, text=True, timeout=90)
            need(run.returncode == 0, "OLE model replacement failed: " + run.stderr[-600:])
            return storage.read_bytes()
    return read_streams, xml, naming_inventory, link_contract, replace_ole


def ids_of(root: E._Element) -> dict[str, E._Element]:
    result = {node.get("id"): node for node in root if node.get("id")}
    need(len(result) == sum(node.get("id") is not None for node in root), "duplicate native model IDs")
    return result


def one(parent: E._Element, path: str) -> E._Element:
    found = parent.xpath(path, namespaces=NS) if ":" in path else parent.findall(path)
    need(len(found) == 1, f"missing or repeated {path}")
    return found[0]


def tag_map(z: zipfile.ZipFile, rels: E._Element) -> dict[str, str]:
    result: dict[str, str] = {}
    for rel in rels:
        if not rel.get("Type", "").endswith("/tags"):
            continue
        part = "ppt/tags/" + rel.get("Target", "").rsplit("/", 1)[-1]
        need(part in z.namelist(), "tag relationship points to a missing part")
        nodes = E.fromstring(z.read(part)).findall("{%s}tag" % P)
        values = [x.get("val") for x in nodes if x.get("name", "").upper() == "THINKCELLSHAPEDONOTDELETE"]
        need(len(values) == 1 and values[0], "invalid think-cell shape tag part")
        result[rel.get("Id")] = values[0]
    return result


def physical_by_tag(slide: E._Element, tags: dict[str, str], value: str, kind: str) -> E._Element:
    path = ".//*[p:nvSpPr/p:cNvPr]" if kind == "sp" else ".//*[p:nvCxnSpPr/p:cNvPr]"
    matches = []
    for node in slide.xpath(path, namespaces=NS):
        refs = node.xpath(".//p:nvPr/p:custDataLst/p:tags/@r:id", namespaces=NS)
        if len(refs) == 1 and tags.get(refs[0]) == value:
            matches.append(node)
    need(len(matches) == 1, f"native tag does not resolve to one physical {kind}: {value}")
    return matches[0]


def chart_candidates(raw: bytes, read_streams, read_xml) -> list[dict]:
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        candidates = []
        for part in z.namelist():
            if not (part.startswith("ppt/embeddings/oleObject") and part.endswith(".bin")):
                continue
            payload = read_streams(z.read(part)).get(("think-cellXML",))
            if not payload:
                continue
            root = read_xml(payload)
            for chart in root:
                if chart.tag != "CSequenceChartSE":
                    continue
                table_ref = chart.find("m_dtable")
                ids = ids_of(root)
                table = ids.get(None if table_ref is None else table_ref.get("idref"))
                if table is None or table.tag != "CSequenceChartDataTable":
                    continue
                vectors = []
                for elem in table.find("ocol").findall("elem") if table.find("ocol") is not None else []:
                    vector = ids.get(elem.get("idref"))
                    if vector is not None and vector.find("m_nIndexInDataSheet") is not None:
                        vectors.append(vector)
                cagr = [x for x in table.find("m_cscdveccon").findall("elem") if ids.get(x.get("idref")) is not None] if table.find("m_cscdveccon") is not None else []
                ver = root.find("version")
                candidates.append({"carrier": part, "chart_id": chart.get("id"), "chart_name": chart.findtext("m_strName") or "", "model_version": "" if ver is None else ver.get("val", ""), "series_count": len(table.find("m_cscdser").findall("elem")) if table.find("m_cscdser") is not None else 0, "category_count": len(vectors), "cagr_count": len(cagr), "vectors": vectors, "root": root})
        return candidates


def endpoint_scalars(root: E._Element, table: E._Element) -> dict[int, list[E._Element]]:
    ids = ids_of(root)
    result: dict[int, list[E._Element]] = {}
    for elem in table.find("ocol").findall("elem"):
        vector = ids.get(elem.get("idref"))
        if vector is None or vector.find("m_nIndexInDataSheet") is None:
            continue
        idx = int(vector.find("m_nIndexInDataSheet").get("val"))
        result[idx] = [ids[e.get("idref")] for e in vector.find("ocol").findall("elem") if ids.get(e.get("idref")) is not None]
    return result


def literal_percent(value: float) -> tuple[str, str]:
    display = f"{value * 100:.1f}%"
    return display, "'" + "''".join(display) + "'"


NUMERIC_PERCENT_XML = """<CTextVariable><m_bstrFormat>'0''.''0''%'</m_bstrFormat><m_prec17834><m_bNumberIsYear val=\"0\"/><m_esigndisplay val=\"0\"/><m_nSignPosition val=\"-2147483648\"/><m_chMinusSymbol>-</m_chMinusSymbol><m_nDecimalDigits17909 val=\"1\"/><m_chDecimalSymbol17909>.</m_chDecimalSymbol17909><m_nGroupingDigits17909 val=\"3\"/><m_chGroupingSymbol17909>,</m_chGroupingSymbol17909><m_strPrefix/><m_strSuffix17909>%</m_strSuffix17909><m_nMagnitude17909 val=\"0\"/><m_yearfmt><begin val=\"0\"/><end val=\"4\"/></m_yearfmt></m_prec17834><m_bUseExcelFont val=\"0\"/><m_bUseExcelFontColor val=\"0\"/></CTextVariable>"""


def prepare(input_path: Path, expected_sha: str, output: Path, report: Path, workdir: Path, skill_dir: Path, chart_name: str | None, source_index: int, sink_index: int) -> dict:
    read_streams, read_xml, naming_inventory, link_contract, replace_ole = init_helpers(skill_dir)
    source = input_path.resolve(); output = output.resolve(); report = report.resolve(); workdir = workdir.resolve()
    need(source.is_file() and sha(source) == expected_sha.upper(), "input SHA-256 mismatch")
    need(workdir.is_dir() and output.parent == workdir and report.parent == workdir and not output.exists() and not report.exists(), "outputs must be new direct children of workdir")
    raw = source.read_bytes()
    candidates = chart_candidates(raw, read_streams, read_xml)
    selected = [x for x in candidates if chart_name is None or x["chart_name"] == chart_name]
    need(len(selected) == 1, f"chart selector must resolve exactly one CSequenceChartSE; found {len(selected)}")
    chosen = selected[0]; need(chosen["category_count"] > sink_index >= 0 and sink_index > source_index, "endpoint indices must be valid and increasing")
    need(chosen["cagr_count"] >= 1, "source chart must contain an existing native CAGR to clone")
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        carrier = z.read(chosen["carrier"]); slide = E.fromstring(z.read("ppt/slides/slide1.xml")); rels = E.fromstring(z.read("ppt/slides/_rels/slide1.xml.rels")); ct = E.fromstring(z.read("[Content_Types].xml"))
        root = chosen["root"]; ids = ids_of(root); table = ids[one(chosen["root"].find("CSequenceChartSE"), "m_dtable").get("idref")]
        scalar_map = endpoint_scalars(root, table); need(source_index in scalar_map and sink_index in scalar_map, "selected endpoint category has no scalar")
        need(len(scalar_map[source_index]) == 1 and len(scalar_map[sink_index]) == 1, "arbitrary endpoint route currently requires one series")
        graph_ref = table.find("m_cscdveccon/elem").get("idref"); graph = ids[graph_ref]; need(graph.tag == "CSequenceChartDataVectorCAGR", "first existing feature is not CAGR")
        old_label = ids[one(graph, "m_scdvecconlabel").get("idref")]; old_tag = old_label.findtext(".//m_bstrShapeName"); tags = tag_map(z, rels); old_shape = physical_by_tag(slide, tags, old_tag, "sp")
        line_refs = one(one(graph, "m_pptgenlineArrow"), "m_cpptline").findall("elem"); old_line_shapes = [physical_by_tag(slide, tags, one(ids[x.get("idref")], "m_bstrShapeName").text, "cxnSp") for x in line_refs]
        next_id = max(int(x) for x in ids if x.isdigit()) + 1; mapping: dict[str, str] = {graph.get("id"): str(next_id)}; next_id += 1
        fields = ("m_varsrcRelative", "m_varsrcAbsolute", "m_varsrcRelativeGroup", "m_varsrcCAGR", "m_scdvecconlabel")
        for field in fields: mapping[one(graph, field).get("idref")] = str(next_id); next_id += 1
        for item in line_refs: mapping[item.get("idref")] = str(next_id); next_id += 1
        source_anchor_id, sink_anchor_id = str(next_id), str(next_id + 1); next_id += 2; textvar_id = str(next_id)
        clone = copy.deepcopy(graph); clone.set("id", mapping[graph.get("id")])
        for node in clone.iter():
            if node.get("idref") in mapping: node.set("idref", mapping[node.get("idref")])
        clone.find("m_anchorSource").set("idref", source_anchor_id); clone.find("m_anchorSink").set("idref", sink_anchor_id)
        cloned_nodes = []
        for old_id in [one(graph, f).get("idref") for f in fields] + [x.get("idref") for x in line_refs]:
            node = copy.deepcopy(ids[old_id]); node.set("id", mapping[old_id])
            if node.tag == "CVariableSource": one(node, "m_guid").set("val", str(uuid.uuid4()))
            if node.tag == "CSequenceChartDataVectorConnectorLabel": one(node, ".//m_bstrShapeName").text = unique_tag("t")
            if node.tag == "CPPTLine": one(node, "m_bstrShapeName").text = unique_tag("t")
            cloned_nodes.append(node)
        new_cagr = next(x for x in cloned_nodes if x.get("id") == mapping[one(graph, "m_varsrcCAGR").get("idref")]); cache, key = literal_percent(float(one(new_cagr, "m_varval").get("val")))
        textvar = E.fromstring(NUMERIC_PERCENT_XML); textvar.set("id", textvar_id); textvar.find("m_bstrFormat").text = key; E.SubElement(one(new_cagr, "m_ctextvar"), "elem", idref=textvar_id)
        for idx, anchor_id in ((source_index, source_anchor_id), (sink_index, sink_anchor_id)):
            scalar = scalar_map[idx][0]; old = scalar.find("m_anchorVectorConnector")
            if old is not None:
                old_anchor = ids.get(old.get("idref")); need(old_anchor is not None, "selected scalar has an unknown existing anchor")
                old_members = old_anchor.find("m_cfeature")
                need(old_members is None or not old_members.findall("elem"), "selected endpoint already owns a native feature; preserve it and choose another endpoint")
            E.SubElement(scalar, "m_anchorVectorConnector", idref=anchor_id) if old is None else scalar.find("m_anchorVectorConnector").set("idref", anchor_id)
            anchor = E.Element("CSequenceChartAnchor", id=anchor_id); members = E.SubElement(anchor, "m_cfeature", length="1"); E.SubElement(members, "elem", idref=clone.get("id")); root.append(anchor)
        E.SubElement(table.find("m_cscdveccon"), "elem", idref=clone.get("id")); table.find("m_cscdveccon").set("length", str(len(table.find("m_cscdveccon").findall("elem"))))
        root.insert(root.index(graph) + 1, clone)
        for node in cloned_nodes + [textvar]: root.insert(root.index(graph) + 1, node)
        changed_carrier = replace_ole(carrier, E.tostring(root, encoding="utf-8"), workdir)
        new_tag = one(next(x for x in cloned_nodes if x.get("id") == mapping[one(graph, "m_scdvecconlabel").get("idref")]), ".//m_bstrShapeName").text
        used = [int(x.get("id")) for x in slide.xpath(".//p:cNvPr", namespaces=NS) if (x.get("id") or "").isdigit()]; used_rids = {x.get("Id") for x in rels}; serial = 1
        new_parts = {}
        def add_shape(source_shape: E._Element, kind: str, tag: str, name: str) -> None:
            nonlocal serial
            shape = copy.deepcopy(source_shape); sid = str(max(used) + 1); used.append(int(sid)); one(shape, ".//p:cNvPr").set("id", sid); one(shape, ".//p:cNvPr").set("name", name + " " + sid)
            while "rId" + str(serial) in used_rids: serial += 1
            rid, part = "rId" + str(serial), "cagr-parameterized-" + str(serial) + ".xml"; serial += 1; used_rids.add(rid); one(shape, ".//p:nvPr/p:custDataLst/p:tags").set("{%s}id" % R, rid); one(slide, ".//p:spTree").append(shape)
            E.SubElement(rels, "{%s}Relationship" % PR, Id=rid, Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags", Target="../tags/" + part); E.SubElement(ct, "{%s}Override" % CT, PartName="/ppt/tags/" + part, ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"); new_parts["ppt/tags/" + part] = (f'<p:tagLst xmlns:p="{P}"><p:tag name="THINKCELLSHAPEDONOTDELETE" val="{tag}"/></p:tagLst>').encode()
        label = copy.deepcopy(old_shape); one(label, ".//p:cNvPr").set("id", str(max(used) + 1)); used.append(int(one(label, ".//p:cNvPr").get("id"))); one(label, ".//p:cNvPr").set("name", "CAGR parameterized label"); while_id = serial
        while "rId" + str(serial) in used_rids: serial += 1
        rid, part = "rId" + str(serial), "cagr-parameterized-label.xml"; used_rids.add(rid); one(label, ".//p:nvPr/p:custDataLst/p:tags").set("{%s}id" % R, rid); para = one(label, ".//a:p"); oldrun = one(para, "a:r"); rpr = copy.deepcopy(one(oldrun, "a:rPr")); [para.remove(x) for x in list(para) if x.tag in {"{%s}r" % A, "{%s}fld" % A}]; ppr = one(para, "a:pPr"); fld = E.Element("{%s}fld" % A, id="{" + str(uuid.uuid4()).upper() + "}", type="datetime" + key); fld.append(rpr); E.SubElement(fld, "{%s}t" % A).text = cache; para.insert(para.index(ppr) + 1, fld); one(slide, ".//p:spTree").append(label); E.SubElement(rels, "{%s}Relationship" % PR, Id=rid, Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags", Target="../tags/" + part); E.SubElement(ct, "{%s}Override" % CT, PartName="/ppt/tags/" + part, ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"); new_parts["ppt/tags/" + part] = (f'<p:tagLst xmlns:p="{P}"><p:tag name="THINKCELLSHAPEDONOTDELETE" val="{new_tag}"/></p:tagLst>').encode()
        for old_shape, clone_node in zip(old_line_shapes, [x for x in cloned_nodes if x.tag == "CPPTLine"]): add_shape(old_shape, "cxnSp", one(clone_node, "m_bstrShapeName").text, "CAGR parameterized arrow")
        modified = {chosen["carrier"]: changed_carrier, "ppt/slides/slide1.xml": E.tostring(slide, xml_declaration=True, encoding="UTF-8", standalone=True), "ppt/slides/_rels/slide1.xml.rels": E.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True), "[Content_Types].xml": E.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True), **new_parts}
        with zipfile.ZipFile(output, "x") as dst:
            for item in z.infolist():
                if item.filename not in modified: dst.writestr(copy.copy(item), z.read(item.filename))
            for name, data in modified.items(): dst.writestr(name, data)
    result = {"status": "PREPARED_PARAMETERIZED_CAGR_REGENERATION_REQUIRED", "source_sha256": expected_sha.upper(), "candidate_sha256": sha(output), "chart": {k: chosen[k] for k in ("carrier", "chart_id", "chart_name", "model_version", "series_count", "category_count", "cagr_count")}, "endpoints": {"source_index": source_index, "sink_index": sink_index, "series_mode": "single-series-scalar"}, "inserted": {"feature_model_id": clone.get("id"), "label_tag": new_tag, "text_variable_id": textvar_id, "source_anchor_id": source_anchor_id, "sink_anchor_id": sink_anchor_id}, "source_unchanged": sha(source) == expected_sha.upper()}
    report.write_text(json.dumps(result, indent=2), encoding="utf-8"); return result


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--input", type=Path, required=True); p.add_argument("--expected-sha256", required=True); p.add_argument("--output", type=Path, required=True); p.add_argument("--report", type=Path, required=True); p.add_argument("--workdir", type=Path, required=True); p.add_argument("--skill-dir", type=Path, required=True); p.add_argument("--chart-name"); p.add_argument("--source-index", type=int, required=True); p.add_argument("--sink-index", type=int, required=True)
    a = p.parse_args(); print(json.dumps(prepare(a.input, a.expected_sha256, a.output, a.report, a.workdir, a.skill_dir.resolve(), a.chart_name, a.source_index, a.sink_index), indent=2))
