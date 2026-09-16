"""Portable offline adapter for a native think-cell series connector.

The adapter copies a complete connector/property/shape relationship closure
from a supplied native donor chart into a supplied target chart. Chart,
category, and series selection is semantic. It never starts Office or ppttc;
the output is a candidate for the official regeneration/native-reopen step.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import importlib
import json
import os
import posixpath
import subprocess
import sys
import tempfile
import uuid
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree as E

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"p": P, "a": A, "r": R}


def fail(message: str) -> None:
    raise RuntimeError(message)


def need(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def expected_digest(path: Path, expected: str | None, role: str) -> str:
    """Read a stable starting fingerprint and enforce an optional caller guard."""
    actual = digest(path)
    if expected is not None:
        need(actual == expected.strip().upper(), f"{role} SHA-256 does not match --expected-{role}-sha256")
    return actual


def assert_stable(path: Path, before: str, role: str) -> None:
    """Reject any mutation of either source fixture during preparation."""
    after = digest(path)
    need(after == before, f"{role} changed during adapter run; refusing to publish candidate")


def resolve_part(base: str, target: str) -> str:
    target = target.split("#", 1)[0]
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(base), target))


def discover_skill_dir(explicit: Path | None) -> Path:
    candidates: list[Path] = []
    if explicit:
        candidates.append(explicit)
    env_dir = os.environ.get("THINKCELL_SKILL_DIR")
    if env_dir:
        candidates.append(Path(env_dir))
    here = Path(__file__).resolve()
    candidates.extend([here.parent, here.parent.parent])
    candidates.extend(parent / "staged-plugin" / "skills" / "thinkcell-edit" / "scripts" for parent in [here, *here.parents])
    for candidate in candidates:
        if (candidate / "chart_geometry.py").is_file() and (candidate / "thinkcell_no_click" / "implementation" / "replace_ole_stream.ps1").is_file():
            return candidate.resolve()
    fail("Could not locate think-cell scripts; pass --skill-dir or set THINKCELL_SKILL_DIR")


def load_helpers(skill_dir: Path):
    sys.path.insert(0, str(skill_dir))
    try:
        geometry = importlib.import_module("chart_geometry")
    except ImportError as exc:
        fail(f"Cannot import chart_geometry from {skill_dir}: {exc}")
    return geometry.streams, geometry.xml, skill_dir / "thinkcell_no_click" / "implementation" / "replace_ole_stream.ps1"


def slide_contexts(z: zipfile.ZipFile, streams_fn):
    pres = E.fromstring(z.read("ppt/presentation.xml"))
    prels = E.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    sld_list = pres.find(f"{{{P}}}sldIdLst")
    need(sld_list is not None, "presentation has no slide list")
    rel_by_id = {node.get("Id"): node for node in prels}
    for ordinal, sld in enumerate(sld_list, 1):
        rel = rel_by_id.get(sld.get(f"{{{R}}}id"))
        if rel is None:
            continue
        slide_part = resolve_part("ppt/presentation.xml", rel.get("Target") or "")
        rel_part = slide_part.rsplit("/", 1)[0] + "/_rels/" + slide_part.rsplit("/", 1)[1] + ".rels"
        if slide_part not in z.namelist() or rel_part not in z.namelist():
            continue
        slide = E.fromstring(z.read(slide_part)); rels = E.fromstring(z.read(rel_part))
        rel_map = {node.get("Id"): node for node in rels}
        carriers = []
        for ole in slide.xpath(".//p:oleObj", namespaces=NS):
            srel = rel_map.get(ole.get(f"{{{R}}}id"))
            if srel is None:
                continue
            part = resolve_part(slide_part, srel.get("Target") or "")
            if part not in z.namelist():
                continue
            try:
                if ("think-cellXML",) in streams_fn(z.read(part)):
                    if part not in carriers:
                        carriers.append(part)
            except Exception:
                continue
        for carrier in carriers:
            yield {"ordinal": ordinal, "slide_part": slide_part, "slide": slide, "rels": rels, "carrier": carrier}


def model_ids(root: E._Element) -> dict[str, E._Element]:
    return {node.get("id"): node for node in root.iter() if node.get("id")}


def text_source(root: E._Element, source_id: str | None) -> str | None:
    if not source_id:
        return None
    node = model_ids(root).get(source_id)
    if node is None or node.tag != "CVariableSource":
        return None
    value = node.find("m_varval")
    return value.text if value is not None else None


def vectors(root: E._Element) -> dict[int, dict[str, Any]]:
    ids = model_ids(root); result: dict[int, dict[str, Any]] = {}
    for vector in root.iter("CSequenceChartDataVector"):
        idx = vector.find("m_nIndexInDataSheet"); cells = vector.find("ocol")
        if idx is None or cells is None:
            continue
        index = int(idx.get("val")); cat = text_source(root, vector.find("m_varsrcCategory").get("idref") if vector.find("m_varsrcCategory") is not None else None)
        result[index] = {"node": vector, "label": cat, "scalars": [ids.get(cell.get("idref")) for cell in cells.findall("elem")]}
    return result


def chart_record(context: dict[str, Any], root: E._Element) -> dict[str, Any]:
    ids = model_ids(root); records = []
    for owner in root.iter("CSequenceChartSE"):
        chart_ref = owner.find("m_pptseqchart")
        chart = ids.get(chart_ref.get("idref")) if chart_ref is not None else None
        if chart is None or chart.tag != "CPPTSequenceChart":
            continue
        table_ref = owner.find("m_dtable"); table = ids.get(table_ref.get("idref")) if table_ref is not None else None
        series = []
        if table is not None:
            for item in table.findall("m_cscdser/elem"):
                ser = ids.get(item.get("idref"))
                if ser is not None:
                    source = ser.find("m_varsrc")
                    series.append(text_source(root, source.get("idref") if source is not None else None))
        records.append({"context": context, "root": root, "owner": owner, "chart": chart, "table": table, "name": (owner.findtext("m_strName") or ""), "range_name": (table.find("m_bstrRangeName").get("val") if table is not None and table.find("m_bstrRangeName") is not None else ""), "shape_name": chart.findtext("m_bstrShapeName") or "", "series": series, "vectors": vectors(root)})
    return records[0] if len(records) == 1 else {"records": records}


def select_chart(z: zipfile.ZipFile, streams_fn, selector: str | None) -> dict[str, Any]:
    all_records = []
    for context in slide_contexts(z, streams_fn):
        root = importlib.import_module("chart_geometry").xml(streams_fn(z.read(context["carrier"]))[("think-cellXML",)])
        # chart_geometry streams uses tuple keys in current releases; retain a
        # clear error if a future release changes its stream contract.
        if not isinstance(root, E._Element):
            fail("think-cell carrier did not expose an XML model")
        rec = chart_record(context, root)
        all_records.extend(rec["records"] if "records" in rec else [rec])
    if selector:
        matches = [r for r in all_records if selector in {r["name"], r["range_name"], r["shape_name"]}]
    else:
        matches = all_records
    need(len(matches) == 1, f"chart selector matched {len(matches)} charts; pass --chart-name (available names: {[r['name'] or r['range_name'] or r['shape_name'] for r in all_records]})")
    return matches[0]


def choose_label(mapping: dict[str, int], label: str | None, index: int | None, kind: str) -> tuple[int, str]:
    if label is not None:
        matches = [idx for text, idx in mapping.items() if text == label]
        need(len(matches) == 1, f"{kind} label {label!r} matched {len(matches)} entries")
        return matches[0], label
    need(index is not None and index in mapping.values(), f"pass --{kind.replace(' ', '-')} or a valid --{kind.replace(' ', '-')}-index")
    return index, next(text for text, idx in mapping.items() if idx == index)


def series_mapping(record: dict[str, Any]) -> dict[str, int]:
    return {label: idx for idx, label in enumerate(record["series"]) if label}


def category_mapping(record: dict[str, Any]) -> dict[str, int]:
    return {entry["label"]: idx for idx, entry in record["vectors"].items() if entry["label"]}


def scalar_at(record: dict[str, Any], category_idx: int, series_idx: int) -> E._Element:
    scalar = record["vectors"].get(category_idx, {}).get("scalars", [None])[series_idx]
    need(scalar is not None, f"series slot {series_idx} is missing in category index {category_idx}")
    return scalar


def physical_shape(z: zipfile.ZipFile, context: dict[str, Any], line_tag: str):
    rel_by_id = {node.get("Id"): node for node in context["rels"]}
    matches = []
    for shape in context["slide"].xpath(".//*[p:nvSpPr/p:cNvPr] | .//*[p:nvCxnSpPr/p:cNvPr]", namespaces=NS):
        for attr in shape.xpath(".//p:nvPr/p:custDataLst/p:tags/@r:id", namespaces=NS):
            rel = rel_by_id.get(attr)
            if rel is None:
                continue
            part = resolve_part(context["slide_part"], rel.get("Target") or "")
            if part in z.namelist():
                node = E.fromstring(z.read(part)); tag = node.find(f"{{{P}}}tag")
                if tag is not None and tag.get("val") == line_tag:
                    matches.append((shape, rel, part))
    need(len(matches) == 1, f"donor physical shape for line tag {line_tag!r} matched {len(matches)} shapes")
    return matches[0]


def clone_feature_closure(target_root: E._Element, donor_root: E._Element, donor_scalar: E._Element, target_source: E._Element, target_sink: E._Element) -> dict[str, Any]:
    donor_ids = model_ids(donor_root); target_ids = model_ids(target_root)
    binding = donor_scalar.find("m_scdscconnect")
    need(binding is not None and binding.get("idref") not in {None, "0"}, "donor semantic source has no connector")
    connector = donor_ids.get(binding.get("idref")); need(connector is not None and connector.tag == "CSequenceChartDataScalarConnector", "donor source connector is not native")
    # Follow every idref in the connector/property closure, including both
    # anchors and the CPPTLine. This keeps future connector properties intact.
    closure: dict[str, E._Element] = {}; pending = [connector]
    while pending:
        node = pending.pop(); old_id = node.get("id")
        if old_id in closure:
            continue
        closure[old_id] = node
        for attr in node.xpath(".//@idref"):
            child = donor_ids.get(attr)
            if child is not None and child.get("id") not in closure:
                pending.append(child)
    numeric = [int(value) for value in target_ids if value.isdigit()]
    next_id = max(numeric, default=0) + 1; remap = {old: str(next_id + i) for i, old in enumerate(closure)}
    clones = []
    for old, node in closure.items():
        clone = copy.deepcopy(node); clone.set("id", remap[old])
        for item in clone.xpath(".//*[@idref]"):
            if item.get("idref") in remap:
                item.set("idref", remap[item.get("idref")])
        clones.append(clone)
    cloned_connector = next(node for node in clones if node.tag == "CSequenceChartDataScalarConnector")
    line_ref = cloned_connector.find(".//m_cpptline/elem"); need(line_ref is not None, "connector closure has no CPPTLine")
    line = next((node for node in clones if node.get("id") == line_ref.get("idref")), None); need(line is not None and line.tag == "CPPTLine", "connector closure line is not CPPTLine")
    target_source_binding = target_source.find("m_scdscconnect")
    need(target_source_binding is not None and target_source_binding.get("idref") in {None, "0"}, "target source already has a connector")
    need(target_source.find("m_anchorSource") is None and target_sink.find("m_anchorSink") is None, "target endpoint already has an anchor")
    target_source_binding.set("idref", cloned_connector.get("id")); E.SubElement(target_source, "m_anchorSource", idref=remap[connector.find("m_anchorSource").get("idref")]); E.SubElement(target_sink, "m_anchorSink", idref=remap[connector.find("m_anchorSink").get("idref")])
    target_root.extend(clones)
    return {"connector": cloned_connector, "line": line, "closure_size": len(clones), "source_scalar_id": target_source.get("id"), "sink_scalar_id": target_sink.get("id")}


def rewrite_package(target: Path, donor: Path, output: Path, target_rec: dict[str, Any], donor_rec: dict[str, Any], line: E._Element, shape: E._Element, donor_rel: E._Element, donor_tag_part: str, target_carrier: bytes, model_payload: bytes, skill_dir: Path, streams_fn, report_path: Path, report: dict[str, Any]) -> None:
    helper = skill_dir / "thinkcell_no_click" / "implementation" / "replace_ole_stream.ps1"
    with tempfile.TemporaryDirectory(prefix="series-connector-") as temp:
        temp_path = Path(temp); storage = temp_path / "carrier.bin"; model = temp_path / "model.xml"; storage.write_bytes(target_carrier); model.write_bytes(model_payload)
        proc = subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(helper), "-StoragePath", str(storage), "-StreamBytesPath", str(model)], capture_output=True, text=True, timeout=90)
        need(proc.returncode == 0, "OLE model replacement failed: " + proc.stderr[-1000:]); replacement = storage.read_bytes()
    with zipfile.ZipFile(target) as src, zipfile.ZipFile(output, "x") as dst:
        for item in src.infolist():
            name = item.filename
            if name == target_rec["context"]["carrier"]:
                dst.writestr(copy.copy(item), replacement)
            elif name == target_rec["context"]["slide_part"]:
                dst.writestr(copy.copy(item), E.tostring(target_rec["context"]["slide"], xml_declaration=True, encoding="UTF-8", standalone=True))
            elif name == target_rec["context"]["slide_part"].rsplit("/", 1)[0] + "/_rels/" + target_rec["context"]["slide_part"].rsplit("/", 1)[1] + ".rels":
                dst.writestr(copy.copy(item), E.tostring(target_rec["context"]["rels"], xml_declaration=True, encoding="UTF-8", standalone=True))
            elif name == "[Content_Types].xml":
                dst.writestr(copy.copy(item), report.pop("content_types_bytes"))
            else:
                dst.writestr(copy.copy(item), src.read(name))
        dst.writestr(f"ppt/tags/{report['closure']['physical_tag']}.xml", report.pop("tag_xml"))


def inspect(path: Path, streams_fn, xml_fn) -> dict[str, Any]:
    with zipfile.ZipFile(path) as z:
        charts = []
        for context in slide_contexts(z, streams_fn):
            root = xml_fn(streams_fn(z.read(context["carrier"]))[("think-cellXML",)])
            rec = chart_record(context, root)
            for item in rec["records"] if "records" in rec else [rec]:
                charts.append({"slide": context["ordinal"], "chart_name": item["name"], "range_name": item["range_name"], "shape_name": item["shape_name"], "series": item["series"], "categories": [item["vectors"][idx]["label"] for idx in sorted(item["vectors"]) if item["vectors"][idx]["label"]]})
        return {"status": "INSPECT_PASS", "input": str(path), "sha256": digest(path), "charts": charts}


def prepare(args: argparse.Namespace) -> dict[str, Any]:
    skill_dir = discover_skill_dir(args.skill_dir); streams_fn, xml_fn, _ = load_helpers(skill_dir)
    target = args.input.resolve()
    need(target.is_file(), "input must be a saved PPTX file")
    if args.inspect:
        return inspect(target, streams_fn, xml_fn)
    need(args.donor is not None and args.output is not None and args.report is not None, "--donor, --output, and --report are required when preparing")
    donor = args.donor.resolve(); output = args.output.resolve(); report_path = args.report.resolve()
    need(donor.is_file(), "donor must be a saved PPTX file"); need(output not in {target, donor} and report_path not in {target, donor, output}, "output/report must be distinct from inputs"); need(not output.exists() and not report_path.exists(), "refusing to overwrite output/report")
    target_sha_before = expected_digest(target, args.expected_input_sha256, "input")
    donor_sha_before = expected_digest(donor, args.expected_donor_sha256, "donor")
    with zipfile.ZipFile(target) as tz, zipfile.ZipFile(donor) as dz:
        target_rec = select_chart(tz, streams_fn, args.chart_name); donor_rec = select_chart(dz, streams_fn, args.donor_chart_name or args.chart_name)
        t_series_map = series_mapping(target_rec); d_series_map = series_mapping(donor_rec)
        series_idx, series_label = choose_label(t_series_map, args.series_name, args.series_index, "series")
        donor_series_idx, _ = choose_label(d_series_map, args.donor_series_name or series_label, args.donor_series_index, "series")
        t_cat_map = category_mapping(target_rec); d_cat_map = category_mapping(donor_rec)
        from_idx, from_label = choose_label(t_cat_map, args.from_category, args.from_category_index, "from-category")
        to_idx, to_label = choose_label(t_cat_map, args.to_category, args.to_category_index, "to-category")
        donor_from_idx, _ = choose_label(d_cat_map, args.donor_from_category or from_label, args.donor_from_category_index, "donor-from-category")
        donor_to_idx, _ = choose_label(d_cat_map, args.donor_to_category or to_label, args.donor_to_category_index, "donor-to-category")
        target_root = target_rec["root"]; donor_root = donor_rec["root"]
        # Give the target chart/table one stable, reportable automation name
        # when the source has not named it. This name is derived from the
        # supplied source file, never from donor fixture IDs.
        owner_name = target_rec["owner"].find("m_strName")
        table_name = target_rec["table"].find("m_strName") if target_rec["table"] is not None else None
        automation_name = args.automation_name or (owner_name.text if owner_name is not None and owner_name.text else None) or ("TC_SC_" + digest(target)[:12])
        if owner_name is None:
            owner_name = E.SubElement(target_rec["owner"], "m_strName")
        owner_name.text = automation_name
        if target_rec["table"] is not None:
            if table_name is None:
                table_name = E.Element("m_strName")
                excel_top = target_rec["table"].find("m_bExcelOnTop")
                need(excel_top is not None, "target data table insertion point missing")
                target_rec["table"].insert(target_rec["table"].index(excel_top), table_name)
            table_name.text = automation_name
        target_source = scalar_at(target_rec, from_idx, series_idx); target_sink = scalar_at(target_rec, to_idx, series_idx); donor_source = scalar_at(donor_rec, donor_from_idx, donor_series_idx); donor_sink = scalar_at(donor_rec, donor_to_idx, donor_series_idx)
        donor_sink_anchor = donor_sink.find("m_anchorSink"); need(donor_sink_anchor is not None, "donor semantic sink has no native sink anchor")
        donor_ids = model_ids(donor_root); donor_connector_ref = donor_source.find("m_scdscconnect"); donor_connector = donor_ids.get(donor_connector_ref.get("idref")) if donor_connector_ref is not None else None
        need(donor_connector is not None and donor_sink_anchor.get("idref") == donor_connector.find("m_anchorSink").get("idref"), "donor source/sink semantics do not resolve to one connector")
        closure = clone_feature_closure(target_root, donor_root, donor_source, target_source, target_sink)
        line_tag = closure["line"].findtext("m_bstrShapeName"); donor_shape, donor_rel, donor_tag_part = physical_shape(dz, donor_rec["context"], line_tag)
        new_tag = (args.tag_prefix or "sc") + "-" + uuid.uuid4().hex; closure["line"].find("m_bstrShapeName").text = new_tag
        clone_shape = copy.deepcopy(donor_shape); c_nv = clone_shape.xpath(".//p:cNvPr", namespaces=NS)[0]; used = [int(n.get("id")) for n in target_rec["context"]["slide"].xpath(".//p:cNvPr", namespaces=NS) if (n.get("id") or "").isdigit()]; new_shape_id = max(used, default=0) + 1; c_nv.set("id", str(new_shape_id)); c_nv.set("name", "Native series connector " + str(new_shape_id))
        donor_tag_refs = donor_shape.xpath(".//p:nvPr/p:custDataLst/p:tags/@r:id", namespaces=NS)
        need(len(donor_tag_refs) == 1, "donor physical connector shape must have exactly one tags relationship")
        old_rid = donor_tag_refs[0]; target_rids = {n.get("Id") for n in target_rec["context"]["rels"]}; rid_num = 1
        while "rId" + str(rid_num) in target_rids: rid_num += 1
        new_rid = "rId" + str(rid_num); clone_shape.xpath(".//p:nvPr/p:custDataLst/p:tags", namespaces=NS)[0].set(f"{{{R}}}id", new_rid); target_rec["context"]["slide"].find(f".//{{{P}}}spTree").append(clone_shape)
        E.SubElement(target_rec["context"]["rels"], f"{{{PR}}}Relationship", Id=new_rid, Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags", Target=f"../tags/{new_tag}.xml")
        cts = E.fromstring(tz.read("[Content_Types].xml")); donor_cts = E.fromstring(dz.read("[Content_Types].xml")); old_ct = next((n.get("ContentType") for n in donor_cts if n.get("PartName") == "/" + donor_tag_part), None); need(old_ct, "donor tag content type missing"); E.SubElement(cts, f"{{{CT}}}Override", PartName=f"/ppt/tags/{new_tag}.xml", ContentType=old_ct)
        tag_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:tagLst xmlns:p="{P}"><p:tag name="THINKCELLSHAPEDONOTDELETE" val="{new_tag}"/></p:tagLst>'.encode()
        report = {"status": "OFFLINE_SERIES_CONNECTOR_ADAPTER_PASS", "input": str(target), "input_sha256": target_sha_before, "donor": str(donor), "donor_sha256": donor_sha_before, "output": str(output), "automation_name": automation_name, "chart": {"target_name": target_rec["name"], "donor_name": donor_rec["name"], "target_slide": target_rec["context"]["ordinal"], "donor_slide": donor_rec["context"]["ordinal"]}, "semantic_endpoint": {"series_name": series_label, "series_slot": series_idx, "from_category": from_label, "to_category": to_label, "from_category_index": from_idx, "to_category_index": to_idx}, "donor_semantic_endpoint": {"series_name": donor_rec["series"][donor_series_idx], "series_slot": donor_series_idx, "from_category": args.donor_from_category or from_label, "to_category": args.donor_to_category or to_label, "from_category_index": donor_from_idx, "to_category_index": donor_to_idx}, "closure": {"model_nodes": closure["closure_size"], "connector_model_id": closure["connector"].get("id"), "line_model_id": closure["line"].get("id"), "physical_tag": new_tag, "physical_shape_id": str(new_shape_id), "physical_relationship_id": new_rid}, "requires": ["official ppttc regeneration", "PowerPoint native save/reopen", "exact named-datasheet readback", "changed-data interior-boundary geometry check"], "evidence_link": "SERIES_CONNECTOR_ADAPTER.md#native-follow-up"}
        report["content_types_bytes"] = E.tostring(cts, xml_declaration=True, encoding="UTF-8", standalone=True); report["tag_xml"] = tag_xml
        rewrite_package(target, donor, output, target_rec, donor_rec, closure["line"], clone_shape, donor_rel, donor_tag_part, tz.read(target_rec["context"]["carrier"]), E.tostring(target_root, encoding="utf-8"), skill_dir, streams_fn, report_path, report)
    assert_stable(target, target_sha_before, "input")
    assert_stable(donor, donor_sha_before, "donor")
    report["input_unchanged"] = True
    report["donor_unchanged"] = True
    report["output_sha256"] = digest(output); report_path.write_text(json.dumps(report, indent=2), encoding="utf-8"); return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True); parser.add_argument("--donor", type=Path); parser.add_argument("--output", type=Path); parser.add_argument("--report", type=Path); parser.add_argument("--skill-dir", type=Path); parser.add_argument("--inspect", action="store_true")
    parser.add_argument("--chart-name"); parser.add_argument("--donor-chart-name"); parser.add_argument("--automation-name"); parser.add_argument("--series-name"); parser.add_argument("--series-index", type=int); parser.add_argument("--donor-series-name"); parser.add_argument("--donor-series-index", type=int); parser.add_argument("--expected-input-sha256", "--expected-target-sha256", dest="expected_input_sha256"); parser.add_argument("--expected-donor-sha256", dest="expected_donor_sha256")
    parser.add_argument("--from-category"); parser.add_argument("--from-category-index", type=int); parser.add_argument("--to-category"); parser.add_argument("--to-category-index", type=int); parser.add_argument("--donor-from-category"); parser.add_argument("--donor-from-category-index", type=int); parser.add_argument("--donor-to-category"); parser.add_argument("--donor-to-category-index", type=int); parser.add_argument("--tag-prefix", default="sc")
    args = parser.parse_args()
    try:
        result = prepare(args); print(json.dumps(result, indent=2))
    except Exception as exc:
        raise SystemExit(f"series_connector_adapter: {exc}") from exc


if __name__ == "__main__":
    main()
