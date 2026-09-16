"""Official-regeneration and native-readback gate for `insert_native_cagr.py`.

The only supported contract is the proved model-38764, 1-series by 8-category
sequence chart with typed January-1 date cells and a new native CAGR from
indices 1 to 7.  The runner never edits its input and writes a final output
only after every generated and native-reopened check passes.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import io
import json
import math
import subprocess
import sys
import zipfile
from pathlib import Path

from lxml import etree as E

HERE = Path(__file__).resolve().parent
SCRIPTS = HERE.parent
IMPL = SCRIPTS / "thinkcell_no_click" / "implementation"
sys.path[:0] = [str(SCRIPTS), str(IMPL)]
from chart_geometry import streams, xml  # noqa: E402
from office_operation_lock import OfficeOperationLock, run_locked_subprocess  # noqa: E402
from runtime import powershell, powershell_env  # noqa: E402
from thinkcell_no_click.implementation.read_named_datasheet import read as read_sheet  # noqa: E402
from prepare_thinkcell_name import inventory as naming_inventory, link_contract  # noqa: E402

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"p": P, "a": A, "r": R}


def need(ok: bool, message: str) -> None:
    if not ok:
        raise RuntimeError(message)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def one(parent: E._Element, tag: str) -> E._Element:
    nodes = parent.findall(tag, NS if ":" in tag else None)
    need(len(nodes) == 1, "missing or repeated " + tag)
    return nodes[0]


def ids_of(root: E._Element) -> dict[str, E._Element]:
    ids = {node.get("id"): node for node in root if node.get("id")}
    need(len(ids) == sum(node.get("id") is not None for node in root), "duplicate native model IDs")
    return ids


def no_other_chart_models(z: zipfile.ZipFile) -> None:
    charts = []
    for part in [x for x in z.namelist() if x.startswith("ppt/embeddings/oleObject") and x.endswith(".bin")]:
        model = streams(z.read(part)).get(("think-cellXML",))
        if model is not None:
            root = xml(model)
            charts.extend((part, node.get("id")) for node in root if node.tag.endswith("ChartSE"))
    need(charts == [("ppt/embeddings/oleObject13.bin", "7")], "additional or noncanonical active chart carrier is unsupported")


def parse_plan(path: Path) -> dict:
    """Read the complete typed data contract required by official JSON."""
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    need(set(data) == {"name", "dates", "values"}, "data contract must contain only name, dates, values")
    need(isinstance(data["name"], str) and data["name"].strip(), "automation name is required")
    need(isinstance(data["dates"], list) and len(data["dates"]) == 8, "requires all 8 typed category dates")
    need(isinstance(data["values"], list) and len(data["values"]) == 8, "requires all 8 values")
    dates = []
    for raw in data["dates"]:
        value = dt.date.fromisoformat(raw)
        need(value.month == value.day == 1, "only January-1 whole-year categories are proved")
        dates.append(value)
    need(all(b > a for a, b in zip(dates, dates[1:])), "dates must be strictly increasing whole years")
    values = []
    for value in data["values"]:
        need(isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value) and value > 0, "values must be finite positive numbers")
        values.append(float(value))
    return {"name": data["name"], "dates": dates, "values": values}


def make_job(template: Path, plan: dict, path: Path) -> None:
    table = [[None] + [{"date": x.isoformat()} for x in plan["dates"]], [None] * 9,
             [{"string": "Label"}] + [{"number": x} for x in plan["values"]]]
    path.write_text(json.dumps([{"template": str(template), "data": [{"name": plan["name"], "table": table}]}], indent=2), encoding="utf-8")


def tag_map(z: zipfile.ZipFile, rels: E._Element) -> dict[str, str]:
    result = {}
    for rel in rels:
        if rel.get("Type", "").endswith("/tags"):
            part = "ppt/tags/" + rel.get("Target", "").rsplit("/", 1)[-1]
            need(part in z.namelist(), "tag relationship has no part")
            tags = E.fromstring(z.read(part)).findall("{%s}tag" % P)
            vals = [x.get("val") for x in tags if x.get("name", "").upper() == "THINKCELLSHAPEDONOTDELETE"]
            need(len(vals) == 1 and vals[0], "invalid think-cell shape tag")
            result[rel.get("Id")] = vals[0]
    return result


def physical_by_tag(slide: E._Element, tags: dict[str, str], tag: str, kind: str = "sp") -> E._Element:
    matches = []
    path = ".//*[p:nvSpPr/p:cNvPr]" if kind == "sp" else ".//*[p:nvCxnSpPr/p:cNvPr]"
    for shape in slide.xpath(path, namespaces=NS):
        refs = shape.xpath(".//p:nvPr/p:custDataLst/p:tags/@r:id", namespaces=NS)
        if len(refs) == 1 and tags.get(refs[0]) == tag:
            matches.append(shape)
    need(len(matches) == 1, "tag does not resolve to one physical shape")
    return matches[0]


def signature(shape: E._Element, *, field: bool = False) -> dict:
    geom = shape.find("p:spPr/a:prstGeom", NS); xfrm = shape.find("p:spPr/a:xfrm", NS)
    para = one(shape, ".//a:p")
    node = one(para, "a:fld" if field else "a:r")
    rpr = one(node, "a:rPr")
    return {"geometry": None if geom is None else geom.get("prst"), "bounds": None if xfrm is None else {"x": xfrm.find("a:off", NS).get("x"), "y": xfrm.find("a:off", NS).get("y"), "cx": xfrm.find("a:ext", NS).get("cx"), "cy": xfrm.find("a:ext", NS).get("cy")}, "style": {k: rpr.get(k) for k in ("sz", "b", "i", "typeface")}, "text": node.findtext("a:t", namespaces=NS)}


def endpoint(root: E._Element, ids: dict[str, E._Element], anchor_id: str) -> dict:
    matches = []
    for scalar in root.iter("CSequenceChartDataScalar"):
        ref = scalar.find("m_anchorVectorConnector")
        if ref is None or ref.get("idref") != anchor_id:
            continue
        vectors = [vector for vector in root.iter("CSequenceChartDataVector") if scalar.get("id") in [x.get("idref") for x in vector.findall("ocol/elem")]]
        need(len(vectors) == 1, "endpoint scalar has ambiguous category vector")
        vector = vectors[0]; variable = ids.get(one(scalar, "m_varsrcAbsolute").get("idref"))
        need(variable is not None and variable.tag == "CVariableSource", "endpoint absolute variable unresolved")
        value = one(variable, "m_varval").get("val")
        date = vector.find("m_varsrcCategory/m_varval/m_datetime")
        need(date is not None and date.get("val"), "model category is not a typed date")
        matches.append({"scalar": scalar.get("id"), "vector": vector.get("id"), "index": int(one(vector, "m_nIndexInDataSheet").get("val")), "date": date.get("val"), "value": float(value)})
    need(len(matches) == 1, "connector endpoint is missing or ambiguous")
    return matches[0]


def complete_model_matrix(root: E._Element, ids: dict[str, E._Element]) -> tuple[list[str], list[float]]:
    rows = []
    for vector in root.iter("CSequenceChartDataVector"):
        index = int(one(vector, "m_nIndexInDataSheet").get("val"))
        cells = one(vector, "ocol").findall("elem")
        if not cells:
            continue
        need(len(cells) == 1 and cells[0].get("idref") in ids, "requires exactly one scalar in every category")
        scalar = ids[cells[0].get("idref")]
        variable = ids.get(one(scalar, "m_varsrcAbsolute").get("idref"))
        date = vector.find("m_varsrcCategory/m_varval/m_datetime")
        need(variable is not None and date is not None and date.get("val"), "model category is not typed/date-backed")
        rows.append((index, date.get("val"), float(one(variable, "m_varval").get("val"))))
    rows.sort()
    need([x[0] for x in rows] == list(range(8)), "current model does not expose exactly 8 category records")
    return [x[1] for x in rows], [x[2] for x in rows]


def model_and_physical(path: Path, plan: dict, prep: dict) -> dict:
    with zipfile.ZipFile(path) as z:
        no_other_chart_models(z)
        root = xml(streams(z.read("ppt/embeddings/oleObject13.bin"))[("think-cellXML",)])
        ids = ids_of(root); charts = [x for x in root if x.tag == "CSequenceChartSE"]
        need(len(charts) == 1 and one(root, "version").get("val") == "38764", "native profile drifted")
        chart = charts[0]; need(one(chart, "m_bConsistent").get("val") == "1" and one(chart, "m_ect").get("val") == "0", "current chart model is stale or not the proved subtype")
        table = ids.get(one(chart, "m_dtable").get("idref")); need(table is not None and table.tag == "CSequenceChartDataTable", "current datasource missing")
        need(one(table, "m_advisesink").get("idref") == "0" and one(table, "m_bNeedsUpdateFromSheetOnMakeTC").get("val") in (None, "0"), "current model retains an active datasource wrapper")
        members = one(table, "m_cscdveccon").findall("elem"); need(len(members) == 2, "native regeneration did not retain exactly two CAGR graphs")
        graphs = [ids.get(x.get("idref")) for x in members]; need(all(x is not None and x.tag == "CSequenceChartDataVectorCAGR" for x in graphs), "CAGR feature closure changed")
        by_tag = {}
        for graph in graphs:
            label = ids.get(one(graph, "m_scdvecconlabel").get("idref")); need(label is not None and label.tag == "CSequenceChartDataVectorConnectorLabel", "feature label model missing")
            by_tag[label.findtext(".//m_bstrShapeName")] = graph
        need(set(by_tag) == {prep["original"]["label_tag"], prep["inserted"]["tag"]}, "original/new semantic label ownership changed")
        original, inserted = by_tag[prep["original"]["label_tag"]], by_tag[prep["inserted"]["tag"]]
        old_cagr = ids.get(one(original, "m_varsrcCAGR").get("idref")); need(old_cagr is not None and one(old_cagr, "m_guid").get("val") == prep["original"]["source_feature_guid"], "original CAGR source GUID changed")
        old_refs = one(old_cagr, "m_ctextvar").findall("elem")
        need(len(old_refs) <= 1 and all(x.get("idref") in ids for x in old_refs), "original CAGR text-variable closure is stale")
        new_cagr = ids.get(one(inserted, "m_varsrcCAGR").get("idref")); need(new_cagr is not None, "inserted CAGR source missing")
        need(one(new_cagr, "m_guid").get("val") == prep["inserted"]["source_feature_guid"], "inserted CAGR source GUID changed")
        line_refs = one(one(inserted, "m_pptgenlineArrow"), "m_cpptline").findall("elem")
        need(len(line_refs) == 2 and all(x.get("idref") in ids and ids[x.get("idref")].tag == "CPPTLine" for x in line_refs), "inserted model does not own two native arrow lines")
        refs = one(new_cagr, "m_ctextvar").findall("elem"); need(len(refs) == 1 and refs[0].get("idref") in ids, "inserted CTextVariable is missing or nonunique")
        textvar = ids[refs[0].get("idref")]; need(textvar.tag == "CTextVariable", "inserted text binding has wrong type")
        fmt = one(textvar, "m_bstrFormat").text; precision = one(one(textvar, "m_prec17834"), "m_nDecimalDigits17909").get("val"); suffix = one(one(textvar, "m_prec17834"), "m_strSuffix17909").text
        need(precision == "1" and suffix == "%" and sum(x.tag == "CTextVariable" and x.findtext("m_bstrFormat") == fmt for x in root) == 1, "numeric percent formatter is incomplete or ambiguous")
        source = endpoint(root, ids, one(inserted, "m_anchorSource").get("idref")); sink = endpoint(root, ids, one(inserted, "m_anchorSink").get("idref"))
        need((source["index"], sink["index"]) == (1, 7), "inserted endpoints drifted from the proved 1-to-7 boundary")
        expected_start, expected_end = plan["values"][1], plan["values"][7]
        expected_start_date, expected_end_date = plan["dates"][1].isoformat() + "T00:00:00", plan["dates"][7].isoformat() + "T00:00:00"
        need(math.isclose(source["value"], expected_start, rel_tol=0, abs_tol=1e-9) and math.isclose(sink["value"], expected_end, rel_tol=0, abs_tol=1e-9) and source["date"] == expected_start_date and sink["date"] == expected_end_date, "model endpoints do not match authoritative requested cells")
        original_source = endpoint(root, ids, one(original, "m_anchorSource").get("idref")); original_sink = endpoint(root, ids, one(original, "m_anchorSink").get("idref"))
        need((original_source["index"], original_sink["index"], original_source["date"], original_sink["date"], original_source["value"], original_sink["value"]) == (1, 7, expected_start_date, expected_end_date, expected_start, expected_end), "original CAGR endpoints changed")
        years = plan["dates"][7].year - plan["dates"][1].year; expected_cagr = (expected_end / expected_start) ** (1 / years) - 1
        actual_cagr = float(one(new_cagr, "m_varval").get("val")); need(math.isclose(actual_cagr, expected_cagr, rel_tol=0, abs_tol=2e-12), "model CAGR is stale")
        all_dates, all_values = complete_model_matrix(root, ids)
        expected_dates = [x.isoformat() + "T00:00:00" for x in plan["dates"]]
        need(all_dates == expected_dates and all(math.isclose(a, b, rel_tol=0, abs_tol=1e-9) for a, b in zip(all_values, plan["values"])), "complete current model matrix differs from requested datasource")
        slide = E.fromstring(z.read("ppt/slides/slide1.xml")); rels = E.fromstring(z.read("ppt/slides/_rels/slide1.xml.rels")); tags = tag_map(z, rels)
        original_shape = physical_by_tag(slide, tags, prep["original"]["label_tag"]); original_signature = signature(original_shape); original_signature.pop("bounds")
        expected_original = dict(prep["original"]["shape"]); expected_original.pop("bounds")
        need(original_signature == expected_original, "original label text/tag/style changed")
        new_shape = physical_by_tag(slide, tags, prep["inserted"]["tag"]); fields = new_shape.findall(".//a:fld", NS); need(len(fields) == 1, "inserted physical label does not own exactly one field")
        field = fields[0]; need(field.get("type") == "datetime" + fmt, "physical field does not exactly match CTextVariable format")
        new_sig = signature(new_shape, field=True); need(new_sig["geometry"] == "ellipse" and new_sig["style"] == prep["original"]["shape"]["style"], "inserted label geometry/style drifted")
        bounds = new_sig["bounds"]; need(bounds is not None and all(int(bounds[k]) > 0 for k in ("cx", "cy")), "inserted native label has no usable frame")
        expected_visible = f"{expected_cagr * 100:.1f}%"; need(new_sig["text"] == expected_visible, "visible cache disagrees with current model CAGR")
        line_tags = [one(ids[x.get("idref")], "m_bstrShapeName").text for x in line_refs]
        physical_lines = [physical_by_tag(slide, tags, tag, "cxnSp") for tag in line_tags]
        need(len({one(x, ".//p:cNvPr").get("id") for x in physical_lines}) == 2, "inserted arrow physical closure is not unique")
    return {"consistent": True, "original_feature_preserved": True, "original_feature": {"source": original_source, "sink": original_sink, "text_reference_count_before_prepare": prep["original"]["cagr_text_reference_count"], "text_reference_count_after_native": len(old_refs)}, "inserted_feature": {"source": source, "sink": sink, "cagr": actual_cagr, "visible_cache": expected_visible, "format": fmt, "label_bounds": bounds, "physical_arrow_count": 2}, "no_active_stale_wrapper": True}


def datasource(path: Path, plan: dict) -> dict:
    raw = read_sheet(path, 1, plan["name"]); need(raw["source_unchanged"] and len(raw["sheets"]) == 1, "datasheet readback failed")
    sheet = raw["sheets"][0]; cells = {(x["row"], x["column"]): x for x in sheet["nonempty_cells"]}
    need(all((1, col) in cells and (3, col) in cells for col in range(2, 10)), "datasheet has missing date or value cells")
    epoch = dt.date(1899, 12, 30)
    actual_dates, actual_values = [], []
    for index, col in enumerate(range(2, 10)):
        date_cell, value_cell = cells[(1, col)], cells[(3, col)]
        need(date_cell.get("record_id") == 5 and isinstance(date_cell["value"], (int, float)), "datasheet date is not an Excel typed numeric date")
        actual_dates.append((epoch + dt.timedelta(days=int(date_cell["value"]))).isoformat())
        need(isinstance(value_cell["value"], (int, float)) and math.isclose(float(value_cell["value"]), plan["values"][index], rel_tol=0, abs_tol=1e-9), "datasheet value mismatch")
        actual_values.append(float(value_cell["value"]))
    want_dates = [x.isoformat() for x in plan["dates"]]; need(actual_dates == want_dates, "datasheet typed dates mismatch")
    return {"typed_dates": actual_dates, "values": actual_values, "storage_kind": raw["storage_kind"]}


def run(a) -> dict:
    source, output, report, work = a.input.resolve(), a.output.resolve(), a.report.resolve(), a.workdir.resolve()
    need(source.is_file() and sha(source) == a.expected_sha256.upper(), "prepared input is missing or its SHA-256 changed")
    _, named_charts, _ = naming_inventory(source.read_bytes())
    need(len(named_charts) == 1, "requires one internal named chart")
    link_contract(named_charts[0])
    need(work.is_dir() and output.parent == work and report.parent == work and output not in (source, report) and not output.exists() and not report.exists(), "new output/report must be direct children of workdir")
    if a.previous_report:
        previous = json.loads(a.previous_report.read_text(encoding="utf-8-sig"))
        need(previous.get("status") == "NATIVE_DYNAMIC_CAGR_GATES_PASS" and previous.get("output_sha256") == sha(source) and isinstance(previous.get("prepared"), dict), "previous native report does not bind this repeat input")
        prep = previous["prepared"]
    else:
        prep = json.loads(a.prepared_report.read_text(encoding="utf-8-sig")); need(prep.get("status") == "PREPARED_NATIVE_REGENERATION_REQUIRED" and prep.get("candidate_sha256") == sha(source), "prepared report does not bind this candidate")
    plan = parse_plan(a.data); source_hash = sha(source)
    stage = work / "native_cagr_work"; need(not stage.exists(), "work stage already exists"); stage.mkdir()
    job, generated = stage / "update.ppttc", stage / "generated.pptx"; make_job(source, plan, job)
    with OfficeOperationLock("bounded-native-cagr"):
        with (stage / "ppttc.stdout.txt").open("w", encoding="utf-8") as stdout, (stage / "ppttc.stderr.txt").open("w", encoding="utf-8") as stderr:
            code = run_locked_subprocess([str(a.ppttc), str(job), "-o", str(generated)], operation="bounded-native-cagr-json", timeout_seconds=240, stdout=stdout, stderr=stderr, creationflags=subprocess.CREATE_NO_WINDOW).returncode
        need(code == 0 and generated.exists(), "official JSON regeneration failed")
        generated_model, generated_sheet = model_and_physical(generated, plan, prep), datasource(generated, plan)
        native, native_report, render = stage / "native-reopened.pptx", stage / "native-report.json", stage / "native.png"
        command = [powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(IMPL / "native_verify_scoped.ps1"), "-InputFile", str(generated), "-OutputFile", str(native), "-ReportFile", str(native_report), "-RenderFile", str(render)]
        with (stage / "native.stdout.txt").open("w", encoding="utf-8") as stdout, (stage / "native.stderr.txt").open("w", encoding="utf-8") as stderr:
            code = run_locked_subprocess(command, operation="bounded-native-cagr-reopen", timeout_seconds=240, stdout=stdout, stderr=stderr, env=powershell_env(), creationflags=subprocess.CREATE_NO_WINDOW).returncode
        need(code == 0 and native.exists() and native_report.exists() and render.exists(), "native save/reopen failed")
        scope = json.loads(native_report.read_text(encoding="utf-8-sig")); need(scope.get("native_reopen_pass") and scope.get("source_unchanged") and scope.get("other_presentations_unchanged"), "native scope gate failed")
        native_model, native_sheet = model_and_physical(native, plan, prep), datasource(native, plan)
    need(sha(source) == source_hash, "prepared input changed during execution")
    with output.open("xb") as handle:
        handle.write(native.read_bytes())
    result = {"status": "NATIVE_DYNAMIC_CAGR_GATES_PASS", "source_sha256": source_hash, "output_sha256": sha(output), "prepared": prep, "generated": {"model": generated_model, "datasheet": generated_sheet}, "native_reopened": {"model": native_model, "datasheet": native_sheet, "scope": scope, "preview": str(render)}, "source_unchanged": True}
    with report.open("x", encoding="utf-8") as handle:
        json.dump(result, handle, indent=2)
    return result


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--input", type=Path, required=True); p.add_argument("--expected-sha256", required=True)
    reports = p.add_mutually_exclusive_group(required=True); reports.add_argument("--prepared-report", type=Path); reports.add_argument("--previous-report", type=Path)
    p.add_argument("--data", type=Path, required=True); p.add_argument("--ppttc", type=Path, required=True)
    p.add_argument("--output", type=Path, required=True); p.add_argument("--report", type=Path, required=True); p.add_argument("--workdir", type=Path, required=True)
    a = p.parse_args(); print(json.dumps(run(a), indent=2))
