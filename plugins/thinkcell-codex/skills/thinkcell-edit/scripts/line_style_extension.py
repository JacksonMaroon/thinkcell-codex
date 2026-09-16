"""Native line series colors through the multi-chart update pipeline.

The installed runner only maps ``c:barChart`` caches.  A native line donor has
the same CSequenceChartSE model and official datasheet-fill setting, but its
visible series live under ``c:lineChart`` and its colors are line/marker fills.
This adapter broadens only the cache readback and mapping needed for that
simple, single-axis line topology.  It never writes the PowerPoint cache.
"""
from __future__ import annotations

import argparse
import copy
import importlib.util
import io
import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree as E

def _discover_script_dir():
    """Find staged think-cell scripts from an env override or this file's tree."""
    candidates = []
    override = os.environ.get("THINKCELL_NATIVE_SCRIPTS")
    if override:
        candidates.append(Path(override).expanduser())
    here = Path(__file__).resolve().parent
    for ancestor in (here, *here.parents):
        candidates.extend((ancestor / "scripts", ancestor / "staged-plugin" / "skills" / "thinkcell-edit" / "scripts", ancestor / "skills" / "thinkcell-edit" / "scripts"))
    for candidate in candidates:
        if (candidate / "multi_chart_update.py").is_file() and (candidate / "thinkcell_no_click" / "implementation").is_dir():
            return candidate
    raise FileNotFoundError("Cannot locate native think-cell scripts from this adapter's directory")


STAGED = _discover_script_dir()
ORIGINAL = STAGED / "multi_chart_update.py"
sys.path.insert(0, str(STAGED / "thinkcell_no_click" / "implementation"))
from prepare_thinkcell_name import inventory, sha, streams, xml
spec = importlib.util.spec_from_file_location("staged_multi_chart_update", ORIGINAL)
multi = importlib.util.module_from_spec(spec)
spec.loader.exec_module(multi)

NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def need(ok, message):
    if not ok:
        raise ValueError(message)


def _value(node):
    return None if node is None else node.get("val")


def _plot(root):
    plots = []
    for tag in ("barChart", "lineChart", "areaChart"):
        plots.extend(root.findall(f".//c:{tag}", NS))
    need(len(plots) == 1, "Line styling requires one simple native plot")
    plot = plots[0]
    kind = E.QName(plot).localname
    need(kind == "lineChart", "Line styling requires a native lineChart cache")
    need(len(root.findall(".//c:valAx", NS)) == 1, "Line styling requires one value axis")
    return plot


def cache_semantics(path, candidate):
    root = _cache_root(path, candidate)
    plot = _plot(root)
    return {"bar_dir": None, "grouping": "line", "vary_colors": None,
            "series_count": len(plot.findall("c:ser", NS))}


def _cache_root(path, candidate):
    part = candidate["frames"][0]["native_chart_part"]
    need(part, "Target has no native chart cache part")
    import zipfile
    with zipfile.ZipFile(path) as archive:
        return E.fromstring(archive.read(part))


def cache_value_vectors(root):
    plot = _plot(root)
    vectors = []
    for series in plot.findall("c:ser", NS):
        points = series.findall("./c:val/c:numRef/c:numCache/c:pt", NS)
        need(points, "Visible line series has no numeric cache")
        indexed = {int(point.get("idx")): float(point.find("c:v", NS).text) for point in points}
        need(set(indexed) == set(range(len(indexed))), "Visible line cache indices are sparse or duplicated")
        vectors.append([indexed[index] for index in range(len(indexed))])
    return vectors


def _line_fill_route(root, candidate):
    owner = candidate["owner"]
    need(owner.tag == "CSequenceChartSE", "Line styling requires an ordinary sequence chart")
    primary = owner.findall("m_daxisPrimaryValue")
    secondary = owner.findall("m_daxisSecondaryValue")
    need(len(primary) == 1 and primary[0].get("idref") not in {None, "0"}, "Line styling requires one primary value axis")
    need(len(secondary) == 1 and secondary[0].get("idref") == "0", "Line styling excludes secondary axes")
    axis = candidate["doc"]["ids"].get(primary[0].get("idref"))
    need(axis is not None and axis.tag == "CSequenceChartDataAxis" and not axis.findall("m_cdaxisbreak/elem"), "Line styling excludes axis breaks")
    model = multi.model_of(candidate)
    return multi.match_unique_series(cache_value_vectors(root), model["series_names"], model["series_values"])


def simple_sequence_contract(candidate, request, flexible):
    """Validate the line donor's blank reserved row without inventing totals.

    The generic sequence contract treats a blank 100% row as a calculated
    extent row.  On this native line donor the model deliberately retains its
    fixed 100 extents while the reserved datasheet row stays blank, so that
    rule would reject a semantically unchanged line chart before JSON runs.
    """
    baseline = multi.model_of(candidate)
    source, orientation, ds = multi.cells_matrix(candidate)
    expected, matrix = request["expected_model"], request["matrix"]
    need(set(expected) in ({"categories", "series_names", "series_values"}, {"categories", "series_names", "series_values", "category_extents"}), "Expected model fields are invalid")
    need(len(matrix) == len(source) and all(len(row) == len(source[0]) for row in matrix), "Line matrix dimensions must remain fixed")
    need(multi.labels_equal(matrix[0][1:], expected["categories"]), "Line category row differs from expected model")
    need(multi.labels_equal(source[0][1:], baseline["categories"]), "Source line category row disagrees with model")
    names, values = expected["series_names"], expected["series_values"]
    need(names == baseline["series_names"] and len(values) == len(names), "Line series dimensions must remain fixed")
    need(len(expected["categories"]) == len(baseline["categories"]), "Line category dimensions must remain fixed")
    need(all(len(row) == len(names) * 0 + len(expected["categories"]) for row in values), "Line series values dimensions are invalid")
    need(all(v is None or (isinstance(v, (int, float)) and not isinstance(v, bool)) for row in values for v in row), "Line series values must be finite numbers/null")
    need("category_extents" in expected and multi.equal(expected["category_extents"], baseline["category_extents"]), "Line extents must retain native fixed values")
    need(matrix[1][0] is None and all(v is None for v in matrix[1][1:]), "Line reserved 100% row must remain blank")
    for index, name in enumerate(names, 2):
        need(matrix[index][0] == name and multi.equal(matrix[index][1:], values[index - 2]), "Line series row differs from expected model")
        need(source[index][0] == baseline["series_names"][index - 2], "Source line series row order changed")
    need(not flexible, "Line style canary does not support count changes")
    return {"family": "sequence", "datasheet_orientation": orientation, "source_dimensions": [len(source), len(source[0])],
            "requested_dimensions": [len(matrix), len(matrix[0])], "count_change": False,
            "only_optional_row": "blank-100%=", "datasource": ds}


def _fill_series_names(path, candidate, contract=None):
    root = _cache_root(path, candidate)
    return root, _line_fill_route(root, candidate), "unique_vectors_line"


def _rgb(node):
    return node.xpath("./a:solidFill/a:srgbClr/@val", namespaces=NS)


def _line_rgb(series):
    values = series.xpath("./c:spPr/a:ln/a:solidFill/a:srgbClr/@val", namespaces=NS)
    need(len(values) == 1, "Visible line series has no exact RGB line fill")
    return values[0].upper()


def _point_rgb(point):
    # Line caches store point appearance under c:dPt/c:marker/c:spPr.
    values = point.xpath("./c:marker/c:spPr/a:solidFill/a:srgbClr/@val", namespaces=NS)
    need(len(values) == 1, "Visible line point has no exact RGB marker fill")
    return values[0].upper()


def cache_fills(path, candidate, contract=None):
    root, names, _ = _fill_series_names(path, candidate, contract)
    plot = _plot(root)
    series_nodes = plot.findall("c:ser", NS)
    need(len(series_nodes) == len(names), "Visible series count differs from native model")
    result = {}
    for name, series in zip(names, series_nodes):
        points = {}
        for point in series.findall("c:dPt", NS):
            idx = point.find("c:idx", NS)
            if idx is None:
                continue
            # The donor's initial points use theme colors and cannot be used as
            # exact RGB preservation evidence; requested fills are RGB after
            # official regeneration and are read back here.
            values = point.xpath("./c:marker/c:spPr/a:solidFill/a:srgbClr/@val", namespaces=NS)
            if len(values) == 1:
                points[int(idx.get("val"))] = values[0].upper()
        result[name] = {"series": _line_rgb(series), "points": points}
    return result


def effective_appearance(prepared, candidate, spec, contract):
    # An explicit complete line mapping is required.  Merging theme-colored
    # donor markers would make the pre-regeneration readback ambiguous.
    if "appearance" in spec:
        requested = spec["appearance"]
        names = spec["data"]["expected_model"]["series_names"]
        need(set(requested["series_fills"]) == set(names), "Line styling requires an explicit fill for every series")
        return requested, "requested_line"
    if multi.datasheet_fill_enabled(candidate):
        raise ValueError("Line styling without explicit appearance is untested")
    return None, "not_enabled"


def _set_rgb(parent, path, color):
    col = parent.find(path)
    need(col is not None, "Line model style slot is missing: " + path)
    idx = col.find("m_msothmcolidx")
    if idx is None:
        idx = E.SubElement(col, "m_msothmcolidx")
    idx.set("val", "0")
    rgb = col.find("m_rgb")
    if rgb is None:
        rgb = E.SubElement(col, "m_rgb")
    value = color.removeprefix("#").upper()
    rgb.attrib.clear()
    rgb.attrib.update({"r": value[0:2], "g": value[2:4], "b": value[4:6]})


def _style_model_bytes(raw, appearance):
    """Change only genuine native line and marker color fields."""
    docs, charts, _ = inventory(raw)
    need(len(charts) == 1, "Line style preparation requires one chart")
    chart = charts[0]
    root = xml(chart["doc"]["streams"][("think-cellXML",)])
    ids = {node.get("id"): node for node in root if node.get("id")}
    series_nodes = root.findall("CSequenceChartDataSeries")
    by_name = {}
    for series in series_nodes:
        source = series.find("m_varsrc")
        variable = ids.get(source.get("idref")) if source is not None else None
        name = variable.findtext("m_varval") if variable is not None else None
        if name:
            by_name[name] = series
    need(set(appearance["series_fills"]) == set(by_name), "Line model series names are not a complete unique set")
    for name, color in appearance["series_fills"].items():
        _set_rgb(by_name[name], "m_linestyle/m_col", color)

    vectors = []
    table = chart["table"]
    for ref in table.findall("ocol/elem"):
        vector = ids.get(ref.get("idref"))
        if vector is not None and vector.find("m_nIndexInDataSheet") is not None:
            vectors.append(vector)
    order = list(by_name)
    for name, colors in appearance.get("point_fills", {}).items():
        position = order.index(name)
        for index, color in enumerate(colors):
            if color is None:
                continue
            matches = [v for v in vectors if v.find("m_nIndexInDataSheet").get("val") == str(index)]
            need(len(matches) == 1, "Line point category does not map to one native vector")
            scalars = matches[0].findall("ocol/elem")
            need(len(scalars) > position, "Line point native scalar slot is missing")
            scalar = ids.get(scalars[position].get("idref"))
            need(scalar is not None, "Line point native scalar is missing")
            _set_rgb(scalar, "m_markerprops/m_col", color)
    return E.tostring(root, encoding="utf-8")


def prepare_model_style(input_path, output_path, appearance):
    """Make a copy with native model colors changed, then regenerate later."""
    source, output = Path(input_path), Path(output_path)
    need(source.resolve() != output.resolve() and not output.exists(), "Style preparation needs a distinct output")
    raw = source.read_bytes()
    docs, charts, _ = inventory(raw)
    need(len(charts) == 1, "Style preparation requires one chart")
    doc = charts[0]["doc"]
    before = doc["streams"][("think-cellXML",)]
    after = _style_model_bytes(raw, appearance)
    with tempfile.TemporaryDirectory(prefix="tc_line_style_", dir=output.parent) as temp:
        temp = Path(temp)
        carrier, payload = temp / "carrier.bin", temp / "model.xml"
        carrier.write_bytes(doc["ole"])
        payload.write_bytes(after)
        command = [multi.powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File",
                   str(STAGED / "thinkcell_no_click/implementation/replace_ole_stream.ps1"),
                   "-StoragePath", str(carrier), "-StreamBytesPath", str(payload)]
        result = subprocess.run(command, capture_output=True, text=True, env=multi.powershell_env(), timeout=60)
        need(result.returncode == 0, "Line model style preparation failed: " + result.stderr[-500:])
        changed = carrier.read_bytes()
        updated = streams(changed)
        need(set(updated) == set(doc["streams"]), "Model style preparation changed OLE stream inventory")
        need(all(value == updated[key] for key, value in doc["streams"].items() if key != ("think-cellXML",)), "Non-model OLE stream changed")
        with zipfile.ZipFile(io.BytesIO(raw)) as zin, zipfile.ZipFile(output, "x") as zout:
            zout.comment = zin.comment
            for item in zin.infolist():
                zout.writestr(copy.copy(item), changed if item.filename == doc["part"] else zin.read(item.filename))
    return {"status": "LINE_MODEL_STYLE_PREPARED", "source_sha256": sha(raw), "output_sha256": sha(output.read_bytes()),
            "only_native_line_and_marker_colors_changed": True}


def install():
    multi.simple_sequence_contract = simple_sequence_contract
    original_prepare_all = multi.prepare_all

    def prepare_all_with_line_fill(src, digest, stage, chosen):
        current, manifests = original_prepare_all(src, digest, stage, chosen)
        import enable_datasheet_fill
        chart = chosen[0][0]
        selection = {"slide_number": 1, "shape_tag": chart["frames"][0]["shape_tag"]}
        fill_plan = enable_datasheet_fill.make_plan(current, selection)
        enabled = stage / "datasheet-fill-enabled.pptx"
        fill_result = enable_datasheet_fill.prepare(current, enabled, fill_plan)
        style_output = stage / "line-model-styled.pptx"
        style_result = prepare_model_style(enabled, style_output, chosen[0][1]["appearance"])
        manifests = list(manifests) + [{"datasheet_fill": fill_plan, "result": fill_result},
                                       {"line_model_style": style_result}]
        return style_output, manifests

    multi.prepare_all = prepare_all_with_line_fill
    multi.cache_semantics = cache_semantics
    multi.cache_semantics_or_none = lambda path, candidate: cache_semantics(path, candidate)
    multi._cache_root = _cache_root
    multi.cache_value_vectors = cache_value_vectors
    multi._fill_series_names = _fill_series_names
    multi.cache_fills = cache_fills
    multi.effective_appearance = effective_appearance
    return multi


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--expected-sha256", required=True)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--prepare", action="store_true")
    parser.add_argument("--execute", action="store_true")
    parser.add_argument("--ppttc")
    args = parser.parse_args()
    runner = install()
    report = runner.run(args)
    runner.dump(args.report, report)
    print(json.dumps({"status": report["status"], "report": str(args.report.resolve())}))


if __name__ == "__main__":
    main()
