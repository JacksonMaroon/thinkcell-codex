"""Portable create/update route for an authentic think-cell error-bar donor.

This route is deliberately bounded to a donor that already contains a native
Min/Max/Marker error-bar feature. It does not claim insertion into a clean
ordinary line chart. File preparation and verification are safe to run here;
PowerPoint, ppttc and native reopen remain executor-owned operations.

For read-only grading after native execution, use --verify-native FILE,
--data EXPECTED.json, --name NAME, --slide-number N, --expected-sha256 HASH.
The verifier checks actual native model vectors and signed custom cache extent;
native reopen evidence and visual review remain separate requirements.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import math
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import olefile
from lxml import etree as E

HERE = Path(__file__).resolve().parent
def installed_impl_dir() -> Path:
    """Resolve scripts from this installed plugin, independent of CWD.

    When packaged under ``skills/thinkcell-edit/scripts/errorbars`` the
    ancestor walk finds the sibling think-cell implementation directory. The
    explicit environment override is only for lane-local smoke tests.
    """
    requested = os.environ.get("KTC_PRESENTATIONS_PLUGIN_ROOT", "")
    roots = [Path(requested)] if requested else []
    for ancestor in (HERE, *HERE.parents):
        roots.extend([ancestor, ancestor / "scripts"])
    for root in roots:
        candidates = [
            root / "skills/thinkcell-edit/scripts/thinkcell_no_click/implementation",
            root / "thinkcell_no_click/implementation",
        ]
        for impl in candidates:
            if (impl / "prepare_thinkcell_name.py").is_file() and (impl / "replace_ole_stream.ps1").is_file():
                return impl.resolve()
    raise RuntimeError("installed think-cell implementation scripts were not found; set KTC_PRESENTATIONS_PLUGIN_ROOT for lane smoke")


IMPL = installed_impl_dir()
PLUGIN_ROOT = IMPL
sys.path.insert(0, str(IMPL))
import prepare_thinkcell_name as naming  # type: ignore  # noqa: E402
from runtime import powershell, powershell_env  # type: ignore  # noqa: E402


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def _distinct_paths(*paths: Path) -> None:
    """Reject aliasing between source and every generated artifact."""
    # normcase makes the guard effective on Windows when the same file is
    # supplied with different drive/filename casing.
    resolved = [os.path.normcase(str(path.resolve())) for path in paths]
    if len(set(resolved)) != len(resolved):
        raise ValueError("source, data, output, plan, and manifest must be distinct paths")


def presentation_content_type(raw: bytes) -> bytes:
    old = b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
    new = b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
    template_count, presentation_count = raw.count(old), raw.count(new)
    if template_count == 1 and presentation_count == 0:
        return raw.replace(old, new)
    if template_count == 0 and presentation_count == 1:
        return raw
    raise ValueError("source must have exactly one presentation or template main content-type declaration")


def table_payload(categories: list[str], series: dict[str, list[float]]) -> list[list[dict | None]]:
    if not isinstance(categories, list) or not categories or not all(isinstance(c, str) and c.strip() for c in categories):
        raise ValueError("categories must be a nonempty list of nonempty strings")
    if not isinstance(series, dict) or set(series) != {"Min", "Max", "Marker"}:
        raise ValueError("series must be exactly Min, Max, Marker")
    n = len(categories)
    if any(not isinstance(series[k], list) or len(series[k]) != n for k in series):
        raise ValueError("all series must match category count")
    if any(isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value)
           for values in series.values() for value in values):
        raise ValueError("error-bar values must be finite numbers")
    if any(lo > hi for lo, hi in zip(series["Min"], series["Max"])):
        raise ValueError("Min must not exceed Max")
    return [[None] + [{"string": c} for c in categories]] + [
        [{"string": key}] + [{"number": value} for value in series[key]]
        for key in ("Min", "Max", "Marker")
    ]


def donor_seed_data(selected: dict) -> dict:
    """Read the donor's existing vectors for authenticity verification."""
    with olefile.OleFileIO(io.BytesIO(selected["doc"]["ole"])) as ole:
        root = E.fromstring(ole.openstream(["think-cellXML"]).read())
    names = {e.get("id"): e for e in root if e.get("id")}
    table = names[selected["table"].get("id")]
    refs = [e.get("idref") for e in table.find("m_cscdser")]
    labels = [names[names[s].find("m_varsrc").get("idref")].findtext("m_varval") for s in refs]
    if labels != ["Min", "Max", "Marker"]:
        raise ValueError(f"selected donor is not semantic Min/Max/Marker: {labels}")
    rows = {label: [None] * len(table.findall("./ocol/elem")) for label in labels}
    for vr in table.findall("./ocol/elem"):
        vector = names[vr.get("idref")]
        j = int(vector.find("m_nIndexInDataSheet").get("val"))
        for i, scalar_ref in enumerate(vector.findall("./ocol/elem")):
            scalar = names[scalar_ref.get("idref")]
            value = names[scalar.find("m_varsrcAbsolute").get("idref")].find("m_varval").get("val")
            rows[labels[i]][j] = float(value)
    categories = [names[vr.get("idref")].findtext("m_varsrcCategory/m_varval") for vr in table.findall("./ocol/elem")]
    return {"categories": categories, "series": rows}


def _chart_cache(root: E._Element, kind: str) -> tuple[str, list[float]]:
    ref = root.xpath(f'./*[local-name()="{kind}"]/*[local-name()="numRef"]')
    if len(ref) != 1:
        raise ValueError(f"chart series requires one {kind} numeric reference")
    formula = ref[0].xpath('string(./*[local-name()="f"])')
    if not formula:
        raise ValueError(f"chart series {kind} numeric formula is missing")
    return formula, _indexed_cache(ref[0])


def _indexed_cache(reference: E._Element) -> list[float]:
    caches = reference.xpath('./*[local-name()="numCache"]')
    if len(caches) != 1:
        raise ValueError("numeric reference requires one cache")
    counts = caches[0].xpath('./*[local-name()="ptCount"]/@val')
    if len(counts) != 1 or not counts[0].isdigit() or int(counts[0]) < 1:
        raise ValueError("numeric cache requires a positive point count")
    count = int(counts[0]); points = {}
    for point in caches[0].xpath('./*[local-name()="pt"]'):
        index = point.get("idx", "")
        values = point.xpath('./*[local-name()="v"]/text()')
        if not index.isdigit() or int(index) in points or len(values) != 1:
            raise ValueError("numeric cache point indices/values are ambiguous")
        value = float(values[0])
        if not math.isfinite(value):
            raise ValueError("numeric cache values must be finite")
        points[int(index)] = value
    if len(points) != count or set(points) != set(range(count)):
        raise ValueError("numeric cache indices must cover the contiguous point count")
    return [points[index] for index in range(count)]


def _error_cache(error: E._Element) -> tuple[str, list[float]]:
    fields = {field: error.xpath(f'./*[local-name()="{field}"]/@val')
              for field in ("errDir", "errValType", "errBarType")}
    if fields != {"errDir": ["x"], "errValType": ["cust"], "errBarType": ["plus"]}:
        raise ValueError("error-bar carrier requires custom X-direction plus-only extents")
    if error.xpath('./*[local-name()="minus"]'):
        raise ValueError("minus error-bar extents are outside the authentic plus-only profile")
    return _chart_cache(error, "plus")


def verify_authentic_donor(path: Path, name: str, slide_number: int, expected: dict) -> dict:
    """Verify model ownership plus signed physical extent cache without lane fixtures."""
    table_payload(expected["categories"], expected["series"])
    _, candidates, _ = naming.inventory(path.read_bytes())
    matches = [c for c in candidates if c["owner_name"] == name and c["doc"]["slide_number"] == slide_number]
    if len(matches) != 1:
        raise ValueError("named error-bar chart on selected slide is missing or ambiguous")
    selected = naming.choose(matches, slide_number=slide_number, shape_tag=matches[0]["frames"][0]["shape_tag"])
    actual_data = donor_seed_data(selected)
    if actual_data != expected:
        raise ValueError("native error-bar model data differs from expected categories/series")
    chart_part = selected["frames"][0].get("native_chart_part")
    if not chart_part:
        raise ValueError("selected error-bar donor has no native chart cache")
    with zipfile.ZipFile(path) as z:
        ole_part = selected["doc"]["part"]
        with olefile.OleFileIO(io.BytesIO(z.read(ole_part))) as ole:
            model = E.fromstring(ole.openstream(["think-cellXML"]).read())
        ids = {e.get("id"): e for e in model if e.get("id")}
        charts = [e for e in model.findall("CSequenceChartSE") if e.findtext("m_strName") == name]
        if len(charts) != 1:
            raise ValueError(f"named native chart count {len(charts)} != 1")
        chart = charts[0]
        table_ref = chart.find("m_dtable")
        table = ids.get(table_ref.get("idref")) if table_ref is not None else None
        if table is None or table.tag != "CSequenceChartDataTable" or table.findtext("m_strName") != name:
            raise ValueError("named chart/table ownership closure missing")
        series_refs = [e.get("idref") for e in table.findall("./m_cscdser/elem")]
        labels = [ids[ids[s].find("m_varsrc").get("idref")].findtext("m_varval") for s in series_refs]
        if labels != ["Min", "Max", "Marker"]:
            raise ValueError(f"semantic series identity is {labels}, expected Min/Max/Marker")
        ranges = model.findall("CSequenceChartSeriesRange")
        if len(ranges) != 1:
            raise ValueError(f"serialized native range count {len(ranges)} != 1")
        rid = ranges[0].get("id")
        if rid not in [e.get("idref") for e in table.findall("./m_cscserrange/elem")]:
            raise ValueError("data table does not own serialized error-bar range")
        range_refs = [e.get("idref") for e in ranges[0].findall("./m_cscdseries/elem")]
        range_labels = [ids[ids[s].find("m_varsrc").get("idref")].findtext("m_varval") for s in range_refs]
        if set(range_labels) != {"Min", "Max"}:
            raise ValueError(f"range semantic labels are {range_labels}")
        with zipfile.ZipFile(path) as z2:
            drawing = E.fromstring(z2.read(chart_part))
        series = drawing.xpath('.//*[local-name()="plotArea"]/*[local-name()="lineChart" or local-name()="scatterChart"]/*[local-name()="ser"]')
        caches = [_chart_cache(s, "xVal") for s in series]
        category_positions = [index + .5 for index in range(len(expected["categories"]))]
        if any(_chart_cache(s, "yVal")[1] != category_positions for s in series):
            raise ValueError("error-bar category Y positions differ from the bounded donor profile")
        want = {tuple(float(v) for v in expected["series"][k]): k for k in ("Min", "Max", "Marker")}
        semantic = {tuple(values): (formula, s) for (formula, values), s in zip(caches, series)}
        if len(want) != 3 or len(semantic) != 3 or len(series) != 3:
            raise ValueError("Min/Max/Marker cache vectors must be distinct and unambiguous")
        if set(semantic) != set(want):
            raise ValueError("drawing chart caches do not contain exact Min/Max/Marker donor values")
        min_formula = next(formula for (values, (formula, _)) in semantic.items() if want[values] == "Min")
        max_formula = next(formula for (values, (formula, _)) in semantic.items() if want[values] == "Max")
        marker_formula = next(formula for (values, (formula, _)) in semantic.items() if want[values] == "Marker")
        err_series = [s for s in series if s.xpath('./*[local-name()="errBars"]')]
        if len(err_series) != 1:
            raise ValueError(f"drawing native error-bar carrier count {len(err_series)} != 1")
        err_owner_formula, _ = _chart_cache(err_series[0], "xVal")
        if err_owner_formula not in {min_formula, max_formula}:
            raise ValueError("error-bar carrier is not attached to semantic Min or Max")
        err = err_series[0].xpath('./*[local-name()="errBars"]')[0]
        err_formula, err_values = _error_cache(err)
        lo, hi = expected["series"]["Min"], expected["series"]["Max"]
        signed_expected = [float(h) - float(l) for l, h in zip(lo, hi)] if err_owner_formula == min_formula else [float(l) - float(h) for l, h in zip(lo, hi)]
        if err_values != signed_expected:
            raise ValueError(f"signed custom extent mismatch: {err_values} vs {signed_expected}")
    return {"model_range_refs": range_labels, "model_range_id": rid,
            "chart_part": chart_part, "min_formula": min_formula,
            "max_formula": max_formula, "marker_formula": marker_formula,
            "error_formula": err_formula, "error_owner_formula": err_owner_formula,
            "signed_extent": err_values,
            "error_semantics": "Max-Min attached to Min" if err_owner_formula == min_formula else "Min-Max attached to Max",
        "marker_fixed_in_seed": True}


def verify_native(path: Path, name: str, slide_number: int, expected: dict, expected_sha256: str) -> dict:
    """Grade a native-saved same/changed-data artifact without touching Office."""
    before = sha(path)
    if before != expected_sha256.upper():
        raise ValueError("native artifact SHA-256 mismatch")
    verification = verify_authentic_donor(path, name, slide_number, expected)
    if sha(path) != before:
        raise RuntimeError("native artifact changed during verification")
    return {"status": "ERRORBAR_MODEL_AND_CACHE_GATES_PASS", "input": str(path.resolve()),
            "input_sha256": before, "source_unchanged": True, "verification": verification,
            "remaining_review": ["native save/reopen evidence", "target-slide visual review"]}


def build(source: Path, output: Path, plan: Path, name: str, data: dict,
          slide_number: int, shape_id: int | None, shape_tag: str | None,
          expected_source_sha256: str | None = None) -> dict:
    if not source.exists():
        raise FileNotFoundError(source)
    _distinct_paths(source, output, plan)
    if output.exists() or plan.exists():
        raise FileExistsError("refusing to overwrite route outputs")
    categories = data["categories"]
    series = data["series"]
    payload_table = table_payload(categories, series)
    source_hash = sha(source)
    if expected_source_sha256 is not None:
        expected = expected_source_sha256.strip().upper()
        if len(expected) != 64 or any(c not in "0123456789ABCDEF" for c in expected):
            raise ValueError("expected source SHA-256 must be 64 hexadecimal characters")
        if source_hash != expected:
            raise ValueError(f"source SHA-256 changed: {source_hash} != {expected}")
    docs, candidates, names = naming.inventory(source.read_bytes())
    selected = naming.choose(candidates, slide_number=slide_number, shape_id=shape_id, shape_tag=shape_tag)
    seed = donor_seed_data(selected)
    table_payload(seed["categories"], seed["series"])
    if len(categories) != len(seed["categories"]):
        raise ValueError("error-bar update must preserve donor category count")
    if series["Marker"] != seed["series"]["Marker"]:
        raise ValueError("error-bar update must preserve the donor Marker series")
    chosen, reused = naming.name_plan(selected, names, source_hash, name)
    if chosen != name:
        raise ValueError("automation name is not the requested semantic name")
    model_bytes, name_changes = naming.rewrite(selected, name, reused)
    with tempfile.TemporaryDirectory(prefix="errorbar_route_", dir=str(output.parent)) as td:
        temp = Path(td)
        carrier, model = temp / "carrier.bin", temp / "model.xml"
        carrier.write_bytes(selected["doc"]["ole"])
        model.write_bytes(model_bytes)
        helper = IMPL / "replace_ole_stream.ps1"
        result = subprocess.run(
            [powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(helper),
             "-StoragePath", str(carrier), "-StreamBytesPath", str(model)],
            capture_output=True, text=True, env=powershell_env(), check=False, timeout=60,
        )
        if result.returncode:
            raise RuntimeError(result.stderr[:2000])
        changed_ole = carrier.read_bytes()
        with zipfile.ZipFile(source) as src, zipfile.ZipFile(output, "w") as dst:
            dst.comment = src.comment
            for entry in src.infolist():
                if entry.filename == selected["doc"]["part"]:
                    payload = changed_ole
                elif entry.filename == "[Content_Types].xml":
                    payload = presentation_content_type(src.read(entry.filename))
                else:
                    payload = src.read(entry.filename)
                dst.writestr(entry, payload)
    if sha(source) != source_hash:
        raise RuntimeError("source changed while preparing route")
    # Verify the prepared copy directly. This certifies the donor's model
    # range and signed native custom extent cache without a lane fixture;
    # requested update data is intentionally left for ppttc/native work.
    verification = verify_authentic_donor(output, name, slide_number, seed)
    payload = [{"template": str(output.resolve()),
                "data": [{"name": name, "table": payload_table}]}]
    plan.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    return {
        "status": "PREPARED_AUTHENTIC_DONOR_CREATE_UPDATE_ROUTE",
        "source": str(source.resolve()), "source_sha256": source_hash,
        "plugin_root": str(PLUGIN_ROOT),
        "candidate": str(output.resolve()), "candidate_sha256": sha(output),
        "plan": str(plan.resolve()), "plan_sha256": sha(plan),
        "automation_name": name,
        "target": {"slide_number": slide_number, "shape_id": shape_id,
                   "shape_tag": selected["frames"][0]["shape_tag"],
                   "active_document_part": selected["doc"]["part"],
                   "logical_chart_id": selected["owner"].get("id"),
                   "logical_table_id": selected["table"].get("id")},
        "semantic_series": ["Min", "Max", "Marker"],
        "native_feature": "existing CSequenceChartSeriesRange pairing Min and Max",
        "portable_scope": "authentic installed donor import and named data update",
        "clean_line_insertion": "unverified and intentionally unsupported",
        "data": data,
        "seed_verification": verification,
        "name_changes": name_changes,
        "verification": {
            "offline": "test_candidate.py verifies the authentic donor model ownership, exact donor vectors, signed Max-Min or Min-Max custom extent cache, and Marker stability",
            "native": "executor must run ppttc, native reopen/save, target-slide render, and changed-data readback",
            "source_preservation": "source hash must remain source_sha256; native scope report must pass source and sibling gates",
        },
    }


def main() -> None:
    if "--verify-native" in sys.argv:
        ap = argparse.ArgumentParser(description="Read-only verification of a native-saved error-bar update")
        ap.add_argument("--verify-native", type=Path, required=True)
        ap.add_argument("--data", type=Path, required=True)
        ap.add_argument("--name", required=True)
        ap.add_argument("--slide-number", type=int, required=True)
        ap.add_argument("--expected-sha256", required=True)
        args = ap.parse_args()
        data = json.loads(args.data.read_text(encoding="utf-8-sig"))
        print(json.dumps(verify_native(args.verify_native, args.name, args.slide_number, data,
                                       args.expected_sha256), indent=2))
        return
    ap = argparse.ArgumentParser()
    ap.add_argument("--verify-native", type=Path,
                    help="Alternative read-only mode: supply data, name, slide-number and expected-sha256 only")
    ap.add_argument("--source", type=Path, required=True)
    ap.add_argument("--data", type=Path, required=True)
    ap.add_argument("--output", type=Path, required=True)
    ap.add_argument("--plan", type=Path)
    ap.add_argument("--name", default="TC_AUTO_ERRORBAR_PORTABLE")
    ap.add_argument("--slide-number", type=int, required=True)
    selector = ap.add_mutually_exclusive_group(required=True)
    selector.add_argument("--shape-id", type=int)
    selector.add_argument("--shape-tag")
    ap.add_argument("--manifest", type=Path)
    ap.add_argument("--expected-sha256", "--expected-source-sha256", dest="expected_source_sha256", required=True)
    args = ap.parse_args()
    source, output = args.source.resolve(), args.output.resolve()
    plan = (args.plan or output.with_suffix(".ppttc")).resolve()
    manifest_path = (args.manifest or output.with_suffix(".manifest.json")).resolve()
    data_path = args.data.resolve()
    _distinct_paths(source, data_path, output, plan, manifest_path)
    if manifest_path.exists():
        raise FileExistsError("refusing to overwrite route manifest")
    data = json.loads(data_path.read_text(encoding="utf-8"))
    manifest = build(source, output, plan, args.name, data, args.slide_number, args.shape_id,
                     args.shape_tag, args.expected_source_sha256)
    manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    if sha(source) != manifest["source_sha256"]:
        raise RuntimeError("source changed before route completion")
    print(json.dumps(manifest, indent=2))


if __name__ == "__main__":
    main()
