"""Execute the bounded native axis-break insertion or proven repeat update.

This module lives under <bundled skill>/scripts and uses only its sibling
helpers.  It is copy-only: input, output and report must be distinct paths.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import math
import subprocess
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path[:0] = [str(HERE), str(HERE / "thinkcell_no_click" / "implementation")]

from chart_geometry import inventory  # noqa: E402
from office_operation_lock import serialized_office, run_locked_subprocess  # noqa: E402
from runtime import powershell, powershell_env  # noqa: E402
from thinkcell_no_click.implementation.prepare_thinkcell_name import link_contract  # noqa: E402
from thinkcell_no_click.implementation.read_named_datasheet import read as read_sheet  # noqa: E402
from thinkcell_no_click.implementation.update_thinkcell_json import equal, model_of  # noqa: E402
from audit_thinkcell_integrity import chart_details, inspect_presentation, logical_slides  # noqa: E402
import experimental_axis_break_insert as prepare  # noqa: E402

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"p": P, "r": R}
Q = lambda name: "{" + R + "}" + name
FIELDS = ("m_pptshpLow", "m_pptshpHigh", "m_pptshpBody")


def need(condition, message):
    if not condition:
        raise RuntimeError(message)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def exclusive_write(path: Path, payload: bytes):
    with path.open("xb") as handle:
        handle.write(payload)


def target(path: Path, plan: dict):
    docs, charts, _ = inventory(path.read_bytes())
    need(len(docs) == len(charts) == 1, "requires exactly one logical slide, carrier, and native chart")
    wanted = plan["targets"][0]
    choices = [c for c in charts if c["exact"] and c["owner_name"] == wanted["name"]]
    need(len(choices) == 1, "target automation name did not resolve uniquely")
    chart = choices[0]
    need(chart["owner"].tag == "CSequenceChartSE", "unsupported chart subtype")
    ect = chart["owner"].find("m_ect")
    need(ect is not None and ect.get("val") == "1", "requires clustered-column native model subtype")
    need(chart["frames"][0]["shape_tag"] == wanted["selector"]["shape_tag"], "target shape tag changed")
    link_contract(chart)
    cache_part = chart["frames"][0]["native_chart_part"]
    need(cache_part, "target has no native chart cache")
    with zipfile.ZipFile(path) as z:
        need(len(logical_slides(z)) == 1, "requires exactly one logical slide")
        need(cache_part in z.namelist(), "native chart cache is missing")
        cache = chart_details(z.read(cache_part))
    subtype = cache.get("subtypes", [])
    need(len(subtype) == 1 and subtype[0].get("type") == "barChart" and subtype[0].get("bar_direction") == "col" and subtype[0].get("grouping") == "clustered", "requires clustered-column native cache")
    return chart, wanted


def profile(plan: dict):
    need(set(plan) == {"targets"} and len(plan["targets"]) == 1, "one target plan is required")
    data = plan["targets"][0]["data"]
    matrix, expected = data["matrix"], data["expected_model"]
    need(len(matrix) == 5 and all(len(row) == 5 for row in matrix), "requires the tested 3-series x 4-category matrix")
    need(len(expected["categories"]) == 4 and len(expected["series_names"]) == 3, "requires exactly 4 categories and 3 series")
    values = expected["series_values"]
    need(len(values) == 3 and all(len(row) == 4 for row in values), "requires exactly 12 scalar values")
    scalars = [v for row in values for v in row]
    need(all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) and v >= 0 for v in scalars), "requires finite nonnegative numeric values")
    maximum = max(scalars)
    need(scalars.count(maximum) == 1 and maximum > max(v for v in scalars if v != maximum), "requires one unique outlier")
    need(len(expected["category_extents"]) == 4, "category extents must cover all four categories")
    need(matrix[0][1:] == expected["categories"] and [row[0] for row in matrix[2:]] == expected["series_names"], "matrix labels disagree with expected model")
    need(all(equal(row[1:], values[i]) for i, row in enumerate(matrix[2:])), "matrix values disagree with expected model")
    need(all(equal(total, expected["category_extents"][i]) for i, total in enumerate([sum(row[i] for row in values) for i in range(4)])), "category extents disagree with 3-series values")
    return data


def relationships(z: zipfile.ZipFile, part: str):
    relpart = str(Path(part).parent / "_rels" / (Path(part).name + ".rels")).replace("\\", "/")
    root = __import__("lxml.etree", fromlist=["etree"]).fromstring(z.read(relpart))
    import posixpath
    base = str(Path(part).parent).replace("\\", "/")
    return {x.get("Id"): posixpath.normpath(posixpath.join(base, x.get("Target"))) for x in root}


def physical_tags(raw: bytes, names: list[str]):
    from lxml import etree as E
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        slide = E.fromstring(z.read("ppt/slides/slide1.xml"))
        rels = relationships(z, "ppt/slides/slide1.xml")
        found = []
        for node in list(slide.find("p:cSld/p:spTree", NS)):
            for tag in node.findall(".//p:tags", NS):
                part = rels.get(tag.get(Q("id")))
                if part not in z.namelist():
                    continue
                vals = [x.get("val") for x in E.fromstring(z.read(part)) if x.get("name", "").upper() == "THINKCELLSHAPEDONOTDELETE"]
                hit = set(vals).intersection(names)
                need(len(hit) <= 1, "one physical node maps multiple break tags")
                found.extend(hit)
        need(len(found) == len(names) and set(found) == set(names), "linked physical break parts are missing")


def specialized_integrity(path: Path) -> dict:
    """Use the native break scope: ordinary cache parity is inapplicable.

    A broken axis transforms the visible cache.  The normal model-to-cache
    comparison therefore fails even when the visible chart values are right.
    Preserve the general package/content audit and explicitly record only the
    three cache assertions that are inapplicable to this native topology.
    """
    audit = inspect_presentation(path, strict_parity=False)
    excluded = {"no_model_visible_chart_mismatch", "strict_parity_available", "strict_parity_pass"}
    failures = [key for key, value in audit["assertions"].items() if not value and key not in excluded]
    need(not failures, "integrity checks failed: " + ", ".join(failures))
    return {"integrity_assertions_pass": True, "native_cache_parity": "known_axis_break_cache_transform", "excluded_inapplicable_checks": sorted(excluded)}


def verify(path: Path, plan: dict, expected_shape_count: int) -> dict:
    data = profile(plan)
    chart, wanted = target(path, plan)
    actual = model_of(chart)
    expected = data["expected_model"]
    for key in ("categories", "series_names", "series_values", "category_extents"):
        need(equal(actual.get(key), expected[key]), "exact model mismatch: " + key)
    sheets = read_sheet(path, 1, chart["owner_name"])["sheets"]
    cells = {(x["row"], x["column"]): x["value"] for x in sheets[0]["nonempty_cells"]}
    wanted_cells = {(r + 1, c + 1): value for r, row in enumerate(data["matrix"]) for c, value in enumerate(row) if value not in (None, "")}
    need(cells == wanted_cells, "exact datasource cells mismatch")
    ids = chart["doc"]["ids"]
    axis_ref = chart["owner"].find("m_daxisPrimaryValue")
    axis = ids[axis_ref.get("idref")]
    refs = axis.findall("m_cdaxisbreak/elem")
    need(len(refs) == 1 and refs[0].get("idref") in ids, "one native axis break is required")
    br = ids[refs[0].get("idref")]
    fraction = br.find("m_fFraction")
    need(br.tag == "CDataAxisBreak" and br.find("m_bUser").get("val") == "1" and fraction is not None and math.isclose(float(fraction.get("val")), 0.2, rel_tol=0, abs_tol=1e-12), "break is not the proved user-owned 20% native state")
    shapes = [ids[x.get("idref")] for x in br.findall("m_cpptbreakshp/elem")]
    need(len(shapes) == expected_shape_count and all(x.tag == "CPPTBreakShape" for x in shapes), "unexpected generated break-shape topology")
    names = [x.find(field + "/m_bstrShapeName").text for x in shapes for field in FIELDS]
    need(len(names) == 3 * expected_shape_count and len(set(names)) == len(names) and all(names), "incomplete linked break tags")
    physical_tags(path.read_bytes(), names)
    intervals = axis.findall("m_vecintvlOrdinal")
    if expected_shape_count == 2:
        need(len(intervals) == 1 and intervals[0].get("length") == "3", "unexpected prepared seed interval topology")
        return {"status": "PREPARED_SEED_NATIVE_REGENERATION_REQUIRED", "exact_model": True, "exact_datasheet": True,
                "native_break_shape_count": 2, "physical_tag_count": 6, "fraction": float(fraction.get("val"))}
    need(len(intervals) == 1 and intervals[0].get("length") == "12", "unexpected regenerated interval topology")
    need(br.find("m_nOrdinal") is not None and br.find("m_nOrdinal").get("val") == "0", "unexpected selected break interval")
    gap = intervals[0].find("elem")
    low, high = float(gap.find("begin").get("val")), float(gap.find("end").get("val"))
    crossed = [value for row in expected["series_values"] for value in row if value > low and value <= high]
    need(len(crossed) == 1 and equal(crossed[0], max(v for row in expected["series_values"] for v in row)), "axis gap is not crossed by exactly the unique outlier")
    body = shapes[0].find("m_pptshpBody/m_rectPPTShape")
    low_rect = shapes[0].find("m_pptshpLow/m_rectPPTShape")
    high_rect = shapes[0].find("m_pptshpHigh/m_rectPPTShape")
    rect = lambda node: {k: int(node.get(k)) for k in ("left", "top", "right", "bottom")}
    body, low_rect, high_rect = rect(body), rect(low_rect), rect(high_rect)
    need(body["left"] < body["right"] and body["top"] < body["bottom"] and low_rect["left"] == high_rect["left"] == body["left"] and low_rect["right"] == high_rect["right"] == body["right"] and body["top"] <= low_rect["top"] < high_rect["top"] <= body["bottom"], "break rectangles are not a gap within one crossed bar")
    integrity = specialized_integrity(path)
    return {"exact_model": True, "exact_datasheet": True, "integrity": integrity, "chart": {"name": chart["owner_name"], "shape_tag": chart["frames"][0]["shape_tag"], "subtype": chart["owner"].tag, "ect": "1", "cache": "barChart/col/clustered"}, "break_id": br.get("id"), "fraction": float(fraction.get("val")), "gap": [low, high], "crossing_values": crossed, "native_break_shape_count": len(shapes), "physical_tag_count": len(names), "axis_interval_versions": [{"reqver": x.get("reqver"), "endver": x.get("endver"), "length": x.get("length")} for x in intervals]}


def make_job(template: Path, plan: dict, job: Path):
    matrix = profile(plan)["matrix"]
    cell = lambda value: None if value is None else ({"number": value} if isinstance(value, (int, float)) else {"string": value})
    name = plan["targets"][0]["name"]
    job.write_text(json.dumps([{"template": str(template), "data": [{"name": name, "table": [[cell(v) for v in row] for row in matrix]}]}]), encoding="utf-8")


@serialized_office
def run(args):
    source = args.input.resolve()
    output, report = args.output.resolve(), args.report.resolve()
    donor = args.break_donor.resolve() if args.break_donor else None
    input_paths = [source, args.plan.resolve(), args.pre_plan.resolve() if args.pre_plan else None, donor]
    input_paths = [p for p in input_paths if p is not None]
    need(len(input_paths) == len(set(input_paths)), "input, plan, pre-plan and donor must be distinct files")
    need(output not in input_paths and report not in input_paths and output != report, "output and report must be exclusive new files")
    need(sha(source) == args.expected_source_sha256.upper(), "source hash mismatch before execution")
    plan = json.loads(args.plan.read_text(encoding="utf-8-sig"))
    profile(plan)
    pre_plan = json.loads(args.pre_plan.read_text(encoding="utf-8-sig")) if args.pre_plan else plan
    profile(pre_plan)
    need(not output.exists() and not report.exists(), "output and report paths must be new")
    stage = output.parent / (output.stem + "_axis_break_work")
    need(not stage.exists(), "stage path already exists")
    stage.mkdir()
    if args.mode == "insert":
        need(args.expected_donor_sha256, "insert mode requires donor SHA-256")
        prepared = stage / "prepared.pptx"
        prep_report = stage / "prepare-report.json"
        prepare.build(source, donor, args.expected_source_sha256, args.expected_donor_sha256, prepared, prep_report)
        template = prepared
        pre = verify(prepared, pre_plan, expected_shape_count=2)
        need(sha(source) == args.expected_source_sha256.upper() and sha(donor) == args.expected_donor_sha256.upper(), "source or donor changed during preparation")
    else:
        template = source
        pre = verify(source, pre_plan, expected_shape_count=1)
    if not args.execute:
        result = {"status": "PREPARED_NATIVE_EXECUTION_REQUIRED", "mode": args.mode, "source_sha256": sha(source), "pre": pre, "stage": str(stage), "source_unchanged": sha(source) == args.expected_source_sha256.upper()}
        exclusive_write(report, json.dumps(result, indent=2).encode("utf-8"))
        return result
    job, generated = stage / "update.ppttc", stage / "generated.pptx"
    make_job(template, plan, job)
    with (stage / "ppttc.stdout.txt").open("w", encoding="utf-8") as stdout, (stage / "ppttc.stderr.txt").open("w", encoding="utf-8") as stderr:
        code = run_locked_subprocess([str(args.ppttc), str(job), "-o", str(generated)], operation="bounded-native-axis-break-json", timeout_seconds=240, stdout=stdout, stderr=stderr, creationflags=subprocess.CREATE_NO_WINDOW).returncode
    need(code == 0 and generated.exists(), "official JSON generation failed")
    generated_state = verify(generated, plan, expected_shape_count=1)
    native, native_report, render = stage / "native-reopened.pptx", stage / "native-report.json", stage / "native.png"
    command = [powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(HERE / "thinkcell_no_click" / "implementation" / "native_verify_scoped.ps1"), "-InputFile", str(generated), "-OutputFile", str(native), "-ReportFile", str(native_report), "-RenderFile", str(render)]
    with (stage / "native.stdout.txt").open("w", encoding="utf-8") as stdout, (stage / "native.stderr.txt").open("w", encoding="utf-8") as stderr:
        code = run_locked_subprocess(command, operation="bounded-native-axis-break-reopen", timeout_seconds=240, stdout=stdout, stderr=stderr, env=powershell_env(), creationflags=subprocess.CREATE_NO_WINDOW).returncode
    need(code == 0 and native.exists() and native_report.exists() and render.exists(), "native save/reopen or render failed")
    native_scope = json.loads(native_report.read_text(encoding="utf-8-sig"))
    need(native_scope.get("native_reopen_pass") and native_scope.get("source_unchanged") and native_scope.get("other_presentations_unchanged"), "native scope verification failed")
    reopened = verify(native, plan, expected_shape_count=1)
    need(sha(source) == args.expected_source_sha256.upper(), "source changed during execution")
    if donor is not None:
        need(sha(donor) == args.expected_donor_sha256.upper(), "donor changed during execution")
    exclusive_write(output, native.read_bytes())
    result = {"status": "NATIVE_BREAK_INSERTION_GATES_PASS", "mode": args.mode, "source_sha256": sha(source), "donor_sha256": sha(donor) if donor else None, "output_sha256": sha(output), "pre": pre, "generated": generated_state, "native_reopened": reopened, "native_scope": native_scope, "preview": str(render), "source_unchanged": True, "donor_unchanged": donor is None or sha(donor) == args.expected_donor_sha256.upper()}
    exclusive_write(report, json.dumps(result, indent=2).encode("utf-8"))
    return result


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--mode", choices=("insert", "repeat"), required=True)
    p.add_argument("--input", type=Path, required=True)
    p.add_argument("--expected-source-sha256", required=True)
    p.add_argument("--plan", type=Path, required=True, help="Data plan for this update")
    p.add_argument("--pre-plan", type=Path, help="Current-data plan; required when repeat data differs")
    p.add_argument("--break-donor", type=Path)
    p.add_argument("--expected-donor-sha256")
    p.add_argument("--ppttc", type=Path)
    p.add_argument("--execute", action="store_true")
    p.add_argument("--output", type=Path, required=True)
    p.add_argument("--report", type=Path, required=True)
    a = p.parse_args()
    need(a.mode == "repeat" or (a.break_donor and a.expected_donor_sha256), "insert mode requires --break-donor and --expected-donor-sha256")
    need(not a.execute or a.ppttc, "--execute requires --ppttc")
    print(json.dumps(run(a), indent=2))
