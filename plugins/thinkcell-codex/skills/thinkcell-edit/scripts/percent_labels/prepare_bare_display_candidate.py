"""Prepare a native bare-display percent candidate on a guarded copy.

This adapter first uses the production semantic percent candidate path to
retain the selected relative text field and its model binding, then replaces
only the physical wrapper runs with U+200B.  It never writes a static label or
changes chart geometry.  Native regeneration and reopen remain required.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import io
import json
import os
import posixpath
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree as E

HERE = Path(__file__).resolve().parent


def _load_production_modules():
    """Resolve the packaged script rail, honoring THINKCELL_PLUGIN_SCRIPTS."""
    root = HERE
    portable = root / "portable_scripts.py"
    if not portable.is_file():
        raise FileNotFoundError("installed percent_labels scripts rail is missing portable_scripts.py")
    sys.path.insert(0, str(root))
    from portable_scripts import resolve_plugin_scripts  # type: ignore
    scripts = resolve_plugin_scripts(root)
    percent = scripts / "percent_labels"
    if not percent.is_dir():
        raise FileNotFoundError("resolved scripts rail has no percent_labels directory: " + str(percent))
    sys.path.insert(0, str(percent))
    from prepare_percent_candidate import prepare  # type: ignore
    from discover_percent_semantics import discover, select  # type: ignore
    return scripts, prepare, discover, select


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
ZWSP = "\u200b"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def _find_shape(package: zipfile.ZipFile, tag: str):
    """Find one physical shape carrying the selected model shape tag."""
    for slide_name in package.namelist():
        if not (slide_name.startswith("ppt/slides/slide") and slide_name.endswith(".xml")):
            continue
        rel_name = f"{slide_name.rsplit('/', 1)[0]}/_rels/{slide_name.rsplit('/', 1)[1]}.rels"
        if rel_name not in package.namelist():
            continue
        slide = E.fromstring(package.read(slide_name))
        rels = E.fromstring(package.read(rel_name))
        by_id = {rel.get("Id"): rel for rel in rels}
        for shape in slide.xpath(".//p:sp", namespaces=NS):
            for rid in shape.xpath(".//p:tags/@r:id", namespaces=NS):
                rel = by_id.get(rid)
                if rel is None:
                    continue
                target = posixpath.normpath(
                    posixpath.join(posixpath.dirname(slide_name), rel.get("Target", "")))
                if target in package.namelist() and tag in package.read(target).decode("utf-8"):
                    return slide_name, slide, shape
    raise ValueError("semantic physical shape tag unresolved")


def _fresh(path: Path, *others: Path) -> None:
    if path.exists() or any(path.resolve() == other.resolve() for other in others):
        raise FileExistsError("output/report must be new and distinct from input")


def prepare(source: Path, output: Path, report: Path, *, expected_sha: str,
            category: str, series: str, chart_name: str | None, digits: int) -> dict:
    scripts, production_prepare, discover, select = _load_production_modules()
    _fresh(output, source, report)
    _fresh(report, source, output)
    actual_sha = sha256(source)
    if actual_sha != expected_sha.upper():
        raise ValueError(f"source SHA-256 mismatch: {actual_sha} != {expected_sha.upper()}")

    selected = select(discover(source), category=category, series=series, chart_name=chart_name)
    if len(selected.get("physical_shapes", [])) != 1:
        raise ValueError("selected label must have exactly one physical shape")
    expected_display = selected.get("relative_suffix", "")
    if expected_display != "%":
        raise ValueError("selected field is not a relative percent field")

    with tempfile.TemporaryDirectory(prefix="bare_display_", dir=str(output.parent)) as temp:
        temp_dir = Path(temp)
        parenthesized = temp_dir / "parenthesized.pptx"
        parent_report_path = temp_dir / "parenthesized.json"
        parent = production_prepare(
            source, parenthesized, parent_report_path,
            category=category, series=series, digits=digits,
            wrapper="parentheses", expected_sha=expected_sha,
            chart_name=chart_name)
        parent_shape_tag = parent["selected"]["shape_tag"]
        with zipfile.ZipFile(parenthesized) as zin:
            entries = {info.filename: zin.read(info.filename) for info in zin.infolist()}
            infos = {info.filename: info for info in zin.infolist()}
            slide_name, slide, shape = _find_shape(zin, parent_shape_tag)
        texts = shape.xpath(".//a:t", namespaces=NS)
        left = [node for node in texts if node.text == "("]
        right = [node for node in texts if node.text == ")"]
        if len(left) != 1 or len(right) != 1:
            raise ValueError("production parenthesized candidate did not expose one wrapper pair")
        left[0].text = ZWSP
        right[0].text = ZWSP
        entries[slide_name] = E.tostring(
            slide, encoding="UTF-8", xml_declaration=True, standalone=True)
        with zipfile.ZipFile(output, "x") as zout:
            for info in infos.values():
                zout.writestr(copy.copy(info), entries[info.filename])

    with zipfile.ZipFile(source) as before, zipfile.ZipFile(output) as after:
        changed = [name for name in after.namelist()
                   if before.read(name) != after.read(name)]
    if set(changed) != {selected["chart_part"], slide_name}:
        raise ValueError(f"unexpected package changes: {changed}")
    final_rows = discover(output)
    final = select(final_rows, category=category, series=series, chart_name=chart_name)
    shapes = final.get("physical_shapes", [])
    literals = [str(value or "") for shape_info in shapes for value in shape_info.get("literal_texts", [])]
    if len(shapes) != 1 or len(shapes[0].get("fields", [])) != 1:
        raise ValueError("bare-display candidate lost its single relative physical field")
    if literals.count(ZWSP) < 2:
        raise ValueError("bare-display candidate lacks raw U+200B wrapper runs")
    result = {
        "status": "PREPARED_BARE_DISPLAY_CANDIDATE_NATIVE_PROOF_REQUIRED",
        "source_sha256": actual_sha,
        "output_sha256": sha256(output),
        "selector": {"category": category, "series": series,
                      "chart_name": selected["chart_name"],
                      "chart_part": selected["chart_part"],
                      "shape_tag": selected["shape_tag"]},
        "requested_precision": digits,
        "expected_display_from_source": parent["expected_display"],
        "wrapper_mode": "bare-display-zero-width",
        "wrapper_codepoint": "U+200B ZERO WIDTH SPACE",
        "model_route": "production semantic relative field preparation",
        "physical_field_id_retained": shapes[0]["fields"][0]["id"],
        "changed_entries": changed,
        "runtime_scripts": str(scripts),
        "native_gates": [
            "official regeneration",
            "native save/reopen",
            "exact changed datasheet numerator and denominator",
            "selected relative field and physical label survive",
            "raw U+200B prefix/suffix and visible bare label",
            "changed-data behavior and sibling preservation",
        ],
    }
    report.write_text(json.dumps(result, indent=2), encoding="utf-8")
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--expected-sha256", required=True)
    parser.add_argument("--category", required=True)
    parser.add_argument("--series", required=True)
    parser.add_argument("--chart-name")
    parser.add_argument("--digits", type=int, required=True)
    args = parser.parse_args()
    result = prepare(args.input, args.output, args.report,
                     expected_sha=args.expected_sha256, category=args.category,
                     series=args.series, chart_name=args.chart_name,
                     digits=args.digits)
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
