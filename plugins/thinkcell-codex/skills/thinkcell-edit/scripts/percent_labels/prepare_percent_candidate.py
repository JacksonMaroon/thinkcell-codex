"""Prepare a coherent semantic percent-label candidate on a copy.

The candidate keeps the scalar's native relative source and physical field,
changes only the requested precision, and removes the absolute field from the
physical shape. Wrapper mode "bare" is intentionally experimental: native
regeneration must still prove that the model owns the resulting bare label.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import io
import json
import posixpath
import subprocess
import sys
import tempfile
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from lxml import etree as E

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from portable_scripts import resolve_plugin_scripts  # noqa: E402
SCRIPTS = resolve_plugin_scripts(HERE)
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(SCRIPTS / "thinkcell_no_click" / "implementation"))
sys.path.insert(0, str(HERE))
from chart_geometry import inventory, streams, xml  # noqa: E402
from runtime import powershell, powershell_env  # noqa: E402
from discover_percent_semantics import discover, select  # noqa: E402

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def need(value, message):
    if not value:
        raise ValueError(message)


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def format_percent(numerator, denominator, digits):
    value = (Decimal(str(numerator)) * Decimal("100") / Decimal(str(denominator))).quantize(
        Decimal("1." + "0" * digits), rounding=ROUND_HALF_UP)
    return f"{value}%"


def format_key(display: str) -> str:
    """Encode the visible label as the native single-character format key."""
    return "".join("'" + char + "'" for char in display)


def find_shape(package, tag):
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
                target = posixpath.normpath(posixpath.join(posixpath.dirname(slide_name), rel.get("Target")))
                if target.startswith("../"):
                    target = posixpath.normpath(posixpath.join(posixpath.dirname(slide_name), target))
                if target not in package.namelist():
                    continue
                if tag in package.read(target).decode("utf-8", errors="strict"):
                    return slide_name, slide, shape
    raise ValueError("semantic physical shape tag unresolved")


def prepare(source: Path, output: Path, report: Path, *, category: str, series: str,
            digits: int, wrapper: str, expected_sha: str, chart_name: str | None = None) -> dict:
    need(0 <= digits <= 3, "precision must be 0 through 3 decimal places")
    need(wrapper in {"parentheses", "bare"}, "wrapper must be parentheses or bare")
    need(source.resolve() not in {output.resolve(), report.resolve()}, "paths must differ")
    need(not output.exists() and not report.exists(), "outputs must be new")
    raw = source.read_bytes()
    need(sha(raw) == expected_sha.upper(), "source SHA-256 mismatch")
    selected = select(discover(source), category=category, series=series, chart_name=chart_name)
    physical = selected["physical_shapes"]
    need(len(physical) == 1 and len(physical[0]["fields"]) == 2, "selected label is not a dual native label")
    expected = format_percent(selected["numerator"], selected["denominator"], digits)
    expected_format_key = format_key(expected)
    with zipfile.ZipFile(io.BytesIO(raw)) as zin:
        entries = {info.filename: zin.read(info.filename) for info in zin.infolist()}
        infos = {info.filename: info for info in zin.infolist()}
        slide_name, slide, shape = find_shape(zin, selected["shape_tag"])
    para = shape.find("p:txBody", NS).find("a:p", NS)
    fields = para.findall("a:fld", NS)
    need(len(fields) == 2, "selected physical shape has unexpected field count")
    need(fields[0].findtext("a:t", namespaces=NS) == physical[0]["fields"][0]["text"], "absolute field identity mismatch")
    need(fields[1].findtext("a:t", namespaces=NS) == physical[0]["fields"][1]["text"], "relative field identity mismatch")
    for node in list(para):
        if node is fields[0] or node.tag == "{%s}br" % NS["a"]:
            para.remove(node)
        elif wrapper == "bare" and node.tag == "{%s}r" % NS["a"] and node.findtext("a:t", namespaces=NS) in {"(", ")"}:
            para.remove(node)
    remain = para.findall("a:fld", NS)
    need(len(remain) == 1 and remain[0] is fields[1], "relative field was replaced")
    remain[0].set("type", "datetime" + expected_format_key)
    remain[0].find("a:t", NS).text = expected
    entries[slide_name] = E.tostring(slide, xml_declaration=True, encoding="UTF-8", standalone=True)

    _, charts, _ = inventory(raw)
    chart = next(c for c in charts if c["doc"]["part"] == selected["chart_part"])
    model_before = chart["doc"]["streams"][("think-cellXML",)]
    root = xml(model_before)
    ids = {n.get("id"): n for n in root if n.get("id")}
    relative = ids[selected["relative_source_id"]]
    text_ref = relative.find("m_ctextvar/elem")
    need(text_ref is not None and text_ref.get("idref") == selected["relative_text_variable"], "relative model binding changed")
    textvar = ids[text_ref.get("idref")]
    format_node = textvar.find("m_bstrFormat")
    need(format_node is not None and format_node.text, "relative format key is missing")
    previous_format_key = format_node.text
    format_node.text = expected_format_key
    precision = textvar.findall("m_prec17834")
    need(len(precision) == 1, "relative precision is missing or ambiguous")
    decimal_node = precision[0].find("m_nDecimalDigits17909")
    need(decimal_node is not None, "relative decimal precision is missing")
    old_digits = decimal_node.get("val")
    decimal_node.set("val", str(digits))
    model_after = E.tostring(root, encoding="utf-8")
    with tempfile.TemporaryDirectory(prefix="percent_candidate_", dir=output.parent) as tmp:
        tmp = Path(tmp)
        carrier = tmp / "carrier.bin"
        payload = tmp / "model.xml"
        carrier.write_bytes(chart["doc"]["ole"])
        payload.write_bytes(model_after)
        command = [powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File",
                   str(SCRIPTS / "thinkcell_no_click" / "implementation" / "replace_ole_stream.ps1"),
                   "-StoragePath", str(carrier), "-StreamBytesPath", str(payload)]
        proc = subprocess.run(command, capture_output=True, text=True, env=powershell_env(), timeout=60)
        need(proc.returncode == 0, "OLE model patch failed: " + proc.stderr[-400:])
        changed_carrier = carrier.read_bytes()
    entries[selected["chart_part"]] = changed_carrier
    with zipfile.ZipFile(output, "x") as zout:
        for info in infos.values():
            zout.writestr(copy.copy(info), entries[info.filename])
    with zipfile.ZipFile(io.BytesIO(raw)) as before_zip, zipfile.ZipFile(output) as after_zip:
        changed = [name for name in after_zip.namelist() if before_zip.read(name) != after_zip.read(name)]
    need(set(changed) == {selected["chart_part"], slide_name}, "unexpected package changes: " + str(changed))
    result = {
        "status": "PREPARED_PERCENT_SEMANTIC_CANDIDATE_NATIVE_PROOF_REQUIRED",
        "source_sha256": expected_sha.upper(), "output_sha256": sha(output.read_bytes()),
        "selector": {"category": category, "series": series, "chart_name": selected["chart_name"],
                     "chart_part": selected["chart_part"]},
        "selected": selected, "requested_precision": digits, "previous_precision": old_digits,
        "wrapper_mode": wrapper, "expected_display": expected,
        "previous_format_key": previous_format_key, "format_key": expected_format_key,
        "model_relative_source_id": selected["relative_source_id"],
        "physical_field_id_retained": physical[0]["fields"][1]["id"],
        "changed_entries": changed, "native_gates": ["official regeneration", "native save/reopen",
        "changed numerator", "changed denominator", "visual wrapper review"],
    }
    report.write_text(json.dumps(result, indent=2), encoding="utf-8")
    return result


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--input", type=Path, required=True)
    p.add_argument("--output", type=Path, required=True)
    p.add_argument("--report", type=Path, required=True)
    p.add_argument("--expected-sha256", required=True)
    p.add_argument("--category", required=True)
    p.add_argument("--series", required=True)
    p.add_argument("--chart-name")
    p.add_argument("--digits", type=int, required=True)
    p.add_argument("--wrapper", choices=["parentheses", "bare"], default="parentheses")
    a = p.parse_args()
    print(json.dumps(prepare(a.input, a.output, a.report, category=a.category, series=a.series,
                             digits=a.digits, wrapper=a.wrapper, expected_sha=a.expected_sha256,
                             chart_name=a.chart_name), indent=2))
