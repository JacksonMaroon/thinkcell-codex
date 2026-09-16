"""Parameterised offline rebind for one existing think-cell Excel link.

This is a bounded portable adapter. It discovers one linked OLE carrier by its model
fields, optionally narrows the selection with caller-supplied identity guards,
and replaces only the UTF-16 workbook path in the binary moniker. Arbitrary
length changes remain deliberately unsupported: the replacement path must have
the same UTF-16 byte length as the discovered path.

No Office, COM, UI, or production file is touched by this module.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import io
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import olefile


FIELDS = (
    "m_advisesink", "m_bExternalStorage", "m_bstrRangeName",
    "m_bNeedsUpdateFromSheetOnMakeTC", "CAdviseSink", "m_guidLink",
    "m_lnkid", "m_bAutoUpdate", "m_vecbMoniker",
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def _xml(raw: bytes) -> bytes:
    ole = olefile.OleFileIO(io.BytesIO(raw))
    try:
        return ole.openstream(["think-cellXML"]).read()
    finally:
        ole.close()


def _moniker(value: str | None) -> tuple[str, bytes]:
    if not value:
        return "", b""
    encoded = value.strip() + "=" * ((-len(value.strip())) % 4)
    payload = base64.b64decode(encoded)
    runs = [m.group(0).decode("utf-16-le", "ignore") for m in re.finditer(rb"(?:[\x20-\x7e]\x00){3,}", payload)]
    for _, _, path in _path_runs(payload):
        if path not in runs:
            runs.insert(0, path)
    return ";".join(runs), payload


def _path_runs(payload: bytes) -> list[tuple[int, int, str]]:
    """Find null-terminated Windows workbook paths, including BMP chars."""
    found = []
    for prefix in re.finditer(rb"[A-Za-z]\x00:\x00\\\x00", payload):
        start = prefix.start()
        end = start
        while end + 1 < len(payload):
            unit = payload[end:end + 2]
            if unit == b"\x00\x00":
                break
            codepoint = int.from_bytes(unit, "little")
            if codepoint < 0x20 or 0xD800 <= codepoint <= 0xDFFF:
                break
            end += 2
            text = payload[start:end].decode("utf-16-le", "strict")
            excel_extension = re.search(r"\.(?:xlsx|xlsm|xlsb)$", text, re.IGNORECASE)
            if not excel_extension and text.lower().endswith(".xls"):
                next_unit = payload[end:end + 2]
                next_codepoint = int.from_bytes(next_unit, "little") if len(next_unit) == 2 else 0
                excel_extension = end + 1 >= len(payload) or not (0x30 <= next_codepoint <= 0x7A)
            if excel_extension:
                found.append((start, end, text))
                break
    return found


def _norm_ref(value: str) -> str:
    return re.sub(r"\s+", "", value or "").replace("$", "").upper()


def named_range(path: Path, name: str, expected_sheet: str) -> dict[str, str]:
    """Read the XLSX defined name that anchors the think-cell link."""
    with zipfile.ZipFile(path) as package:
        workbook = ET.fromstring(package.read("xl/workbook.xml"))
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    sheets = [node.get("name", "") for node in workbook.findall("x:sheets/x:sheet", ns)]
    names = workbook.find("x:definedNames", ns)
    matches = [] if names is None else [node for node in names.findall("x:definedName", ns) if node.get("name") == name]
    if len(matches) != 1:
        raise RuntimeError(f"target workbook must contain exactly one defined name {name!r}; found {len(matches)}")
    node = matches[0]
    local_id = node.get("localSheetId")
    if local_id is None or not local_id.isdigit() or int(local_id) >= len(sheets):
        raise RuntimeError(f"defined name {name!r} has no valid local sheet identity")
    sheet = sheets[int(local_id)]
    if sheet != expected_sheet:
        raise RuntimeError(f"defined name {name!r} is bound to sheet {sheet!r}, expected {expected_sheet!r}")
    reference = (node.text or "").strip()
    if not reference or "!" not in reference:
        raise RuntimeError(f"defined name {name!r} has no worksheet range reference")
    ref_sheet, ref_range = reference.rsplit("!", 1)
    ref_sheet = ref_sheet.strip("'")
    if ref_sheet != expected_sheet:
        raise RuntimeError(f"defined name reference sheet {ref_sheet!r} differs from {expected_sheet!r}")
    return {"name": name, "sheet": sheet, "reference": reference, "range": _norm_ref(ref_range), "hidden": node.get("hidden", "0")}


def _fields(xml: bytes) -> tuple[dict[str, str | list[str]], str, bytes]:
    root = ET.fromstring(xml)
    nodes = {tag: list(root.iter(tag)) for tag in FIELDS}
    if any(len(items) != 1 for items in nodes.values()):
        raise ValueError("linked carrier does not have one complete identity field set")
    fields: dict[str, str | list[str]] = {
        "advisesink_idref": nodes["m_advisesink"][0].get("idref", ""),
        "external_storage": nodes["m_bExternalStorage"][0].get("val", ""),
        "range_name": nodes["m_bstrRangeName"][0].get("val", ""),
        "needs_update_from_sheet_on_make_tc": nodes["m_bNeedsUpdateFromSheetOnMakeTC"][0].get("val", ""),
        "advisesink_id": nodes["CAdviseSink"][0].get("id", ""),
        "guid_link": nodes["m_guidLink"][0].get("val", ""),
        "link_id": (nodes["m_lnkid"][0].text or "").strip(),
        "auto_update": nodes["m_bAutoUpdate"][0].get("val", ""),
    }
    moniker_text, payload = _moniker(nodes["m_vecbMoniker"][0].text)
    fields["moniker_text"] = moniker_text
    return fields, moniker_text, payload


def discover(path: Path, guid: str | None = None, link_id: str | None = None, range_name: str | None = None):
    candidates = []
    with zipfile.ZipFile(path) as package:
        for part in sorted(package.namelist()):
            if not (part.startswith("ppt/embeddings/") and part.endswith(".bin")):
                continue
            raw = package.read(part)
            try:
                xml = _xml(raw)
            except Exception:
                continue
            try:
                fields, moniker_text, payload = _fields(xml)
            except Exception as exc:
                # A think-cell XML part that advertises external-link fields but
                # is malformed must stop discovery. Silently skipping it could
                # turn a malformed carrier into a false unique match.
                external_marker = re.search(rb"<m_bExternalStorage\b[^>]*\bval=[\"']1[\"']", xml)
                identity_markers = sum(marker in xml for marker in (b"<CAdviseSink", b"<m_guidLink", b"<m_lnkid", b"<m_vecbMoniker"))
                # A data-table's storage flag alone is common for ordinary
                # think-cell carriers. Treat it as an external-link candidate
                # only when it also advertises at least one link identity
                # field; a complete identity set is likewise unambiguous.
                if (external_marker and identity_markers > 0) or identity_markers == 4:
                    raise RuntimeError(f"malformed linked carrier at {part}: {exc}") from exc
                continue
            if fields["external_storage"] != "1" or not fields["guid_link"] or not fields["link_id"] or not moniker_text:
                continue
            if guid is not None and fields["guid_link"] != guid:
                continue
            if link_id is not None and fields["link_id"] != link_id:
                continue
            if range_name is not None and fields["range_name"] != range_name:
                continue
            candidates.append({"part": part, "raw": raw, "xml": xml, "fields": fields, "moniker_text": moniker_text, "payload": payload})
    if len(candidates) != 1:
        raise RuntimeError(f"expected exactly one existing linked carrier after optional guards, found {len(candidates)}")
    return candidates[0]


def rebind_xml(xml: bytes, payload: bytes, new_path: str) -> tuple[bytes, str]:
    path_runs = _path_runs(payload)
    paths = [(i, value) for i, (_, _, value) in enumerate(path_runs)]
    if len(paths) != 1:
        raise RuntimeError(f"expected exactly one absolute Windows path in moniker, found {paths!r}")
    index, old_path = paths[0]
    if len(new_path) != len(old_path):
        raise RuntimeError(f"equal-moniker-length guard: discovered path has {len(old_path)} characters, target has {len(new_path)}")
    if new_path == old_path:
        raise RuntimeError("target workbook path is already bound; refusing a no-op output")
    replacement = new_path.encode("utf-16-le")
    start, old_end, _ = path_runs[index]
    if len(replacement) != old_end - start:
        raise RuntimeError(f"equal-moniker-UTF16-byte-length guard: discovered path has {old_end - start} bytes, target has {len(replacement)}")
    payload2 = payload[:start] + replacement + payload[old_end:]
    marker = re.search(rb"<m_vecbMoniker>([^<]*)</m_vecbMoniker>", xml)
    if not marker:
        raise RuntimeError("m_vecbMoniker serialization not found")
    encoded = base64.b64encode(payload2)
    if len(encoded) != len(marker.group(1)):
        raise RuntimeError("moniker base64 length changed")
    xml2 = xml[:marker.start(1)] + encoded + xml[marker.end(1):]
    if len(xml2) != len(xml):
        raise RuntimeError("think-cell XML length changed")
    return xml2, old_path


def patch_part(raw: bytes, xml: bytes, payload: bytes, new_path: str) -> tuple[bytes, str]:
    xml2, old_path = rebind_xml(xml, payload, new_path)
    buffer = io.BytesIO(raw)
    writer = olefile.OleFileIO(buffer, write_mode=True)
    try:
        writer.write_stream(["think-cellXML"], xml2)
    finally:
        writer.close()
    return buffer.getvalue(), old_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Rebind one existing linked think-cell Excel moniker offline.")
    parser.add_argument("--input-presentation", type=Path, required=True)
    parser.add_argument("--output-presentation", type=Path, required=True)
    parser.add_argument("--target-workbook", type=Path, required=True)
    parser.add_argument("--source-workbook", type=Path, required=True, help="Original workbook used to prove named-range compatibility.")
    parser.add_argument("--expected-range", help="Optional additional A1 range assertion.")
    parser.add_argument("--input-sha256", required=True, help="Required SHA-256 guard for the input presentation.")
    parser.add_argument("--source-workbook-sha256", required=True, help="Required SHA-256 guard for the source workbook.")
    parser.add_argument("--target-workbook-sha256", help="Optional SHA-256 guard for the target workbook.")
    parser.add_argument("--guid", help="Optional exact GUID selector.")
    parser.add_argument("--link-id", help="Optional exact link-ID selector.")
    parser.add_argument("--range-name", help="Optional exact think-cell range-name selector.")
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    input_path, output_path, workbook_path, report_path = tuple(p.resolve() for p in (args.input_presentation, args.output_presentation, args.target_workbook, args.report))
    if not input_path.exists() or not workbook_path.exists():
        raise SystemExit("input presentation and target workbook must exist")
    if input_path == workbook_path:
        raise SystemExit("input presentation and target workbook must be distinct")
    if output_path in {input_path, workbook_path} or report_path in {input_path, output_path, workbook_path}:
        raise SystemExit("output and report must be distinct from input/workbook")
    if output_path.exists() or report_path.exists():
        raise SystemExit("fresh output and report paths are required")
    source_workbook = args.source_workbook.resolve() if args.source_workbook else None
    if source_workbook is not None:
        if not source_workbook.exists():
            raise SystemExit(f"source workbook does not exist: {source_workbook}")
        if source_workbook in {input_path, output_path, workbook_path, report_path}:
            raise SystemExit("source workbook must be distinct from presentation, target, and report paths")
    input_hash = sha256(input_path)
    workbook_hash = sha256(workbook_path)
    if input_hash != args.input_sha256.upper():
        raise SystemExit(f"input SHA-256 mismatch: expected {args.input_sha256.upper()}, got {input_hash}")
    if args.target_workbook_sha256 and workbook_hash != args.target_workbook_sha256.upper():
        raise SystemExit(f"target workbook SHA-256 mismatch: expected {args.target_workbook_sha256.upper()}, got {workbook_hash}")
    source_workbook_hash = sha256(source_workbook)
    if source_workbook_hash != args.source_workbook_sha256.upper():
        raise SystemExit(f"source workbook SHA-256 mismatch: expected {args.source_workbook_sha256.upper()}, got {source_workbook_hash}")
    selected = discover(input_path, args.guid, args.link_id, args.range_name)
    moniker_tail = selected["moniker_text"].split(";")[-1]
    if moniker_tail.startswith("W"):
        moniker_tail = moniker_tail[1:]
    if "!" not in moniker_tail:
        raise RuntimeError("linked moniker has no worksheet/name identity")
    moniker_sheet, moniker_name = moniker_tail.split("!", 1)
    target_named_range = named_range(workbook_path, moniker_name, moniker_sheet)
    old_paths = [item for item in selected["moniker_text"].split(";") if re.match(r"^[A-Za-z]:\\", item)]
    if len(old_paths) != 1:
        raise RuntimeError(f"linked moniker must contain exactly one source workbook path; found {old_paths!r}")
    source_named_range = named_range(source_workbook, moniker_name, moniker_sheet)
    compatibility_keys = ("name", "sheet", "range")
    mismatches = {key: (source_named_range[key], target_named_range[key]) for key in compatibility_keys if source_named_range[key] != target_named_range[key]}
    if mismatches:
        raise RuntimeError(f"target named-range identity differs from source: {mismatches}")
    if args.expected_range and target_named_range["range"] != _norm_ref(args.expected_range):
        raise RuntimeError(f"target named range range {target_named_range['range']!r} differs from expected {_norm_ref(args.expected_range)!r}")
    patched, old_path = patch_part(selected["raw"], selected["xml"], selected["payload"], str(workbook_path))
    parts: dict[str, bytes]; infos: dict[str, zipfile.ZipInfo]
    with zipfile.ZipFile(input_path) as package:
        parts = {name: package.read(name) for name in package.namelist()}
        infos = {name: package.getinfo(name) for name in package.namelist()}
    parts[selected["part"]] = patched
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        for name, data in parts.items():
            package.writestr(infos[name], data)
    # Compare against the sealed input bytes rather than the selected payload.
    with zipfile.ZipFile(input_path) as before, zipfile.ZipFile(output_path) as after:
        names = sorted(set(before.namelist()) | set(after.namelist()))
        changed = [name for name in names if before.read(name) != after.read(name)]
    if changed != [selected["part"]]:
        raise RuntimeError(f"portable rebind changed unexpected package parts: {changed}")
    rebound = discover(output_path, selected["fields"]["guid_link"], selected["fields"]["link_id"], selected["fields"]["range_name"])
    preservation = {"input_presentation_unchanged": sha256(input_path) == input_hash, "target_workbook_unchanged": sha256(workbook_path) == workbook_hash, "source_workbook_unchanged": sha256(source_workbook) == source_workbook_hash}
    if not all(preservation.values()):
        raise RuntimeError(f"source or target changed during offline rebind: {preservation}")
    report = {
        "schema": "thinkcell-portable-rebind-v1",
        "status": "PREPARED_OFFLINE_PORTABLE_REBIND",
        "input_presentation": str(input_path),
        "output_presentation": str(output_path),
        "target_workbook": str(workbook_path),
        "input_sha256": input_hash,
        "output_sha256": sha256(output_path),
        "target_workbook_sha256": workbook_hash,
        "source_workbook": str(source_workbook),
        "source_workbook_sha256": source_workbook_hash,
        "selected_part": selected["part"],
        "changed_parts": changed,
        "old_moniker_workbook_path": old_path,
        "new_moniker_workbook_path": str(workbook_path),
        "identity_before": selected["fields"],
        "identity_after": rebound["fields"],
        "moniker_after": rebound["moniker_text"],
        "guards": {"guid": args.guid, "link_id": args.link_id, "range_name": args.range_name, "equal_moniker_length": True, "named_range_compatibility": {"target": target_named_range, "source": source_named_range, "expected_range": _norm_ref(args.expected_range) if args.expected_range else None}},
        "preservation_after_readback": preservation,
        "native_certification": "NOT_PERFORMED_OFFLINE_ADAPTER_ONLY",
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
