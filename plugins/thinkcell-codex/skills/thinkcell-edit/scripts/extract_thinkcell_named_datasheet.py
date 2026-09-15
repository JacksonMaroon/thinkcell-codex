from __future__ import annotations

import argparse
import hashlib
import io
import json
from pathlib import Path
import struct
import sys
import zipfile

from lxml import etree

VENDOR = Path(__file__).resolve().parent / "vendor"
if str(VENDOR) not in sys.path:
    sys.path.insert(0, str(VENDOR))

try:
    import olefile  # type: ignore
except ImportError as exc:  # pragma: no cover - environment failure
    raise SystemExit(f"Missing olefile dependency; install scripts/requirements.txt: {exc}")

from audit_thinkcell_integrity import NS, R, logical_slides, relationship_map


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def node_text(node: etree._Element | None) -> str:
    if node is None:
        return ""
    return (node.text or node.get("val") or "").strip()


FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


def _directory_entry(
    name: str,
    object_type: int,
    child: int = NOSTREAM,
    start_sector: int = ENDOFCHAIN,
    stream_size: int = 0,
) -> bytes:
    """Build one version-3 Compound File directory entry.

    Legacy think-cell elements can retain the Excel BIFF Workbook stream
    directly under their child storage instead of packaging an XLSB ZIP.
    Excel still expects that BIFF stream inside a Compound File Binary
    container, so the extractor reconstructs the narrow one-stream wrapper
    rather than treating the raw stream as an XLSB file.
    """
    encoded = (name + "\0").encode("utf-16le")
    if len(encoded) > 64:
        raise ValueError(f"Compound-file directory name is too long: {name}")
    entry = bytearray(128)
    entry[: len(encoded)] = encoded
    struct.pack_into("<H", entry, 64, len(encoded))
    entry[66] = object_type
    entry[67] = 1  # black node in the directory red-black tree
    struct.pack_into("<III", entry, 68, NOSTREAM, NOSTREAM, child)
    struct.pack_into("<I", entry, 116, start_sector)
    struct.pack_into("<Q", entry, 120, stream_size)
    return bytes(entry)


def wrap_biff_workbook_stream(workbook: bytes) -> bytes:
    """Wrap a legacy BIFF Workbook stream in a minimal valid CFB .xls file."""
    if len(workbook) < 8:
        raise ValueError("The legacy BIFF Workbook stream is too short to contain a BIFF8 BOF record.")
    record_id, record_length, biff_version = struct.unpack_from("<HHH", workbook, 0)
    if record_id != 0x0809 or record_length < 4 or biff_version != 0x0600:
        raise ValueError(
            "The legacy Workbook stream is not a supported BIFF8 workbook (expected BOF 0x0809, version 0x0600)."
        )
    # A stream below 4096 bytes belongs in the CFB mini-stream. Pad this
    # temporary reader wrapper to the regular-stream threshold instead.
    # The source BIFF stream and presentation remain unchanged.
    workbook = workbook.ljust(4096, b"\0")
    sector_size = 512
    data_sector_count = (len(workbook) + sector_size - 1) // sector_size
    directory_sector = data_sector_count
    fat_sector = directory_sector + 1
    if fat_sector >= 128:
        raise ValueError(
            "The legacy BIFF Workbook stream is too large for the supported one-FAT-sector wrapper."
        )

    header = bytearray(sector_size)
    header[:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HH", header, 24, 0x003E, 0x0003)
    struct.pack_into("<H", header, 28, 0xFFFE)
    struct.pack_into("<HH", header, 30, 9, 6)
    struct.pack_into("<I", header, 40, 0)
    struct.pack_into("<I", header, 44, 1)
    struct.pack_into("<I", header, 48, directory_sector)
    struct.pack_into("<I", header, 52, 0)
    struct.pack_into("<I", header, 56, 4096)
    struct.pack_into("<I", header, 60, ENDOFCHAIN)
    struct.pack_into("<I", header, 64, 0)
    struct.pack_into("<I", header, 68, ENDOFCHAIN)
    struct.pack_into("<I", header, 72, 0)
    struct.pack_into("<I", header, 76, fat_sector)
    for offset in range(80, sector_size, 4):
        struct.pack_into("<I", header, offset, FREESECT)

    data = workbook + b"\0" * (data_sector_count * sector_size - len(workbook))
    directory = bytearray(sector_size)
    directory[:128] = _directory_entry("Root Entry", 5, child=1)
    directory[128:256] = _directory_entry(
        "Workbook", 2, start_sector=0, stream_size=len(workbook)
    )

    fat = [FREESECT] * 128
    for sector in range(data_sector_count):
        fat[sector] = sector + 1 if sector + 1 < data_sector_count else ENDOFCHAIN
    fat[directory_sector] = ENDOFCHAIN
    fat[fat_sector] = FATSECT
    fat_bytes = struct.pack("<128I", *fat)
    return bytes(header) + data + bytes(directory) + fat_bytes


def named_storage_candidates(
    model_root: etree._Element, name: str
) -> list[tuple[etree._Element, etree._Element, str]]:
    ids = {node.get("id"): node for node in model_root if node.get("id")}
    candidates = []
    seen_storage_owners: set[tuple[str, str]] = set()
    for node in model_root:
        if node_text(node.find("m_strName")).casefold() != name.casefold():
            continue
        storage_owner = node
        storage = node_text(storage_owner.find("m_bstrRangeName"))
        if not storage:
            # Some families, including pie and scatter/bubble, own the
            # automation name on the chart while the linked data table owns
            # m_bstrRangeName. Resolve the genuine table through m_dtable
            # rather than guessing by child position.
            table_reference = node.find("m_dtable")
            storage_owner = (
                ids.get(table_reference.get("idref"))
                if table_reference is not None and table_reference.get("idref")
                else None
            )
            storage = (
                node_text(storage_owner.find("m_bstrRangeName"))
                if storage_owner is not None
                else ""
            )
        if storage and storage_owner is not None:
            storage_key = (storage_owner.get("id") or "", storage.casefold())
            if storage_key in seen_storage_owners:
                continue
            seen_storage_owners.add(storage_key)
            candidates.append((node, storage_owner, storage))
    return candidates


def extract_named_datasheet(
    presentation: Path, slide_number: int, name: str
) -> tuple[bytes, str, dict[str, object]]:
    with zipfile.ZipFile(presentation) as package:
        slides = logical_slides(package)
        if slide_number < 1 or slide_number > len(slides):
            raise ValueError(
                f"Slide {slide_number} is outside the presentation's {len(slides)} slides."
            )
        slide_part = str(slides[slide_number - 1]["part"])
        slide_root = etree.fromstring(package.read(slide_part))
        rels = relationship_map(package, slide_part)
        matches: list[tuple[bytes, str, dict[str, object]]] = []
        seen_active_document_parts: set[str] = set()

        for ole in slide_root.xpath(".//p:oleObj", namespaces=NS):
            relationship = rels.get(ole.get(R + "id"))
            if not relationship or relationship["resolved"] not in package.namelist():
                continue
            active_document_part = relationship["resolved"]
            if active_document_part in seen_active_document_parts:
                continue
            seen_active_document_parts.add(active_document_part)
            active_document = package.read(active_document_part)
            if (
                "TCLayout.ActiveDocument" not in ole.get("progId", "")
                and b"think-cellXML" not in active_document
            ):
                continue
            with olefile.OleFileIO(io.BytesIO(active_document)) as compound:
                if not compound.exists("think-cellXML"):
                    continue
                model_root = etree.fromstring(compound.openstream("think-cellXML").read())
                candidates = named_storage_candidates(model_root, name)
                for named_owner, storage_owner, storage in candidates:
                    package_stream = [storage, "Package"]
                    legacy_stream = [storage, "Workbook"]
                    if compound.exists(package_stream):
                        workbook = compound.openstream(package_stream).read()
                        storage_kind = "xlsb_package"
                        source_workbook_stream = None
                    elif compound.exists(legacy_stream):
                        source_workbook_stream = compound.openstream(legacy_stream).read()
                        workbook = wrap_biff_workbook_stream(source_workbook_stream)
                        storage_kind = "legacy_biff_cfb"
                    else:
                        raise ValueError(
                            f"Named element '{name}' uses storage '{storage}', but neither its Package nor Workbook stream is present."
                        )
                    matches.append(
                        (
                            workbook,
                            storage_kind,
                            {
                                "presentation": str(presentation),
                                "presentation_sha256": sha256(presentation.read_bytes()),
                                "slide": slide_number,
                                "slide_id": slides[slide_number - 1]["id"],
                                "slide_part": slide_part,
                                "name": name,
                                "name_owner": etree.QName(named_owner).localname,
                                "table_owner": etree.QName(storage_owner).localname,
                                "active_document_part": active_document_part,
                                "active_document_sha256": sha256(active_document),
                                "storage": storage,
                                "workbook_storage_kind": storage_kind,
                                "workbook_sha256": sha256(workbook),
                                "workbook_bytes": len(workbook),
                                **(
                                    {
                                        "source_workbook_stream_sha256": sha256(source_workbook_stream),
                                        "source_workbook_stream_bytes": len(source_workbook_stream),
                                    }
                                    if source_workbook_stream is not None
                                    else {}
                                ),
                            },
                        )
                    )

    if len(matches) != 1:
        raise ValueError(
            f"Expected exactly one embedded datasheet for named element '{name}' on slide {slide_number}; found {len(matches)}."
        )
    return matches[0]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract the embedded datasheet for one genuine named think-cell element."
    )
    parser.add_argument("presentation", type=Path)
    parser.add_argument("--slide", type=int, required=True)
    parser.add_argument("--name", required=True)
    parser.add_argument("--output-xlsb", type=Path, required=True)
    parser.add_argument("--output-xls", type=Path)
    parser.add_argument("--output-json", type=Path)
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()

    presentation = args.presentation.resolve()
    output_xlsb = args.output_xlsb.resolve()
    output_xls = args.output_xls.resolve() if args.output_xls else None
    output_json = args.output_json.resolve() if args.output_json else None
    if (
        output_xlsb == presentation
        or (output_xls and output_xls == presentation)
        or (output_json and output_json == presentation)
    ):
        raise SystemExit("Outputs must be distinct from the presentation input.")
    if output_xls and output_xls == output_xlsb:
        raise SystemExit("--output-xls and --output-xlsb must be distinct paths.")
    for output in [output_xlsb, output_xls, output_json]:
        if output and output.exists() and not args.force:
            raise SystemExit(f"Output already exists: {output}")

    try:
        workbook, storage_kind, report = extract_named_datasheet(
            presentation, args.slide, args.name
        )
    except (ValueError, zipfile.BadZipFile, etree.XMLSyntaxError, OSError) as exc:
        raise SystemExit(str(exc)) from exc

    if storage_kind == "legacy_biff_cfb":
        if output_xls is None:
            raise SystemExit(
                "This named element uses a legacy BIFF Workbook stream; pass --output-xls for its reconstructed workbook."
            )
        output_workbook = output_xls
    else:
        output_workbook = output_xlsb
    output_workbook.parent.mkdir(parents=True, exist_ok=True)
    output_workbook.write_bytes(workbook)
    report["output_workbook"] = str(output_workbook)
    if storage_kind == "xlsb_package":
        report["output_xlsb"] = str(output_workbook)
    else:
        report["output_xls"] = str(output_workbook)
    payload = json.dumps(report, indent=2)
    if output_json:
        output_json.parent.mkdir(parents=True, exist_ok=True)
        output_json.write_text(payload, encoding="utf-8")
    print(payload)


if __name__ == "__main__":
    main()
