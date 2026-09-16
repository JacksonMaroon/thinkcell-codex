"""Read-only native percent-label trace, including the scalar's owning graph.

The report deliberately records XML rather than inferring label semantics from
the rendered percentage.  It is usable on the Format Library donor and on a
candidate after each regeneration/native reopen.
"""
from __future__ import annotations

import argparse
import io
import json
import sys
import zipfile
from pathlib import Path

from lxml import etree as E

SCRIPTS = Path(__file__).resolve().parent
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / "thinkcell_no_click" / "implementation")]
from chart_geometry import streams, xml  # noqa: E402
from audit_thinkcell_integrity import NS, logical_slides, relationship_map  # noqa: E402


def stringify(node):
    return E.tostring(node, encoding="unicode") if node is not None else None


def field_shapes(data, shape_tag):
    found = []
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        for slide in logical_slides(package):
            root = xml(package.read(slide["part"]))
            relationships = relationship_map(package, slide["part"])
            for shape in root.findall(".//p:sp", NS):
                tags = []
                for ref in shape.findall(".//p:tags", NS):
                    target = relationships[ref.get("{" + NS["r"] + "}id")]["resolved"]
                    tags.extend(n.get("val") for n in xml(package.read(target))
                                if n.get("name", "").upper() == "THINKCELLSHAPEDONOTDELETE")
                if shape_tag in tags:
                    found.append({
                        "slide_id": slide["id"],
                        "shape_id": shape.find(".//p:cNvPr", NS).get("id"),
                        "shape_name": shape.find(".//p:cNvPr", NS).get("name"),
                        "fields": [{"id": f.get("id"), "type": f.get("type"),
                                    "text": f.findtext("a:t", namespaces=NS)}
                                   for f in shape.findall(".//a:fld", NS)],
                        "literal_texts": [n.text for n in shape.findall(".//a:t", NS)],
                        "shape_xml": stringify(shape),
                    })
    return found


def trace(path, carrier, scalar_id, shape_tag):
    raw = Path(path).read_bytes()
    with zipfile.ZipFile(io.BytesIO(raw)) as package:
        carrier_bytes = package.read(carrier)
    root = xml(streams(carrier_bytes)[("think-cellXML",)])
    ids = {n.get("id"): n for n in root if n.get("id")}
    scalar = ids[str(scalar_id)]
    parents = [n for n in root.iter() if scalar in n]
    label_ref = scalar.find("m_scdlabel")
    label = ids.get(label_ref.get("idref")) if label_ref is not None else None
    source_rows = {}
    for name in ("m_varsrcAbsolute", "m_varsrcAccuAbsolute", "m_varsrcRelative", "m_varsrcAccuRelative"):
        ref = scalar.find(name)
        source = ids.get(ref.get("idref")) if ref is not None else None
        field = source.find("m_ctextvar/elem") if source is not None else None
        text = ids.get(field.get("idref")) if field is not None else None
        source_rows[name] = {"ref": ref.get("idref") if ref is not None else None,
                             "xml": stringify(source), "text_variable_xml": stringify(text)}
    return {"input": str(Path(path).resolve()), "carrier": carrier, "scalar_id": str(scalar_id),
            "shape_tag": shape_tag, "scalar_xml": stringify(scalar), "label_xml": stringify(label),
            "parent_xml": [stringify(p) for p in parents], "sources": source_rows,
            "visible_shapes": field_shapes(raw, shape_tag)}


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--carrier", required=True)
    parser.add_argument("--scalar", required=True)
    parser.add_argument("--shape-tag", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()
    Path(args.out).write_text(json.dumps(trace(args.input, args.carrier, args.scalar, args.shape_tag), indent=2), encoding="utf-8")
