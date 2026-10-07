"""Bounded ordinary-object edits on a source slide, without touching native parts.

Selectors bind a top-level PowerPoint shape ID and canonical XML SHA-256.
All edits are validated against the original source, then applied together.
Geometry uses integer EMUs. Text replacement binds every existing text run;
formatting and paragraph structure remain intact. Additions are ordinary,
editable PowerPoint rectangles/text boxes with explicit style and geometry.
"""
import copy
import hashlib
import io
import re
import zipfile
from lxml import etree as E

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"p": P, "a": A}


def need(condition, message):
    if not condition:
        raise ValueError(message)


def shape_sha(shape):
    return hashlib.sha256(E.tostring(shape, method="c14n")).hexdigest().upper()


def protection(shape):
    if shape.tag != "{" + P + "}sp":
        return "Only ordinary top-level PowerPoint shapes are editable"
    if shape.findall(".//p:tags", NS):
        return "Tagged shapes may be native-owned"
    if shape.findall(".//a:fld", NS):
        return "Field-linked shapes are protected"
    if any(key.startswith("{" + R + "}") for node in shape.iter() for key in node.attrib):
        return "Relationship-bearing shapes are protected"
    harmless = {"{http://schemas.microsoft.com/office/drawing/2014/main}creationId",
                "{http://schemas.microsoft.com/office/powerpoint/2010/main}creationId",
                "{http://schemas.microsoft.com/office/powerpoint/2010/main}modId"}
    for extension in shape.findall(".//a:extLst", NS) + shape.findall(".//p:extLst", NS):
        if any(leaf.tag not in harmless for record in extension for leaf in record):
            return "Unknown extension-bearing shapes require a dedicated adapter"
    return None


def _tree(raw):
    root = E.fromstring(raw, E.XMLParser(resolve_entities=False))
    tree = root.find("p:cSld/p:spTree", NS)
    need(tree is not None, "Slide shape tree missing")
    # An OLE object's fallback picture repeats its carrier frame identity.
    # It is a representation inside that object, not another slide shape.
    ids = [node.get("id") for node in tree.findall(".//p:cNvPr", NS)
           if not any(parent.tag == "{" + P + "}oleObj" for parent in node.iterancestors())]
    need(all(value and re.fullmatch(r"[1-9][0-9]*", value) for value in ids), "Invalid shape ID")
    need(len(set(ids)) == len(ids), "Duplicate shape IDs")
    shapes = {}
    for node in tree:
        identity = node.find(".//p:cNvPr", NS)
        if identity is not None:
            shapes[int(identity.get("id"))] = node
    return root, tree, shapes, {int(value) for value in ids}


def inspect_shapes(raw):
    """Read-only selectors and protection reasons for a slide XML part."""
    _, _, shapes, _ = _tree(raw)
    return [{"shape_id": key, "sha256": shape_sha(node),
             "name": node.find(".//p:cNvPr", NS).get("name"),
             "text_runs": [text.text or "" for text in node.findall(".//a:t", NS)],
             "protected_reason": protection(node)} for key, node in shapes.items()]


def _bounds(value):
    need(isinstance(value, dict) and set(value) == {"left", "top", "width", "height"}, "Invalid bounds_emu")
    need(all(type(item) is int for item in value.values()), "Geometry must use integer EMUs")
    need(value["width"] > 0 and value["height"] > 0, "Geometry dimensions must be positive")
    need(all(abs(item) <= 2147483647 for item in value.values()), "Geometry out of bounds")
    return value


def _set_bounds(shape, bounds):
    transform = shape.find("p:spPr/a:xfrm", NS)
    need(transform is not None and transform.find("a:off", NS) is not None and transform.find("a:ext", NS) is not None,
         "Shape must have explicit geometry")
    transform.find("a:off", NS).attrib.update({"x": str(bounds["left"]), "y": str(bounds["top"])})
    transform.find("a:ext", NS).attrib.update({"cx": str(bounds["width"]), "cy": str(bounds["height"])})


def _add(entry):
    need(set(entry) == {"op", "shape_id", "name", "kind", "bounds_emu", "text", "font", "fill", "line"}, "Invalid add operation")
    need(type(entry["shape_id"]) is int and 0 < entry["shape_id"] <= 2147483647, "Invalid added shape ID")
    need(isinstance(entry["name"], str) and entry["name"].strip(), "Added shape requires a name")
    need(entry["kind"] in {"textbox", "rectangle"}, "Unsupported ordinary shape kind")
    need(isinstance(entry["text"], str), "Added text must be a string")
    bounds = _bounds(entry["bounds_emu"])
    font = entry["font"]
    need(isinstance(font, dict) and set(font) == {"name", "size", "color", "bold"}, "Invalid font")
    need(isinstance(font["name"], str) and font["name"].strip(), "Font name missing")
    need(type(font["size"]) is int and 100 <= font["size"] <= 400000, "Font size must use hundredths of a point")
    need(type(font["bold"]) is bool, "Font bold must be Boolean")
    for color in (font["color"], entry["fill"], entry["line"]):
        need(color is None or isinstance(color, str) and re.fullmatch(r"[0-9A-Fa-f]{6}", color), "Colors must be six hex digits or null")
    need(font["color"] is not None, "Font color is required")
    shape = E.Element("{" + P + "}sp")
    nv = E.SubElement(shape, "{" + P + "}nvSpPr")
    E.SubElement(nv, "{" + P + "}cNvPr", id=str(entry["shape_id"]), name=entry["name"])
    E.SubElement(nv, "{" + P + "}cNvSpPr", txBox="1" if entry["kind"] == "textbox" else "0")
    E.SubElement(nv, "{" + P + "}nvPr")
    props = E.SubElement(shape, "{" + P + "}spPr")
    transform = E.SubElement(props, "{" + A + "}xfrm")
    E.SubElement(transform, "{" + A + "}off", x=str(bounds["left"]), y=str(bounds["top"]))
    E.SubElement(transform, "{" + A + "}ext", cx=str(bounds["width"]), cy=str(bounds["height"]))
    E.SubElement(E.SubElement(props, "{" + A + "}prstGeom", prst="rect"), "{" + A + "}avLst")
    def color_node(parent, color):
        if color is None:
            E.SubElement(parent, "{" + A + "}noFill")
        else:
            E.SubElement(E.SubElement(parent, "{" + A + "}solidFill"), "{" + A + "}srgbClr", val=color.upper())
    color_node(props, entry["fill"])
    color_node(E.SubElement(props, "{" + A + "}ln"), entry["line"])
    body = E.SubElement(shape, "{" + P + "}txBody")
    E.SubElement(body, "{" + A + "}bodyPr")
    E.SubElement(body, "{" + A + "}lstStyle")
    for line in entry["text"].split("\n"):
        paragraph = E.SubElement(body, "{" + A + "}p")
        run = E.SubElement(paragraph, "{" + A + "}r")
        style = E.SubElement(run, "{" + A + "}rPr", sz=str(font["size"]), b="1" if font["bold"] else "0")
        color_node(style, font["color"])
        E.SubElement(style, "{" + A + "}latin", typeface=font["name"])
        E.SubElement(run, "{" + A + "}t").text = line
    return shape


def edit_slide(raw, edits):
    need(isinstance(edits, list), "Edits must be a list")
    root, tree, shapes, occupied = _tree(raw)
    seen, additions, planned = set(), set(), []
    for entry in edits:
        need(isinstance(entry, dict) and entry.get("op") in {"set", "remove", "add"}, "Invalid ordinary object operation")
        if entry["op"] == "add":
            node = _add(entry)
            identity = entry["shape_id"]
            need(identity not in occupied and identity not in additions, "Added shape ID collides with source or another addition")
            additions.add(identity)
            planned.append((entry, node))
            continue
        required = {"op", "shape_id", "sha256"}
        need(required.issubset(entry) and set(entry).issubset(required | {"text_runs", "bounds_emu"}), "Invalid edit selector/fields")
        need(type(entry["shape_id"]) is int and entry["shape_id"] in shapes, "Shape selector did not resolve")
        identity = entry["shape_id"]
        need(identity not in seen, "Repeated shape edit selector")
        seen.add(identity)
        shape = shapes[identity]
        need(isinstance(entry["sha256"], str) and re.fullmatch(r"[0-9A-Fa-f]{64}", entry["sha256"]), "Malformed shape hash")
        need(shape_sha(shape) == entry["sha256"].upper(), "Shape SHA-256 precondition mismatch")
        reason = protection(shape)
        need(reason is None, reason or "Protected shape")
        if entry["op"] == "remove":
            need(set(entry) == required, "Remove accepts only an exact selector")
            need(not shape.findall(".//p:ph", NS), "Inherited placeholders cannot be removed")
            need(not root.xpath(".//a:stCxn[@id=$id] | .//a:endCxn[@id=$id]", id=str(identity), namespaces=NS), "Shape is referenced by a connector")
            need(not any(node.get("spid") == str(identity) for node in root.iter()), "Shape is referenced by slide timing or build settings")
        else:
            need(set(entry) != required, "Set operation has no changes")
            if "text_runs" in entry:
                runs = shape.findall(".//a:t", NS)
                texts = entry["text_runs"]
                need(isinstance(texts, list) and len(texts) == len(runs) and len(runs) > 0 and all(isinstance(text, str) for text in texts), "Text must bind every existing text run")
            if "bounds_emu" in entry:
                _set_bounds(copy.deepcopy(shape), _bounds(entry["bounds_emu"]))
        planned.append((entry, shape))
    for entry, shape in planned:
        if entry["op"] == "add":
            extension = tree.find("p:extLst", NS)
            if extension is None:
                tree.append(shape)
            else:
                tree.insert(tree.index(extension), shape)
        elif entry["op"] == "remove":
            tree.remove(shape)
        else:
            if "text_runs" in entry:
                for text, value in zip(shape.findall(".//a:t", NS), entry["text_runs"]):
                    text.text = value
            if "bounds_emu" in entry:
                _set_bounds(shape, entry["bounds_emu"])
    _tree(E.tostring(root))
    return E.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def prepare_source(data, edits, slide_part):
    """Return a new package, retaining every other part byte-for-byte."""
    if not edits:
        return data
    source, destination = io.BytesIO(data), io.BytesIO()
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(destination, "w") as after:
        need(len(set(before.namelist())) == len(before.namelist()), "Duplicate ZIP entries")
        changed = edit_slide(before.read(slide_part), edits)
        for item in before.infolist():
            after.writestr(item, changed if item.filename == slide_part else before.read(item.filename))
    return destination.getvalue()


def main():
    import argparse
    import json
    from pathlib import Path
    from audit_thinkcell_integrity import logical_slides
    parser = argparse.ArgumentParser(description="Inspect exact ordinary-object assembly selectors without writing")
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--expected-sha256", required=True)
    args = parser.parse_args()
    data = args.input.read_bytes()
    need(hashlib.sha256(data).hexdigest().upper() == args.expected_sha256.upper(), "Source SHA-256 mismatch")
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        need(len(package.namelist()) == len(set(package.namelist())), "Duplicate ZIP entries")
        slides = logical_slides(package)
        need(len(slides) == 1, "Use an exact one-slide source")
        print(json.dumps({"source_sha256": args.expected_sha256.upper(), "slide": slides[0],
                          "objects": inspect_shapes(package.read(slides[0]["part"]))}, indent=2))


if __name__ == "__main__":
    main()
