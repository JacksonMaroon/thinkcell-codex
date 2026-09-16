from __future__ import annotations

import argparse
import hashlib
import io
import json
import math
import posixpath
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path, PurePosixPath

from lxml import etree


VENDOR = Path(__file__).resolve().parent / "vendor"
sys.path.insert(0, str(VENDOR))
try:
    import olefile  # type: ignore
except ModuleNotFoundError as exc:  # pragma: no cover - installation error
    raise SystemExit(f"Missing olefile dependency; install scripts/requirements.txt: {exc}")


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
R = "{" + NS["r"] + "}"
ACTIVE_DOC_TAG_VALUE = "thinkcellActiveDocDoNotDelete"
THINKCELL_TAG_NAME = "THINKCELLSHAPEDONOTDELETE"


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def rels_part(source_part: str) -> str:
    part = PurePosixPath(source_part)
    return str(part.parent / "_rels" / (part.name + ".rels"))


def source_from_rels(rels_name: str) -> str:
    if rels_name == "_rels/.rels":
        return ""
    path = PurePosixPath(rels_name)
    return str(path.parent.parent / path.name[:-5])


def resolve_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def relationship_map(package: zipfile.ZipFile, source_part: str) -> dict[str, dict[str, str]]:
    name = rels_part(source_part)
    if name not in package.namelist():
        return {}
    result: dict[str, dict[str, str]] = {}
    for rel in etree.fromstring(package.read(name)):
        target = rel.get("Target", "")
        result[rel.get("Id")] = {
            "type": rel.get("Type", ""),
            "target": target,
            "mode": rel.get("TargetMode", ""),
            "resolved": resolve_target(source_part, target),
        }
    return result


def logical_slides(package: zipfile.ZipFile) -> list[dict[str, object]]:
    presentation = etree.fromstring(package.read("ppt/presentation.xml"))
    rels = relationship_map(package, "ppt/presentation.xml")
    slides = []
    for node in presentation.xpath("./p:sldIdLst/p:sldId", namespaces=NS):
        rid = node.get(R + "id")
        slides.append({"id": int(node.get("id")), "rid": rid, "part": rels[rid]["resolved"]})
    return slides


def related_part(package: zipfile.ZipFile, source_part: str, suffix: str) -> str | None:
    for rel in relationship_map(package, source_part).values():
        if rel["type"].endswith("/" + suffix):
            return rel["resolved"]
    return None


def tag_values(package: zipfile.ZipFile, part: str) -> list[dict[str, str]]:
    if part not in package.namelist():
        return []
    try:
        root = etree.fromstring(package.read(part))
    except etree.XMLSyntaxError:
        return []
    return [{"name": node.get("name", ""), "value": node.get("val", "")} for node in root.xpath(".//p:tag", namespaces=NS)]


def dangling_relationships(package: zipfile.ZipFile) -> list[dict[str, str]]:
    names = set(package.namelist())
    errors = []
    for rels_name in sorted(name for name in names if name.endswith(".rels")):
        source = source_from_rels(rels_name)
        try:
            root = etree.fromstring(package.read(rels_name))
        except etree.XMLSyntaxError as exc:
            errors.append({"rels": rels_name, "id": "", "target": f"XML error: {exc}"})
            continue
        for rel in root:
            if rel.get("TargetMode") == "External":
                continue
            target = resolve_target(source, rel.get("Target", ""))
            if target not in names:
                errors.append({"rels": rels_name, "id": rel.get("Id", ""), "target": target})
    return errors


def duplicate_relationship_ids(package: zipfile.ZipFile) -> list[dict[str, object]]:
    duplicates = []
    for rels_name in sorted(name for name in package.namelist() if name.endswith(".rels")):
        try:
            root = etree.fromstring(package.read(rels_name))
        except etree.XMLSyntaxError:
            continue
        counts = Counter(rel.get("Id", "") for rel in root)
        for rel_id, count in counts.items():
            if rel_id and count > 1:
                duplicates.append({"rels": rels_name, "id": rel_id, "count": count})
    return duplicates


def shared_embedding_owners(package: zipfile.ZipFile) -> list[dict[str, object]]:
    owners: dict[str, set[str]] = {}
    for rels_name in sorted(name for name in package.namelist() if name.endswith(".rels")):
        source = source_from_rels(rels_name)
        try:
            root = etree.fromstring(package.read(rels_name))
        except etree.XMLSyntaxError:
            continue
        for rel in root:
            if rel.get("TargetMode") == "External":
                continue
            target = resolve_target(source, rel.get("Target", ""))
            if not target.startswith("ppt/embeddings/"):
                continue
            if not source.startswith(
                ("ppt/charts/", "ppt/slides/", "ppt/slideMasters/", "ppt/slideLayouts/")
            ):
                continue
            owners.setdefault(target, set()).add(source)
    return [
        {"embedding": target, "owners": sorted(source_parts)}
        for target, source_parts in sorted(owners.items())
        if len(source_parts) > 1
        # think-cell's own template deliberately reuses some layout-level
        # ActiveDocument payloads across multiple slide layouts.  That is a
        # valid package pattern and is preserved in the embedding inventory.
        # Sharing a payload between slides, charts, a master, or mixed owner
        # kinds remains unsafe and must fail closed.
        and not all(source.startswith("ppt/slideLayouts/") for source in source_parts)
    ]


def embedding_inventory(package: zipfile.ZipFile) -> list[dict[str, object]]:
    names = set(package.namelist())
    records: dict[str, dict[str, object]] = {
        name: {"part": name, "owners": []}
        for name in sorted(names)
        if name.startswith("ppt/embeddings/") and not name.endswith("/")
    }
    xml_cache: dict[str, etree._Element | None] = {}

    def source_kind(source: str) -> str:
        for prefix, kind in (
            ("ppt/slides/", "slide"),
            ("ppt/slideMasters/", "master"),
            ("ppt/slideLayouts/", "layout"),
            ("ppt/charts/", "chart"),
            ("ppt/notesSlides/", "notes"),
        ):
            if source.startswith(prefix):
                return kind
        return "other"

    for rels_name in sorted(name for name in names if name.endswith(".rels")):
        source = source_from_rels(rels_name)
        try:
            relationships = etree.fromstring(package.read(rels_name))
        except etree.XMLSyntaxError:
            continue
        for relationship in relationships:
            if relationship.get("TargetMode") == "External":
                continue
            target = resolve_target(source, relationship.get("Target", ""))
            if target not in records:
                continue
            relationship_kind = relationship.get("Type", "").rsplit("/", 1)[-1]
            owner: dict[str, object] = {
                "source": source,
                "source_kind": source_kind(source),
                "relationship_id": relationship.get("Id", ""),
                "relationship_kind": relationship_kind,
            }
            if source in names and source.endswith(".xml"):
                if source not in xml_cache:
                    try:
                        xml_cache[source] = etree.fromstring(package.read(source))
                    except etree.XMLSyntaxError:
                        xml_cache[source] = None
                source_root = xml_cache[source]
                if source_root is not None:
                    relationship_id = relationship.get("Id", "")
                    nodes = source_root.xpath(
                        ".//*[@r:id=$relationship_id]",
                        namespaces=NS,
                        relationship_id=relationship_id,
                    )
                    if nodes:
                        owner.update(
                            {
                                "element": etree.QName(nodes[0]).localname,
                                "prog_id": nodes[0].get("progId"),
                                "name": nodes[0].get("name"),
                            }
                        )
            if (
                owner["source_kind"] == "master"
                and relationship_kind == "oleObject"
                and str(owner.get("prog_id", "")).startswith("TCLayout.ActiveDocument")
            ):
                owner["role"] = "master_thinkcell_active_document"
            elif (
                owner["source_kind"] == "slide"
                and relationship_kind == "oleObject"
                and str(owner.get("prog_id", "")).startswith("TCLayout.ActiveDocument")
            ):
                owner["role"] = "slide_thinkcell_active_document"
            elif owner["source_kind"] == "chart" and relationship_kind == "package":
                owner["role"] = "chart_package"
            else:
                owner["role"] = f"{owner['source_kind']}_{relationship_kind or 'relationship'}"
            records[target]["owners"].append(owner)

    for record in records.values():
        owners = record["owners"]
        roles = sorted({str(owner["role"]) for owner in owners})
        record["roles"] = roles
        record["bytes"] = len(package.read(str(record["part"])))
    return list(records.values())


def missing_content_types(package: zipfile.ZipFile) -> list[str]:
    names = set(package.namelist())
    root = etree.fromstring(package.read("[Content_Types].xml"))
    defaults = {node.get("Extension", "").lower() for node in root if etree.QName(node).localname == "Default"}
    overrides = {node.get("PartName", "").lstrip("/") for node in root if etree.QName(node).localname == "Override"}
    missing = []
    for name in names:
        if name.endswith("/") or name == "[Content_Types].xml" or name.endswith(".rels"):
            continue
        extension = PurePosixPath(name).suffix.lstrip(".").lower()
        if name not in overrides and extension not in defaults:
            missing.append(name)
    return sorted(missing)


def invalid_mc_prefixes(package: zipfile.ZipFile) -> list[dict[str, object]]:
    invalid = []
    for name in sorted(part for part in package.namelist() if part.endswith((".xml", ".rels"))):
        try:
            root = etree.fromstring(package.read(name))
        except etree.XMLSyntaxError as exc:
            invalid.append({"part": name, "parse_error": str(exc)})
            continue
        for element in root.iter():
            value = element.get("{" + NS["mc"] + "}Ignorable")
            if not value:
                continue
            missing = [prefix for prefix in value.split() if element.nsmap.get(prefix) is None]
            if missing:
                invalid.append({"part": name, "element": element.tag, "missing": missing})
    return invalid


def notes_substance(package: zipfile.ZipFile, slide_part: str) -> str | None:
    notes = related_part(package, slide_part, "notesSlide")
    if notes is None:
        return None
    root = etree.fromstring(package.read(notes))
    rels = relationship_map(package, notes)
    ignored_placeholders = {"sldImg", "dt", "ftr", "hdr", "sldNum"}
    authored = []
    for node in root.xpath(".//p:sp | .//p:pic | .//p:graphicFrame", namespaces=NS):
        placeholders = node.xpath(".//p:nvPr/p:ph[1]", namespaces=NS)
        placeholder = placeholders[0].get("type", "body") if placeholders else None
        if placeholder in ignored_placeholders:
            continue
        paragraphs = ["".join(paragraph.xpath(".//a:t/text()", namespaces=NS)) for paragraph in node.xpath(".//a:p", namespaces=NS)]
        relationships = []
        for descendant in node.iter():
            for key, value in descendant.attrib.items():
                if not key.startswith(R) or value not in rels:
                    continue
                rel = rels[value]
                entry: dict[str, object] = {"type": rel["type"], "target_mode": rel["mode"]}
                if rel["mode"] == "External":
                    entry["target"] = rel["target"]
                elif rel["resolved"] in package.namelist():
                    entry["content_sha256"] = sha256(package.read(rel["resolved"]))
                relationships.append(entry)
        authored.append(
            {
                "kind": etree.QName(node).localname,
                "placeholder": placeholder,
                "paragraphs": paragraphs,
                "relationships": relationships,
            }
        )
    payload = {"notes_slide_present": True, "authored_content": authored}
    return sha256(json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8"))


def _id_map(root: etree._Element) -> dict[str, etree._Element]:
    return {node.get("id"): node for node in root if node.get("id")}


def _variable_value(ids: dict[str, etree._Element], owner: etree._Element, source_tag: str) -> object | None:
    reference = owner.find(source_tag)
    if reference is None:
        return None
    source_reference = reference.get("idref")
    if source_reference:
        source = ids.get(source_reference)
        if source is None:
            return None
    else:
        source = reference
    value = source.find("m_varval")
    if value is None:
        return None
    raw = value.text or value.get("val")
    if raw is None:
        return None
    try:
        return float(raw)
    except ValueError:
        return raw.strip()


def _node_value(owner: etree._Element | None, tag: str) -> str | None:
    if owner is None:
        return None
    node = owner.find(tag)
    if node is None:
        return None
    raw = node.get("val") if node.get("val") is not None else node.text
    return raw.strip() if isinstance(raw, str) else None


_CURRENT_PERCENT_AXIS_VERSION = 38764


def _typed_axis_value(node: etree._Element | None, expected: float) -> bool:
    """Current regeneration's bounded replacement for legacy m_bPercentage."""
    try:
        return node is not None and node.get("type") == "1" and float(node.get("val")) == expected
    except (TypeError, ValueError):
        return False


def _percent_axis(chart: etree._Element, ids: dict[str, etree._Element], version: int | None) -> bool | None:
    axis_reference = chart.find("m_daxisPrimaryValue")
    axis = ids.get(axis_reference.get("idref")) if axis_reference is not None else None
    percent_value = _node_value(axis, "m_bPercentage")
    if percent_value is not None:
        return percent_value not in {"0", "false", "False"}
    precision = axis.find("m_precUser") if axis is not None else None
    migrated_current_signature = (
        axis is not None and axis.tag == "CSequenceChartDataAxis"
        and _typed_axis_value(axis.find("m_fMinValue"), 0.0)
        and _typed_axis_value(axis.find("m_fMaxValue"), 1.0)
        and _typed_axis_value(axis.find("m_fUserScaleUnit"), 0.5)
        and _node_value(axis, "m_edaxistype") == "1"
        and _node_value(precision, "m_strSuffix17909") == "%"
    )
    return True if version is not None and version >= _CURRENT_PERCENT_AXIS_VERSION and migrated_current_signature else None


def sequence_tables(root: etree._Element) -> list[dict[str, object]]:
    ids = _id_map(root)
    try:
        version = int(root.find("version").get("val"))
    except (AttributeError, TypeError, ValueError):
        version = None
    charts_by_table: dict[str, etree._Element] = {}
    for chart in root.findall("CSequenceChartSE"):
        table_reference = chart.find("m_dtable")
        if table_reference is not None and table_reference.get("idref"):
            charts_by_table[table_reference.get("idref")] = chart
    result = []
    for table in root.findall("CSequenceChartDataTable"):
        series_container = table.find("m_cscdser")
        series_ids = [node.get("idref") for node in series_container] if series_container is not None else []
        names = []
        for series_id in series_ids:
            series = ids.get(series_id)
            names.append(_variable_value(ids, series, "m_varsrc") if series is not None else None)
        categories = []
        category_extents = []
        category_values: list[list[float | None]] = []
        for reference in table.findall("./ocol/elem"):
            vector = ids.get(reference.get("idref"))
            if vector is None:
                continue
            categories.append(_variable_value(ids, vector, "m_varsrcCategory"))
            category_extents.append(_variable_value(ids, vector, "m_varsrcAbsoluteExtent"))
            row: list[float | None] = []
            for scalar_reference in vector.findall("./ocol/elem"):
                scalar = ids.get(scalar_reference.get("idref"))
                raw = _variable_value(ids, scalar, "m_varsrcAbsolute") if scalar is not None else None
                row.append(float(raw) if isinstance(raw, (int, float)) else None)
            category_values.append(row)
        series_values = []
        for index in range(len(series_ids)):
            series_values.append([row[index] if index < len(row) else None for row in category_values])
        name_node = table.find("m_strName")
        automation_name = ((name_node.text or name_node.get("val") or "").strip() if name_node is not None else "")
        chart = charts_by_table.get(table.get("id") or "")
        orientation_code = _node_value(chart, "m_eorient")
        orientation = {"0": "vertical", "1": "horizontal"}.get(orientation_code)
        percent_axis = _percent_axis(chart, ids, version) if chart is not None else None
        result.append(
            {
                "id": table.get("id"),
                "automation_name": automation_name,
                "series_names": names,
                "categories": categories,
                "series_values": series_values,
                "category_extents": category_extents,
                "orientation": orientation,
                "percent_axis": percent_axis,
            }
        )
    return result


def scatter_tables(root: etree._Element) -> list[dict[str, object]]:
    """Extract exact point-row data from think-cell scatter and bubble models.

    CScatterChartDataTable stores one CScatterChartDataVector per populated
    datasheet row.  Each vector independently references its point label,
    group/color key, X value, Y value, and optional bubble size.  Preserve the
    raw row order and raw group cells because blank group cells can carry
    donor-specific formatting semantics that should not be guessed here.
    """
    ids = _id_map(root)
    result = []
    for chart in root.findall("CScatterChartSE"):
        table_reference = chart.find("m_dtable")
        table = (
            ids.get(table_reference.get("idref"))
            if table_reference is not None and table_reference.get("idref")
            else None
        )
        if table is None or etree.QName(table).localname != "CScatterChartDataTable":
            continue

        points = []
        for reference in table.findall("./ocol/elem"):
            vector = ids.get(reference.get("idref"))
            if vector is None or etree.QName(vector).localname != "CScatterChartDataVector":
                continue
            raw_x = _variable_value(ids, vector, "m_varsrcX")
            raw_y = _variable_value(ids, vector, "m_varsrcY")
            raw_size = _variable_value(ids, vector, "m_varsrcSize")
            # Separator, header, or total rows can exist as vectors but are
            # not plotted unless both X and Y are numeric.  Keep the audited
            # point set aligned to the visible native chart cache.
            if not isinstance(raw_x, (int, float)) or not isinstance(raw_y, (int, float)):
                continue
            points.append(
                {
                    "label": _variable_value(ids, vector, "m_varsrcText"),
                    "group": _variable_value(ids, vector, "m_varsrcGroup"),
                    "x": float(raw_x) if isinstance(raw_x, (int, float)) else None,
                    "y": float(raw_y) if isinstance(raw_y, (int, float)) else None,
                    "size": float(raw_size) if isinstance(raw_size, (int, float)) else None,
                }
            )

        name_node = chart.find("m_strName")
        automation_name = (
            (name_node.text or name_node.get("val") or "").strip()
            if name_node is not None
            else ""
        )
        result.append(
            {
                "id": table.get("id"),
                "chart_id": chart.get("id"),
                "automation_name": automation_name,
                "point_labels": [point["label"] for point in points],
                "group_labels": [point["group"] for point in points],
                "x_values": [point["x"] for point in points],
                "y_values": [point["y"] for point in points],
                "size_values": [point["size"] for point in points],
                "points": points,
                "point_count": len(points),
                "z_value_is_area": _node_value(chart, "m_bZValueIsArea"),
                "storage": _node_value(table, "m_bstrRangeName"),
            }
        )
    return result


def pie_tables(root: etree._Element) -> list[dict[str, object]]:
    """Extract exact slice order and raw values from think-cell pie models.

    CPieChartDataScalar stores the category label and absolute value on each
    slice.  m_nIndexInDataSheet is the authoritative order; XML element order
    is not assumed.  Native PowerPoint pie caches commonly contain normalized
    percentages, so both raw values and calculated percentages are retained.
    """
    ids = _id_map(root)
    charts_by_table: dict[str, etree._Element] = {}
    for chart in root.findall("CPieChartSE"):
        table_reference = chart.find("m_dtable")
        if table_reference is not None and table_reference.get("idref"):
            charts_by_table[table_reference.get("idref")] = chart

    result = []
    for table in root.findall("CPieChartDataTable"):
        chart = charts_by_table.get(table.get("id") or "")
        vector_references = table.findall("./ocol/elem")
        vectors = [ids.get(reference.get("idref")) for reference in vector_references]
        vectors = [vector for vector in vectors if vector is not None]
        if len(vectors) != 1:
            # Pie certification requires one value vector.  Keep malformed or
            # unsupported models visible to the audit rather than guessing.
            result.append(
                {
                    "id": table.get("id"),
                    "chart_id": chart.get("id") if chart is not None else None,
                    "automation_name": "",
                    "series_count": len(vectors),
                    "categories": [],
                    "values": [],
                    "total": None,
                    "calculated_total": None,
                    "percentages": [],
                    "hole_percent": None,
                    "exploded": [],
                    "storage": _node_value(table, "m_bstrRangeName"),
                }
            )
            continue

        vector = vectors[0]
        ordered_scalars = []
        for position, reference in enumerate(vector.findall("./ocol/elem")):
            scalar = ids.get(reference.get("idref"))
            if scalar is None:
                continue
            raw_index = _node_value(scalar, "m_nIndexInDataSheet")
            try:
                data_index = int(raw_index) if raw_index is not None else position
            except ValueError:
                data_index = position
            ordered_scalars.append((data_index, position, scalar))
        ordered_scalars.sort(key=lambda item: (item[0], item[1]))

        categories: list[object | None] = []
        values: list[float | None] = []
        exploded: list[bool] = []
        input_modes = []
        for _, _, scalar in ordered_scalars:
            categories.append(_variable_value(ids, scalar, "m_varsrcSeries"))
            raw_value = _variable_value(ids, scalar, "m_varsrcAbsolute")
            relative = _variable_value(ids, scalar, "m_varsrcRelative")
            input_modes.append('percentage' if raw_value is None and isinstance(relative,(int,float)) else 'absolute')
            if input_modes[-1]=='percentage':raw_value=relative
            values.append(float(raw_value) if isinstance(raw_value, (int, float)) else None)
            raw_exploded = _node_value(scalar, "m_bExploded")
            exploded.append(raw_exploded not in {None, "0", "false", "False"})

        numeric_values = [value for value in values if value is not None and math.isfinite(value)]
        calculated_total = sum(numeric_values) if len(numeric_values) == len(values) else None
        model_total = _variable_value(ids, vector, "m_varsrcAbsoluteSum")
        total = float(model_total) if isinstance(model_total, (int, float)) else None
        percentages = (
            [float(value) * 100.0 / calculated_total for value in values if value is not None]
            if calculated_total not in {None, 0}
            else []
        )
        name_node = chart.find("m_strName") if chart is not None else None
        automation_name = (
            (name_node.text or name_node.get("val") or "").strip()
            if name_node is not None
            else ""
        )
        result.append(
            {
                "id": table.get("id"),
                "chart_id": chart.get("id") if chart is not None else None,
                "automation_name": automation_name,
                "series_count": 1,
                "categories": categories,
                "values": values,
                "total": total,
                "calculated_total": calculated_total,
                "percentages": percentages,
                "input_mode": input_modes[0] if len(set(input_modes))==1 else 'mixed',
                "hole_percent": _node_value(vector, "m_nHolePercent"),
                "exploded": exploded,
                "storage": _node_value(table, "m_bstrRangeName"),
            }
        )
    return result


def inspect_thinkcell_ole(data: bytes) -> dict[str, object]:
    with olefile.OleFileIO(io.BytesIO(data)) as compound:
        streams = ["/".join(parts) for parts in compound.listdir()]
        if not compound.exists("think-cellXML"):
            return {"readable": False, "streams": streams, "error": "Missing think-cellXML stream"}
        xml = compound.openstream("think-cellXML").read()
    try:
        root = etree.fromstring(xml)
    except etree.XMLSyntaxError as exc:
        return {"readable": False, "streams": streams, "xml_sha256": sha256(xml), "error": str(exc)}
    family_classes = Counter(
        etree.QName(node).localname
        for node in root
        if re.search(r"(?:ChartSE|GanttSE|TableSE)$", etree.QName(node).localname)
    )
    raw_names = []
    for node in root.iter():
        name = node.find("m_strName")
        if name is None:
            continue
        value = (name.text or name.get("val") or "").strip()
        if value:
            raw_names.append({"owner": etree.QName(node).localname, "id": node.get("id"), "name": value})
    names = []
    for value in dict.fromkeys(entry["name"] for entry in raw_names):
        candidates = [entry for entry in raw_names if entry["name"] == value]
        preferred = next(
            (
                entry
                for entry in candidates
                if re.search(r"(?:ChartSE|GanttSE|TableSE)$", str(entry["owner"]))
            ),
            candidates[0],
        )
        family_candidates = [
            entry
            for entry in candidates
            if re.search(r"(?:ChartSE|GanttSE|TableSE)$", str(entry["owner"]))
        ]
        logical_occurrences = len({(entry["owner"], entry["id"]) for entry in family_candidates or candidates})
        names.append(
            {
                **preferred,
                "logical_occurrences": logical_occurrences,
                "raw_occurrences": len(candidates),
            }
        )
    return {
        "readable": True,
        "streams": streams,
        "xml_sha256": sha256(xml),
        "model_families": dict(family_classes),
        "automation_names": names,
        "sequence_tables": sequence_tables(root),
        "scatter_tables": scatter_tables(root),
        "pie_tables": pie_tables(root),
    }


def chart_details(xml: bytes) -> dict[str, object]:
    root = etree.fromstring(xml)
    plot_area = root.find(".//c:plotArea", namespaces=NS)
    chart_types = []
    subtypes = []
    series = []
    if plot_area is not None:
        for chart in plot_area:
            local = etree.QName(chart).localname
            if not local.endswith("Chart"):
                continue
            chart_types.append(local)
            subtype: dict[str, object] = {"type": local}
            bar_direction = chart.find("c:barDir", namespaces=NS)
            grouping = chart.find("c:grouping", namespaces=NS)
            hole_size = chart.find("c:holeSize", namespaces=NS)
            first_slice_angle = chart.find("c:firstSliceAng", namespaces=NS)
            scatter_style = chart.find("c:scatterStyle", namespaces=NS)
            bubble_scale = chart.find("c:bubbleScale", namespaces=NS)
            size_represents = chart.find("c:sizeRepresents", namespaces=NS)
            show_negative_bubbles = chart.find("c:showNegBubbles", namespaces=NS)
            if bar_direction is not None:
                subtype["bar_direction"] = bar_direction.get("val")
            if grouping is not None:
                grouping_value = grouping.get("val")
                subtype["grouping"] = grouping_value
                subtype["percent_axis"] = grouping_value == "percentStacked"
            if hole_size is not None:
                subtype["hole_size"] = hole_size.get("val")
            if first_slice_angle is not None:
                subtype["first_slice_angle"] = first_slice_angle.get("val")
            if scatter_style is not None:
                subtype["scatter_style"] = scatter_style.get("val")
            if bubble_scale is not None:
                subtype["bubble_scale"] = float(bubble_scale.get("val"))
            if size_represents is not None:
                subtype["size_represents"] = size_represents.get("val")
            if show_negative_bubbles is not None:
                subtype["show_negative_bubbles"] = show_negative_bubbles.get("val") not in {"0", "false", "False"}
            if local == "scatterChart":
                scatter_series = chart.findall("c:ser", namespaces=NS)
                subtype["all_series_lines_hidden"] = all(
                    bool(node.xpath("./c:spPr/a:ln/a:noFill", namespaces=NS))
                    for node in scatter_series
                )
                subtype["all_series_smooth_false"] = all(
                    not node.xpath("./c:smooth[@val != '0' and @val != 'false' and @val != 'False']", namespaces=NS)
                    for node in scatter_series
                )
            subtype["explosions"] = [
                float(node.get("val"))
                for node in chart.xpath("./c:ser/c:explosion | ./c:ser/c:dPt/c:explosion", namespaces=NS)
                if node.get("val") not in (None, "")
            ]
            subtypes.append(subtype)
            for node in chart.findall("c:ser", namespaces=NS):
                names = node.xpath("./c:tx//c:strCache/c:pt/c:v | ./c:tx/c:v", namespaces=NS)
                values = node.xpath(
                    "./c:yVal/c:numRef/c:numCache/c:pt/c:v | ./c:yVal/c:numLit/c:pt/c:v | "
                    "./c:val/c:numRef/c:numCache/c:pt/c:v | ./c:val/c:numLit/c:pt/c:v",
                    namespaces=NS,
                )
                categories = node.xpath(
                    "./c:cat/c:strRef/c:strCache/c:pt/c:v | ./c:cat/c:strLit/c:pt/c:v | "
                    "./c:cat/c:numRef/c:numCache/c:pt/c:v | ./c:cat/c:numLit/c:pt/c:v",
                    namespaces=NS,
                )
                x_values = node.xpath(
                    "./c:xVal/c:numRef/c:numCache/c:pt/c:v | ./c:xVal/c:numLit/c:pt/c:v",
                    namespaces=NS,
                )
                bubble_sizes = node.xpath(
                    "./c:bubbleSize/c:numRef/c:numCache/c:pt/c:v | ./c:bubbleSize/c:numLit/c:pt/c:v",
                    namespaces=NS,
                )
                series.append(
                    {
                        "name": names[0].text if names else None,
                        "values": [float(value.text) for value in values if value.text not in (None, "")],
                        "x_values": [float(value.text) for value in x_values if value.text not in (None, "")],
                        "y_values": [float(value.text) for value in values if value.text not in (None, "")],
                        "bubble_sizes": [float(value.text) for value in bubble_sizes if value.text not in (None, "")],
                        "categories": [value.text for value in categories if value.text is not None],
                    }
                )
    return {"types": chart_types, "subtypes": subtypes, "series": series}


def slide_embedded_automation_names(root: etree._Element) -> list[dict[str, object]]:
    names = []
    for element in root.iter():
        for value in element.attrib.values():
            if not isinstance(value, str) or "thinkcell" not in value or "<m_strName>" not in value:
                continue
            start = value.find("<root")
            if start < 0:
                continue
            try:
                embedded = etree.fromstring(value[start:].encode("utf-8"))
            except etree.XMLSyntaxError:
                continue
            outer_local = etree.QName(element).localname
            owner = "AutomationTextField" if outer_local == "fld" else f"SlideEmbedded:{outer_local}"
            shape = element.xpath("ancestor::p:sp[1] | ancestor::p:graphicFrame[1] | ancestor::p:pic[1]", namespaces=NS)
            properties = shape[0].xpath(".//p:cNvPr[1]", namespaces=NS) if shape else []
            for name_node in embedded.findall(".//m_strName"):
                name = (name_node.text or name_node.get("val") or "").strip()
                if not name:
                    continue
                names.append(
                    {
                        "owner": owner,
                        "id": properties[0].get("id") if properties else None,
                        "shape_name": properties[0].get("name") if properties else None,
                        "name": name,
                        "displayed_text": "".join(element.xpath(".//a:t/text()", namespaces=NS)) if outer_local == "fld" else None,
                        "logical_occurrences": 1,
                        "raw_occurrences": 1,
                    }
                )
    return names


def numeric_lists_match(model: list[float | None], visible: list[float], tolerance: float = 1e-7) -> bool:
    compact = [value for value in model if value is not None and math.isfinite(value)]
    if len(compact) != len(visible):
        return False
    return all(math.isclose(left, right, rel_tol=tolerance, abs_tol=tolerance) for left, right in zip(compact, visible))


def normalized_percent_series(series_values: list[list[float | None]], category_extents=None) -> list[list[float | None]]:
    """Return category-wise 100% values while preserving empty model cells."""
    width = max((len(series) for series in series_values), default=0)
    totals = []
    for category_index in range(width):
        totals.append(
            sum(
                float(series[category_index])
                for series in series_values
                if category_index < len(series)
                and series[category_index] is not None
                and math.isfinite(float(series[category_index]))
            )
        )
    normalized = []
    if category_extents is not None:
        if len(category_extents) != width or any(not isinstance(v, (int, float)) or not math.isfinite(v) for v in category_extents):
            return []
        totals = category_extents
    for series in series_values:
        row: list[float | None] = []
        for category_index, value in enumerate(series):
            total = totals[category_index]
            row.append(None if value is None or total == 0 else float(value) * 100.0 / total)
        normalized.append(row)
    return normalized


def compare_sequence_model_to_chart(model: dict[str, object], chart: dict[str, object]) -> dict[str, object]:
    visible = chart["series"]
    model_values = model["series_values"]
    supported = len(model_values) == len(visible)
    if not supported:
        return {"supported": False, "values_match": False, "names_match": False, "comparison_mode": None}

    variants = [
        ("raw", model_values, model["series_names"]),
        ("raw_reversed", list(reversed(model_values)), list(reversed(model["series_names"]))),
    ]
    normalized = normalized_percent_series(model_values, model.get('category_extents'))
    variants.extend(
        [
            ("normalized_100", normalized, model["series_names"]),
            ("normalized_100_reversed", list(reversed(normalized)), list(reversed(model["series_names"]))),
        ]
    )
    for mode, candidate_values, candidate_names in variants:
        if len(candidate_values) != len(visible):
            continue
        values_match = all(
            numeric_lists_match(model_series, visible_series["values"])
            for model_series, visible_series in zip(candidate_values, visible)
        )
        names_match = all(
            model_name is None or visible_series["name"] is None or str(model_name) == str(visible_series["name"])
            for model_name, visible_series in zip(candidate_names, visible)
        )
        if values_match and names_match:
            return {"supported": True, "values_match": True, "names_match": True, "comparison_mode": mode}
    return {"supported": True, "values_match": False, "names_match": False, "comparison_mode": None}


def _scatter_point_bags_match(
    model_points: list[tuple[float, float, float | None]],
    visible_points: list[tuple[float, float, float | None]],
    tolerance: float = 1e-7,
) -> bool:
    """Compare point tuples without assuming native series or point ordering."""
    if len(model_points) != len(visible_points):
        return False
    remaining = list(visible_points)
    for expected in model_points:
        match_index = None
        for index, actual in enumerate(remaining):
            if all(
                (left is None and right is None)
                or (
                    left is not None
                    and right is not None
                    and math.isclose(float(left), float(right), rel_tol=tolerance, abs_tol=tolerance)
                )
                for left, right in zip(expected, actual)
            ):
                match_index = index
                break
        if match_index is None:
            return False
        remaining.pop(match_index)
    return not remaining


def compare_scatter_model_to_chart(model: dict[str, object], chart: dict[str, object]) -> dict[str, object]:
    chart_types = set(chart.get("types", []))
    supported_types = chart_types.intersection({"scatterChart", "bubbleChart"})
    if len(supported_types) != 1:
        return {"supported": False, "values_match": False, "names_match": False, "comparison_mode": None}
    chart_type = next(iter(supported_types))
    include_size = chart_type == "bubbleChart"

    model_points = []
    for point in model.get("points", []):
        x_value = point.get("x")
        y_value = point.get("y")
        size_value = point.get("size")
        if not isinstance(x_value, (int, float)) or not isinstance(y_value, (int, float)):
            return {"supported": False, "values_match": False, "names_match": False, "comparison_mode": None}
        if include_size and not isinstance(size_value, (int, float)):
            return {"supported": False, "values_match": False, "names_match": False, "comparison_mode": None}
        if not include_size and size_value is not None:
            return {"supported": False, "values_match": False, "names_match": False, "comparison_mode": None}
        model_points.append((float(x_value), float(y_value), float(size_value) if include_size else None))

    visible_points = []
    for series in chart.get("series", []):
        x_values = list(series.get("x_values", []))
        y_values = list(series.get("y_values", []))
        size_values = list(series.get("bubble_sizes", []))
        if len(x_values) != len(y_values) or (include_size and len(size_values) != len(x_values)):
            return {"supported": False, "values_match": False, "names_match": False, "comparison_mode": None}
        if not include_size and size_values:
            return {"supported": False, "values_match": False, "names_match": False, "comparison_mode": None}
        for index, (x_value, y_value) in enumerate(zip(x_values, y_values)):
            visible_points.append((float(x_value), float(y_value), float(size_values[index]) if include_size else None))

    values_match = _scatter_point_bags_match(model_points, visible_points)
    visible_names = [str(series["name"]) for series in chart.get("series", []) if series.get("name") is not None]
    model_groups = []
    for group in model.get("group_labels", []):
        if group is not None and str(group) not in model_groups:
            model_groups.append(str(group))
    names_match = not visible_names or visible_names == model_groups
    return {
        "supported": True,
        "values_match": values_match,
        "names_match": names_match,
        "comparison_mode": "xyz_unordered" if include_size else "xy_unordered",
    }


def compare_pie_model_to_chart(model: dict[str, object], chart: dict[str, object]) -> dict[str, object]:
    visible = chart["series"]
    chart_types = set(chart.get("types", []))
    supported = (
        model.get("series_count") == 1
        and len(visible) == 1
        and bool(chart_types.intersection({"pieChart", "doughnutChart"}))
    )
    if not supported:
        return {"supported": False, "values_match": False, "names_match": False, "comparison_mode": None}

    model_values = list(model.get("values", []))
    visible_values = visible[0]["values"]
    total = model.get("calculated_total")
    reported_total = model.get("total")
    totals_match = (
        isinstance(total, (int, float))
        and isinstance(reported_total, (int, float))
        and math.isfinite(float(total))
        and math.isfinite(float(reported_total))
        and math.isclose(float(total), float(reported_total), rel_tol=1e-9, abs_tol=1e-9)
    )
    if model.get('input_mode')=='percentage':
        totals_match = isinstance(total,(int,float)) and math.isclose(total,1.0,abs_tol=1e-9) and reported_total==100.0
    normalized = (
        [None if value is None else float(value) * 100.0 / float(total) for value in model_values]
        if isinstance(total, (int, float)) and math.isfinite(float(total)) and float(total) != 0
        else []
    )
    model_categories = [None if value is None else str(value) for value in model.get("categories", [])]
    visible_categories = [str(value) for value in visible[0].get("categories", [])]
    # think-cell's generated native pie cache may omit c:cat entirely.  Exact
    # category order remains assertable from the genuine ActiveDocument model;
    # when a native category cache is present it must match exactly.
    categories_match = not visible_categories or (
        len(model_categories) == len(visible_categories)
        and all(left is not None and left == right for left, right in zip(model_categories, visible_categories))
    )
    variants=(("normalized_100",normalized),) if model.get('input_mode')=='percentage' else (("raw",model_values),("normalized_100",normalized))
    for mode, candidate_values in variants:
        if totals_match and numeric_lists_match(candidate_values, visible_values) and categories_match:
            return {"supported": True, "values_match": True, "names_match": True, "comparison_mode": mode}
    return {"supported": True, "values_match": False, "names_match": categories_match, "comparison_mode": None}


def inspect_presentation(path: Path, strict_parity: bool) -> dict[str, object]:
    with zipfile.ZipFile(path) as package:
        bad_zip_entry = package.testzip()
        slides = logical_slides(package)
        presentation = etree.fromstring(package.read("ppt/presentation.xml"))
        slide_reports = []
        global_families: Counter[str] = Counter()
        global_names: list[dict[str, object]] = []
        parity_checks = []
        embeddings = embedding_inventory(package)
        embedding_role_counts = Counter(
            role for record in embeddings for role in record["roles"]
        )
        master_active_documents = []
        for record in embeddings:
            if "master_thinkcell_active_document" not in record["roles"]:
                continue
            data = package.read(str(record["part"]))
            master_active_documents.append(
                {
                    "part": record["part"],
                    "owners": record["owners"],
                    "model": inspect_thinkcell_ole(data),
                }
            )

        for logical_index, slide in enumerate(slides, 1):
            part = str(slide["part"])
            root = etree.fromstring(package.read(part))
            rels = relationship_map(package, part)
            for entry in slide_embedded_automation_names(root):
                global_names.append({"slide": logical_index, **entry})
            tagged_shapes = []
            for shape in root.xpath(".//*[@r:id]/ancestor-or-self::*[p:nvSpPr or p:nvGraphicFramePr or p:nvPicPr][1]", namespaces=NS):
                property_nodes = shape.xpath(".//p:cNvPr[1]", namespaces=NS)
                used_tags = []
                for node in shape.iter():
                    for key, value in node.attrib.items():
                        if key.startswith(R) and value in rels and rels[value]["type"].endswith("/tags"):
                            used_tags.extend(tag_values(package, rels[value]["resolved"]))
                if used_tags:
                    tagged_shapes.append(
                        {
                            "id": int(property_nodes[0].get("id")) if property_nodes else None,
                            "name": property_nodes[0].get("name") if property_nodes else None,
                            "tags": used_tags,
                        }
                    )

            active_docs = []
            seen_active_parts: set[str] = set()
            for ole in root.xpath(".//p:oleObj", namespaces=NS):
                rid = ole.get(R + "id")
                rel = rels.get(rid)
                if not rel or rel["resolved"] not in package.namelist():
                    continue
                prog_id = ole.get("progId", "")
                data = package.read(rel["resolved"])
                model = inspect_thinkcell_ole(data) if ("TCLayout.ActiveDocument" in prog_id or b"think-cellXML" in data) else None
                if model and rel["resolved"] not in seen_active_parts:
                    seen_active_parts.add(rel["resolved"])
                    active_docs.append({"part": rel["resolved"], "prog_id": prog_id, "model": model})
                    global_families.update(model.get("model_families", {}))
                    for entry in model.get("automation_names", []):
                        global_names.append({"slide": logical_index, **entry})

            charts = []
            seen_chart_parts: set[str] = set()
            for chart in root.xpath(".//c:chart", namespaces=NS):
                rid = chart.get(R + "id")
                rel = rels.get(rid)
                if rel and rel["resolved"] in package.namelist() and rel["resolved"] not in seen_chart_parts:
                    seen_chart_parts.add(rel["resolved"])
                    charts.append({"part": rel["resolved"], **chart_details(package.read(rel["resolved"]))})

            sequence = []
            for active in active_docs:
                sequence.extend(active["model"].get("sequence_tables", []))
            # Some think-cell families use sequence-table storage without a
            # native PowerPoint chart cache.  Only assert model/cache parity
            # when the slide actually contains a native chart to pair.
            if not charts:
                sequence = []
            remaining_charts = set(range(len(charts)))
            for model in sequence:
                candidates = []
                for chart_index in sorted(remaining_charts):
                    comparison = compare_sequence_model_to_chart(model, charts[chart_index])
                    candidates.append((chart_index, comparison))
                exact = [candidate for candidate in candidates if candidate[1]["values_match"] and candidate[1]["names_match"]]
                if exact:
                    chart_index, comparison = exact[0]
                elif candidates:
                    same_count = [candidate for candidate in candidates if candidate[1]["supported"]]
                    chart_index, comparison = (same_count or candidates)[0]
                else:
                    chart_index = None
                    comparison = {"supported": False, "values_match": False, "names_match": False, "comparison_mode": None}
                if chart_index is not None:
                    remaining_charts.remove(chart_index)
                    visible_count = len(charts[chart_index]["series"])
                    chart_part = charts[chart_index]["part"]
                else:
                    visible_count = 0
                    chart_part = None
                parity_checks.append(
                    {
                        "slide": logical_index,
                        "automation_name": model.get("automation_name") or None,
                        "chart_part": chart_part,
                        "supported": comparison["supported"],
                        "model_series": len(model["series_values"]),
                        "visible_series": visible_count,
                        "values_match": comparison["values_match"],
                        "names_match": comparison["names_match"],
                        "comparison_mode": comparison["comparison_mode"],
                    }
                )

            scatter = []
            for active in active_docs:
                scatter.extend(active["model"].get("scatter_tables", []))
            for model in scatter:
                candidates = []
                for chart_index in sorted(remaining_charts):
                    comparison = compare_scatter_model_to_chart(model, charts[chart_index])
                    candidates.append((chart_index, comparison))
                exact = [candidate for candidate in candidates if candidate[1]["values_match"] and candidate[1]["names_match"]]
                if exact:
                    chart_index, comparison = exact[0]
                elif candidates:
                    same_count = [candidate for candidate in candidates if candidate[1]["supported"]]
                    chart_index, comparison = (same_count or candidates)[0]
                else:
                    chart_index = None
                    comparison = {"supported": False, "values_match": False, "names_match": False, "comparison_mode": None}
                if chart_index is not None:
                    remaining_charts.remove(chart_index)
                    visible_series = len(charts[chart_index]["series"])
                    visible_points = sum(len(series.get("x_values", [])) for series in charts[chart_index]["series"])
                    chart_part = charts[chart_index]["part"]
                else:
                    visible_series = 0
                    visible_points = 0
                    chart_part = None
                model_groups = []
                for group in model.get("group_labels", []):
                    if group is not None and str(group) not in model_groups:
                        model_groups.append(str(group))
                parity_checks.append(
                    {
                        "kind": "scatter_chart",
                        "slide": logical_index,
                        "automation_name": model.get("automation_name") or None,
                        "chart_part": chart_part,
                        "supported": comparison["supported"],
                        "model_series": len(model_groups),
                        "visible_series": visible_series,
                        "model_points": int(model.get("point_count", 0)),
                        "visible_points": visible_points,
                        "values_match": comparison["values_match"],
                        "names_match": comparison["names_match"],
                        "comparison_mode": comparison["comparison_mode"],
                    }
                )

            pie = []
            for active in active_docs:
                pie.extend(active["model"].get("pie_tables", []))
            for model in pie:
                candidates = []
                for chart_index in sorted(remaining_charts):
                    comparison = compare_pie_model_to_chart(model, charts[chart_index])
                    candidates.append((chart_index, comparison))
                exact = [candidate for candidate in candidates if candidate[1]["values_match"] and candidate[1]["names_match"]]
                if exact:
                    chart_index, comparison = exact[0]
                elif candidates:
                    same_count = [candidate for candidate in candidates if candidate[1]["supported"]]
                    chart_index, comparison = (same_count or candidates)[0]
                else:
                    chart_index = None
                    comparison = {"supported": False, "values_match": False, "names_match": False, "comparison_mode": None}
                if chart_index is not None:
                    remaining_charts.remove(chart_index)
                    visible_count = len(charts[chart_index]["series"])
                    chart_part = charts[chart_index]["part"]
                else:
                    visible_count = 0
                    chart_part = None
                parity_checks.append(
                    {
                        "kind": "pie_chart",
                        "slide": logical_index,
                        "automation_name": model.get("automation_name") or None,
                        "chart_part": chart_part,
                        "supported": comparison["supported"],
                        "model_series": int(model.get("series_count", 0)),
                        "visible_series": visible_count,
                        "values_match": comparison["values_match"],
                        "names_match": comparison["names_match"],
                        "comparison_mode": comparison["comparison_mode"],
                    }
                )

            slide_reports.append(
                {
                    "slide": logical_index,
                    "slide_id": slide["id"],
                    "part": part,
                    "thinkcell_tagged_shapes": tagged_shapes,
                    "active_documents": active_docs,
                    "native_charts": charts,
                    "notes_substance_sha256": notes_substance(package, part),
                }
            )

        consolidated_names = []
        grouped_names: dict[tuple[int, str], list[dict[str, object]]] = {}
        for entry in global_names:
            grouped_names.setdefault((int(entry["slide"]), str(entry["name"]).casefold()), []).append(entry)
        for entries in grouped_names.values():
            family_entries = [
                entry
                for entry in entries
                if re.search(r"(?:ChartSE|GanttSE|TableSE)$", str(entry.get("owner", "")))
            ]
            preferred = (family_entries or entries)[0]
            logical_occurrences = (
                max(int(entry.get("logical_occurrences", 1)) for entry in family_entries)
                if family_entries
                else sum(int(entry.get("logical_occurrences", 1)) for entry in entries)
            )
            consolidated_names.append({**preferred, "logical_occurrences": logical_occurrences})

        active_doc_count = sum(len(slide["active_documents"]) for slide in slide_reports)
        tagged_shape_count = sum(len(slide["thinkcell_tagged_shapes"]) for slide in slide_reports)
        chart_count = sum(len(slide["native_charts"]) for slide in slide_reports)
        unreadable = [
            {"slide": slide["slide"], "part": active["part"], "error": active["model"].get("error")}
            for slide in slide_reports
            for active in slide["active_documents"]
            if not active["model"].get("readable")
        ]
        active_without_visible_tags = [
            slide["slide"] for slide in slide_reports if slide["active_documents"] and not slide["thinkcell_tagged_shapes"]
        ]
        duplicate_names = []
        counts: Counter[tuple[int, str]] = Counter()
        for entry in consolidated_names:
            counts[(int(entry["slide"]), str(entry["name"]).casefold())] += int(
                entry.get("logical_occurrences", 1)
            )
        for key, count in counts.items():
            if count > 1:
                duplicate_names.append({"slide": key[0], "name": key[1], "count": count})

        parity_mismatches = [
            check
            for check in parity_checks
            if (not check["supported"]) or (not check["values_match"]) or (not check["names_match"])
        ]
        strict_checks = [check for check in parity_checks if check["supported"]]
        assertions = {
            "valid_zip": bad_zip_entry is None,
            "no_dangling_relationships": not dangling_relationships(package),
            "no_duplicate_relationship_ids": not duplicate_relationship_ids(package),
            "no_missing_content_types": not missing_content_types(package),
            "valid_mc_prefixes": not invalid_mc_prefixes(package),
            "embedded_parts_have_unique_owners": not shared_embedding_owners(package),
            "active_documents_readable": not unreadable,
            "master_active_documents_readable": all(
                document["model"].get("readable") for document in master_active_documents
            ),
            "active_documents_have_tagged_shapes": not active_without_visible_tags,
            # A single sequence-table model paired with a single visible chart is
            # an exact, auditable pair.  Divergence here is a broken hybrid even
            # when --strict-parity was not requested, so fail closed.
            "no_model_visible_chart_mismatch": not parity_mismatches,
            "strict_parity_available": (not strict_parity)
            or (bool(parity_checks) and len(strict_checks) == len(parity_checks)),
            "strict_parity_pass": (not strict_parity)
            or (
                bool(parity_checks)
                and all(
                    check["supported"] and check["values_match"] and check["names_match"]
                    for check in parity_checks
                )
            ),
        }
        return {
            "path": str(path),
            "sha256": sha256(path.read_bytes()),
            "slides": len(slides),
            "masters": len([name for name in package.namelist() if re.fullmatch(r"ppt/slideMasters/slideMaster\d+\.xml", name)]),
            "layouts": len([name for name in package.namelist() if re.fullmatch(r"ppt/slideLayouts/slideLayout\d+\.xml", name)]),
            "embeddings": len(embeddings),
            "embedding_inventory": embeddings,
            "embedding_role_counts": dict(embedding_role_counts),
            "custom_xml": len([name for name in package.namelist() if name.startswith("customXml/") and name.endswith(".xml")]),
            "active_documents": active_doc_count,
            "master_active_documents": len(master_active_documents),
            "master_active_document_details": master_active_documents,
            "thinkcell_tagged_shapes": tagged_shape_count,
            "native_charts": chart_count,
            "model_families": dict(global_families),
            "automation_names": consolidated_names,
            "duplicate_names_within_slide": duplicate_names,
            "parity_checks": parity_checks,
            "parity_mismatches": parity_mismatches,
            "unreadable_active_documents": unreadable,
            "active_documents_without_tagged_shapes": active_without_visible_tags,
            "slide_details": slide_reports,
            "dangling_relationships": dangling_relationships(package),
            "duplicate_relationship_ids": duplicate_relationship_ids(package),
            "shared_embedding_owners": shared_embedding_owners(package),
            "missing_content_types": missing_content_types(package),
            "invalid_mc_prefixes": invalid_mc_prefixes(package),
            "assertions": assertions,
        }


def baseline_assertions(
    current: dict[str, object],
    baseline: dict[str, object],
    allow_thinkcell_child_count_change: bool,
    allow_master_active_document_addition: bool,
) -> dict[str, bool]:
    current_notes = [slide["notes_substance_sha256"] for slide in current["slide_details"]]
    baseline_notes = [slide["notes_substance_sha256"] for slide in baseline["slide_details"]]

    def parity_fingerprints(presentation: dict[str, object]) -> list[tuple[object, ...]]:
        return sorted(
            (
                mismatch["slide"],
                mismatch.get("automation_name") or "",
                bool(mismatch["supported"]),
                int(mismatch["model_series"]),
                int(mismatch["visible_series"]),
                bool(mismatch["values_match"]),
                bool(mismatch["names_match"]),
            )
            for mismatch in presentation["parity_mismatches"]
        )

    current_roles = Counter(current.get("embedding_role_counts", {}))
    baseline_roles = Counter(baseline.get("embedding_role_counts", {}))
    expected_roles_after_master_addition = baseline_roles.copy()
    expected_roles_after_master_addition["master_thinkcell_active_document"] += 1
    master_details = current.get("master_active_document_details", [])
    allowed_master_addition = bool(
        allow_master_active_document_addition
        and baseline.get("master_active_documents", 0) == 0
        and current.get("master_active_documents", 0) == 1
        and current["embeddings"] == baseline["embeddings"] + 1
        and current_roles == expected_roles_after_master_addition
        and len(master_details) == 1
        and master_details[0]["model"].get("readable")
        and not master_details[0]["model"].get("model_families")
        and not master_details[0]["model"].get("automation_names")
    )
    embedding_roles_preserved = current_roles == baseline_roles or allowed_master_addition
    master_active_document_count_preserved = (
        current.get("master_active_documents", 0)
        == baseline.get("master_active_documents", 0)
        or allowed_master_addition
    )

    return {
        "slide_count_preserved": current["slides"] == baseline["slides"],
        "master_count_preserved": current["masters"] == baseline["masters"],
        "layout_count_preserved": current["layouts"] == baseline["layouts"],
        "embedding_count_preserved": (
            current["embeddings"] == baseline["embeddings"] or allowed_master_addition
        ),
        "embedding_roles_preserved": embedding_roles_preserved,
        "custom_xml_not_lost": current["custom_xml"] >= baseline["custom_xml"],
        "active_document_count_preserved": current["active_documents"] == baseline["active_documents"],
        "master_active_document_count_preserved": master_active_document_count_preserved,
        "thinkcell_tag_ownership_preserved": (
            all(
                (not slide["active_documents"]) or bool(slide["thinkcell_tagged_shapes"])
                for slide in current["slide_details"]
            )
            if allow_thinkcell_child_count_change
            else current["thinkcell_tagged_shapes"] == baseline["thinkcell_tagged_shapes"]
        ),
        "native_chart_count_preserved": current["native_charts"] == baseline["native_charts"],
        "model_families_preserved": current["model_families"] == baseline["model_families"],
        "model_visible_chart_parity_preserved": parity_fingerprints(current)
        == parity_fingerprints(baseline),
        "speaker_notes_substance_preserved": current_notes == baseline_notes,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit native think-cell package integrity and model/chart parity.")
    parser.add_argument("presentation", type=Path)
    parser.add_argument("--baseline", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--strict-parity", action="store_true")
    parser.add_argument("--allow-thinkcell-child-count-change", action="store_true")
    parser.add_argument("--allow-master-active-document-addition", action="store_true")
    args = parser.parse_args()

    presentation = args.presentation.resolve()
    current = inspect_presentation(presentation, args.strict_parity)
    report: dict[str, object] = {"presentation": current}
    checks = dict(current["assertions"])
    if args.baseline:
        baseline = inspect_presentation(args.baseline.resolve(), False)
        comparison = baseline_assertions(
            current,
            baseline,
            args.allow_thinkcell_child_count_change,
            args.allow_master_active_document_addition,
        )
        report["baseline"] = baseline
        report["baseline_comparison"] = comparison
        # In preservation mode, inherited source mismatches are evidence, not a
        # regression. Standalone and --strict-parity audits still require an
        # absolutely clean current package.
        if not args.strict_parity:
            checks.pop("no_model_visible_chart_mismatch", None)
        checks.update(comparison)
    report["checks"] = checks
    report["pass"] = all(checks.values())

    payload = json.dumps(report, indent=2)
    if args.output:
        output = args.output.resolve()
        reserved = {presentation}
        if args.baseline:
            reserved.add(args.baseline.resolve())
        if output in reserved:
            raise SystemExit("Audit output must be distinct from the presentation and baseline inputs.")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(payload, encoding="utf-8")
    print(json.dumps({"pass": report["pass"], "checks": checks, "presentation_sha256": current["sha256"]}, indent=2))
    raise SystemExit(0 if report["pass"] else 2)


if __name__ == "__main__":
    main()
