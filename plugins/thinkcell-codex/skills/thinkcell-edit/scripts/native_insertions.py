"""Offline experimental native feature grafts onto explicitly selected clean charts.

The prepared package is not a verified native result. Official regeneration,
saved native readback, changed-data verification and visual review are required.
No proprietary fixture is embedded or downloaded by this adapter.
"""
from __future__ import annotations

import argparse
import base64
import copy
import json
import posixpath
import subprocess
import sys
import tempfile
import uuid
import zipfile
from pathlib import Path

from lxml import etree as E

HERE = Path(__file__).resolve().parent
IMPL = HERE / "thinkcell_no_click/implementation"
sys.path.insert(0, str(IMPL))
import prepare_thinkcell_name as naming
from runtime import powershell, powershell_env
from trendline_math import polynomial_fit, verify_polynomial_forecast, verify_lower_forecast, forecast_close

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
# Authenticated saved native regeneration profiles, never inferred from menu order.
TRENDLINE_PROFILES = {"linear-trendline": ("0", "linear"),
                      "power-trendline": ("1", "power"),
                      "exponential-trendline": ("2", "exp"),
                      "logarithmic-trendline": ("3", "log"),
                      "quadratic-trendline": ("6", "poly"),
                      "cubic-trendline": ("7", "poly"),
                      "quartic-trendline": ("8", "poly")}
POLYNOMIAL_ORDERS = {"quadratic-trendline": 2, "cubic-trendline": 3, "quartic-trendline": 4}
INSERTION_FEATURES = {"legend", "errorbar-range", "errorbar-caps", *TRENDLINE_PROFILES}



def series_refs(chart):
    return [x.get("idref") for x in chart["table"].findall("./m_cscdser/elem")]


def bound_series(chart):
    ids = chart["doc"]["ids"]
    return [{"id": sid, "source_id": ids[sid].find("m_varsrc").get("idref")}
            for sid in series_refs(chart)]


def dependency_closure(ids, seeds, boundary):
    """Resolve every nonzero reference, without cloning chart data or owners."""
    seen, pending = set(), list(seeds)
    while pending:
        ident = pending.pop()
        if ident == "0" or ident in boundary or ident in seen:
            continue
        naming.need(ident in ids, "Unresolved donor feature reference: " + str(ident))
        seen.add(ident)
        pending.extend(n.get("idref") for n in ids[ident].iter() if n.get("idref"))
    return seen


def clone_graph(root, donor_ids, closure, binding_map):
    mapping = dict(binding_map)
    start = max(int(x.get("id")) for x in root if x.get("id")) + 1
    mapping.update({ident: str(start + i) for i, ident in enumerate(sorted(closure, key=int))})
    tags, clones = {}, []
    for ident in sorted(closure, key=int):
        node = copy.deepcopy(donor_ids[ident])
        node.set("id", mapping[ident])
        for child in node.iter():
            ref = child.get("idref")
            if ref and ref != "0":
                naming.need(ref in mapping, "Unbound graft reference")
                child.set("idref", mapping[ref])
            if child.tag == "m_bstrShapeName" and child.text:
                tags.setdefault(child.text, "t" + base64.urlsafe_b64encode(uuid.uuid4().bytes).decode().rstrip("="))
                child.text = tags[child.text]
            if child.tag == "m_guid" and child.get("val"):
                child.set("val", str(uuid.uuid4()))
        root.append(node)
        clones.append(node)
    return mapping, tags, clones


def _select(path, selector, expected_sha, family="CSequenceChartSE"):
    raw = Path(path).read_bytes()
    naming.need(naming.sha(raw) == expected_sha.upper(), "Source hash changed; inspect again")
    _, charts, _ = naming.inventory(raw)
    selected = naming.choose(charts, **selector)
    # Never detach linked charts or import unknown link state.
    naming.link_contract(selected)
    sink = selected["table"].find("m_advisesink")
    naming.need(sink is not None and sink.get("idref") == "0", "Linked chart insertion is outside this route")
    naming.need(selected["owner"].tag == family, "Insertion requires " + family)
    naming.need(selected["owner"].find("m_bConsistent").get("val") == "1", "Chart is inconsistent")
    return raw, selected


def _physical_import(source_zip, donor_zip, selected, donor, tagmap, field_formats=()):
    ns = naming.NS
    slides = naming.logical_slides(source_zip)
    slide = slides[selected["doc"]["slide_number"] - 1]["part"]
    donor_slide = naming.logical_slides(donor_zip)[donor["doc"]["slide_number"] - 1]["part"]
    relpart = posixpath.dirname(slide) + "/_rels/" + posixpath.basename(slide) + ".rels"
    root = naming.xml(source_zip.read(slide))
    relroot = naming.xml(source_zip.read(relpart))
    donorroot = naming.xml(donor_zip.read(donor_slide))
    donor_rels = naming.relationship_map(donor_zip, donor_slide)
    tree = root.find("p:cSld/p:spTree", ns)
    naming.need(tree is not None, "Missing physical shape tree")
    next_id = max(int(x.get("id")) for x in root.findall(".//p:cNvPr", ns)) + 1
    content_types = naming.xml(source_zip.read("[Content_Types].xml"))
    additions, copied_tags = {}, set()
    for original in donorroot.findall(".//p:sp", ns) + donorroot.findall(".//p:cxnSp", ns):
        imports = []
        for tagref in original.findall(".//p:tags", ns):
            relation = donor_rels[tagref.get("{" + ns["r"] + "}id")]
            tagroot = naming.xml(donor_zip.read(relation["resolved"]))
            matched = {x.get("val") for x in tagroot if x.get("name", "").upper() == "THINKCELLSHAPEDONOTDELETE" and x.get("val") in tagmap}
            if matched:
                imports.append((tagref, relation, tagroot, matched))
        if not imports:
            continue
        allowed = {tagref.get("{" + ns["r"] + "}id") for tagref, _, _, _ in imports}
        dependency_ids = {value for node in original.iter() for key, value in node.attrib.items()
                          if key.startswith("{" + ns["r"] + "}")}
        naming.need(dependency_ids <= allowed, "Legend shape has unsupported physical dependencies")
        shape = copy.deepcopy(original)
        for field in shape.findall(".//a:fld", ns):
            naming.need(field.get("type") in {"datetime" + key for key in field_formats}, "Foreign field has no cloned native variable owner")
            if isinstance(field_formats, dict):
                fmt, text = field_formats[field.get("type")[8:]]
                field.set("type", "datetime" + fmt)
                display = field.find("a:t", ns)
                naming.need(display is not None, "Legend field has no visible text")
                display.text = text
            field.set("id", "{" + str(uuid.uuid4()).upper() + "}")
        cn = shape.find(".//p:cNvPr", ns)
        cn.set("id", str(next_id)); cn.set("name", "Native inserted legend " + str(next_id)); next_id += 1
        for original_tag, relation, tagroot, matched in imports:
            naming.need(not (copied_tags & matched), "Duplicate physical native tag")
            copied_tags.update(matched)
            for tag in tagroot:
                if tag.get("val") in tagmap:
                    tag.set("val", tagmap[tag.get("val")])
            newpart = "ppt/tags/native-insertion-" + uuid.uuid4().hex + ".xml"
            newrid = "rIdNativeInsertion" + uuid.uuid4().hex
            E.SubElement(relroot, "{" + REL_NS + "}Relationship", Id=newrid, Type=relation["type"], Target=posixpath.relpath(newpart, posixpath.dirname(slide)))
            for tagref in shape.findall(".//p:tags", ns):
                if tagref.get("{" + ns["r"] + "}id") == original_tag.get("{" + ns["r"] + "}id"):
                    tagref.set("{" + ns["r"] + "}id", newrid)
            additions[newpart] = E.tostring(tagroot)
            E.SubElement(content_types, "{" + CT_NS + "}Override", PartName="/" + newpart, ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml")
        tree.append(shape)
    naming.need(copied_tags, "No physical native legend shapes were imported")
    return {slide: E.tostring(root), relpart: E.tostring(relroot), "[Content_Types].xml": E.tostring(content_types)}, additions, {tagmap[t] for t in copied_tags}


def prepare(source, output, donor_path, feature, selector, donor_selector,
            source_sha256, donor_sha256, series_pair=None, target_series_id=None, donor_feature_id=None):
    source, output, donor_path = map(lambda p: Path(p).resolve(), (source, output, donor_path))
    naming.need(len({str(p).casefold() for p in (source, output, donor_path)}) == 3, "Source, donor and output must be distinct")
    naming.need(not output.exists(), "Refusing to overwrite output")
    naming.need(feature in INSERTION_FEATURES, "Unknown insertion feature")
    family = "CScatterChartSE" if feature.endswith("trendline") else "CSequenceChartSE"
    raw, target = _select(source, selector, source_sha256, family)
    donor_raw, donor = _select(donor_path, donor_selector, donor_sha256, family)
    root, ids = target["doc"]["root"], target["doc"]["ids"]
    original_nodes = {ident: E.tostring(node) for ident, node in ids.items()}
    allowed_changes = set()
    donor_ids = donor["doc"]["ids"]
    binding = {donor["owner"].get("id"): target["owner"].get("id"), donor["table"].get("id"): target["table"].get("id")}
    physical_changes, additions, physical_tags = {}, {}, set()
    if feature == "legend":
        legends = [n for n in root.findall("CSequenceChartLegendSE") if n.find("m_cse").get("idref") == target["owner"].get("id")]
        naming.need(not legends, "Target already owns a legend")
        seeds = [n for n in donor["doc"]["root"].findall("CSequenceChartLegendSE") if n.find("m_cse").get("idref") == donor["owner"].get("id")]
        naming.need(len(seeds) == 1, "Donor legend is missing or ambiguous")
        source_series, donor_series = bound_series(target), bound_series(donor)
        naming.need(len(source_series) <= len(donor_series), "Legend donor has fewer series than target")
        donor_series = donor_series[:len(source_series)]
        for ds, ts in zip(donor_series, source_series):
            allowed_changes.add(ts["source_id"])
            binding[ds["id"]] = ts["id"]
            binding[ds["source_id"]] = ts["source_id"]
        # Import only entries corresponding to explicitly paired donor series.
        donor_ids = dict(donor_ids)
        seed = copy.deepcopy(seeds[0])
        entries = seed.find("m_clegendentry")
        retained = {x["id"] for x in donor_series}
        for entry in list(entries):
            parent = donor_ids[entry.get("idref")].find("m_featParent")
            naming.need(parent is not None, "Legend entry has no semantic series owner")
            if parent.get("idref") not in retained:
                entries.remove(entry)
        naming.need(len(entries) == len(source_series), "Legend entries do not cover target series uniquely")
        entries.set("length", str(len(entries)))
        donor_ids[seed.get("id")] = seed
        textvars = [e.get("idref") for ds in donor_series for e in donor_ids[ds["source_id"]].findall("./m_ctextvar/elem")]
        naming.need(textvars, "Donor legend has no native text variable bindings")
        closure = dependency_closure(donor_ids, [seed.get("id"), *textvars], binding)
        mapping, tagmap, clones = clone_graph(root, donor_ids, closure, binding)
        field_formats = {}
        cloned_ids = {node.get("id"): node for node in clones}
        for ds, ts in zip(donor_series, source_series):
            target_variables = ids[ts["source_id"]].find("m_ctextvar")
            naming.need(target_variables is not None, "Target series has no text-variable collection")
            for ref in donor_ids[ds["source_id"]].findall("./m_ctextvar/elem"):
                E.SubElement(target_variables, "elem", idref=mapping[ref.get("idref")])
                name = ids[ts["source_id"]].findtext("m_varval")
                fmt = "'" + "''".join(name.replace("'", "''")) + "'"
                old_format = donor_ids[ref.get("idref")].findtext("m_bstrFormat")
                naming.need(old_format not in field_formats, "Donor legend field format binds multiple sources")
                field_formats[old_format] = (fmt, name)
                cloned_ids[mapping[ref.get("idref")]].find("m_bstrFormat").text = fmt
        box = root.find("CContainerSE/m_cse")
        allowed_changes.add(box.getparent().get("id"))
        chart_slot = next(i for i, x in enumerate(box) if x.get("idref") == target["owner"].get("id"))
        box.insert(chart_slot + 1, E.Element("elem", idref=mapping[seed.get("id")]))
        box.set("length", str(len(box)))
        with zipfile.ZipFile(source) as z, zipfile.ZipFile(donor_path) as dz:
            physical_changes, additions, physical_tags = _physical_import(z, dz, target, donor, tagmap, field_formats)
    elif feature in {"errorbar-range", "errorbar-caps"}:
        allowed_changes.add(target["table"].get("id"))
        naming.need(target["owner"].find("m_estDefault").get("val") == "1", "Error-bar insertion requires an existing line chart")
        box = target["table"].find("m_cscserrange")
        naming.need(box is not None and not len(box), "Target already owns ranges or has no range collection")
        seeds = [donor_ids[e.get("idref")] for e in donor["table"].findall("./m_cscserrange/elem")]
        if donor_feature_id is not None:
            seeds = [s for s in seeds if s.get("id") == donor_feature_id]
        naming.need(len(seeds) == 1, "Donor range is missing or ambiguous")
        source_series = series_refs(target)
        pair = series_pair if series_pair is not None else source_series[:2]
        naming.need(len(pair) == 2 and len(set(pair)) == 2 and set(pair) <= set(source_series), "Range pair must bind two distinct target series")
        donor_pair = [x.get("idref") for x in seeds[0].findall("./m_cscdseries/elem")]
        naming.need(len(donor_pair) == 2, "Donor does not have a two-series range")
        binding.update(dict(zip(donor_pair, pair)))
        closure = dependency_closure(donor_ids, [seeds[0].get("id")], binding)
        naming.need(closure == {seeds[0].get("id")}, "Range donor has unsupported dependent objects")
        mapping, _, clones = clone_graph(root, donor_ids, closure, binding)
        E.SubElement(box, "elem", idref=mapping[seeds[0].get("id")]); box.set("length", "1")
        if feature == "errorbar-caps":
            donor_series = series_refs(donor)
            first_vector = donor_ids[donor["table"].find("./ocol/elem").get("idref")]
            donor_scalars = first_vector.findall("./ocol/elem")
            styles = []
            for sid in donor_pair:
                scalar = donor_ids[donor_scalars[donor_series.index(sid)].get("idref")]
                props = scalar.find("m_markerprops")
                naming.need(props is not None and props.find("m_emarkerstyle").get("val") == "dash", "Donor range endpoint lacks native dash cap")
                styles.append(props)
            for vector_ref in target["table"].findall("./ocol/elem"):
                vector = ids[vector_ref.get("idref")]
                scalars = vector.findall("./ocol/elem")
                for sid, props in zip(pair, styles):
                    scalar = ids[scalars[source_series.index(sid)].get("idref")]
                    allowed_changes.add(scalar.get("id"))
                    existing = scalar.find("m_markerprops")
                    naming.need(existing is not None, "Target scalar lacks native marker property slot")
                    scalar.replace(existing, copy.deepcopy(props))
                    scalar.find("m_bAlwaysShown").set("val", "1")
    else:
        allowed_changes.add(target["table"].get("id"))
        import thinkcell
        naming.need(target_series_id is not None and donor_feature_id is not None, "Explicit target series and donor partition required")
        series = [e.get("idref") for e in target["table"].findall("./m_cscatdseries/elem")]
        naming.need(target_series_id in series, "Trendline series does not belong to target")
        seed = donor_ids.get(donor_feature_id)
        naming.need(seed is not None and seed.tag == "CScatterChartPartition", "Donor partition is unresolved")
        type_code, physical_type = TRENDLINE_PROFILES[feature]
        naming.need(seed.find("m_epartitiontype").get("val") == "1" and seed.find("m_etrendlinetype").get("val") == type_code, "Donor native trendline type differs")
        donor_sid = seed.find("m_scatdseries").get("idref")
        naming.need(not any(ids[e.get("idref")].find("m_scatdseries").get("idref") == target_series_id for e in target["table"].findall("./m_cscatpartition/elem")), "Target series already owns a partition")
        binding[donor_sid] = target_series_id
        closure = dependency_closure(donor_ids, [donor_feature_id], binding)
        mapping, _, clones = clone_graph(root, donor_ids, closure, binding)
        box = target["table"].find("m_cscatpartition")
        naming.need(box is not None, "Missing native scatter partition collection")
        E.SubElement(box, "elem", idref=mapping[donor_feature_id]); box.set("length", str(len(box)))
        target_data = thinkcell.baseline_request(source, naming.choose(naming.inventory(raw)[1], **selector))["expected_model"]
        donor_data = thinkcell.baseline_request(donor_path, donor)["expected_model"]
        with zipfile.ZipFile(source) as z, zipfile.ZipFile(donor_path) as dz:
            target_part = target["frames"][0]["native_chart_part"]
            target_cache = naming.xml(z.read(target_part))
            donor_cache = naming.xml(dz.read(donor["frames"][0]["native_chart_part"]))
            physical_target = scatter_carrier(target_cache, target, target_series_id, target_data)
            if feature != "linear-trendline":
                group = ids[ids[target_series_id].find("m_varsrc").get("idref")].findtext("m_varval")
                if feature == "power-trendline":
                    power_fit(target_data, group)
                elif feature in POLYNOMIAL_ORDERS:
                    polynomial_fit(target_data, group, POLYNOMIAL_ORDERS[feature])
                else:
                    nonlinear_fit(target_data, group, physical_type)
            physical_donor = scatter_carrier(donor_cache, donor, donor_sid, donor_data)
            trends = physical_donor.xpath('./*[local-name()="trendline"]')
            naming.need(len(trends) == 1 and trends[0].xpath('./*[local-name()="trendlineType"]/@val') == [physical_type], "Donor has no matching unique physical native trendline")
            orders = trends[0].xpath('./*[local-name()="order"]/@val')
            naming.need(orders == ([str(POLYNOMIAL_ORDERS[feature])] if feature in POLYNOMIAL_ORDERS else []),
                        "Donor polynomial order differs from authenticated native profile")
            naming.need(not trends[0].xpath('./*[local-name()="intercept"]'), "Donor forced intercept is outside calculated profile")
            naming.need(not physical_target.xpath('./*[local-name()="trendline"]'), "Target already has a physical trendline")
            trend = copy.deepcopy(trends[0])
            for extension in list(trend):
                if E.QName(extension).localname in {"forward", "backward"}:
                    trend.remove(extension)
            before = next(x for x in physical_target if E.QName(x).localname in {"xVal", "yVal"})
            physical_target.insert(physical_target.index(before), trend)
            physical_changes[target_part] = E.tostring(target_cache)
    naming.need(all(E.tostring(ids[ident]) == payload for ident, payload in original_nodes.items()
                    if ident not in allowed_changes), "Unrelated native model objects changed")
    # Series sources and scalars must remain private to the selected chart.
    _, siblings, _ = naming.inventory(raw)
    for sibling in siblings:
        if sibling["doc"]["part"] != target["doc"]["part"] or sibling["owner"].get("id") == target["owner"].get("id"):
            continue
        refs = {x.get("idref") for x in sibling["table"].iter() if x.get("idref") not in {None, "0"}}
        pending = list(refs)
        while pending:
            ident = pending.pop()
            naming.need(ident not in allowed_changes, "Insertion touches a shared sibling native object")
            node = sibling["doc"]["ids"].get(ident)
            if node is None:
                continue
            for x in node.iter():
                ref = x.get("idref")
                if ref not in {None, "0"} and ref not in refs:
                    refs.add(ref); pending.append(ref)
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="native_insertion_", dir=output.parent) as td:
        td = Path(td)
        carrier, model = td / "carrier.bin", td / "model.xml"
        carrier.write_bytes(target["doc"]["ole"]); model.write_bytes(E.tostring(root))
        run = subprocess.run([powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(IMPL / "replace_ole_stream.ps1"), "-StoragePath", str(carrier), "-StreamBytesPath", str(model)], capture_output=True, text=True, env=powershell_env(), timeout=60)
        naming.need(run.returncode == 0, "Offline model stream replacement failed: " + run.stderr[:1000])
        physical_changes[target["doc"]["part"]] = carrier.read_bytes()
        with zipfile.ZipFile(source) as z, zipfile.ZipFile(output, "w") as w:
            w.comment = z.comment
            if "[Content_Types].xml" in physical_changes:
                types = physical_changes["[Content_Types].xml"]
            else:
                types = z.read("[Content_Types].xml")
            old = b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
            if old in types:
                physical_changes["[Content_Types].xml"] = types.replace(old, b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
            for entry in z.infolist():
                w.writestr(entry, physical_changes.get(entry.filename, z.read(entry.filename)))
            for part, payload in additions.items():
                w.writestr(part, payload)
    naming.need(naming.sha(source.read_bytes()) == source_sha256.upper(), "Source changed during insertion")
    naming.need(naming.sha(donor_path.read_bytes()) == donor_sha256.upper(), "Donor changed during insertion")
    with zipfile.ZipFile(source) as z, zipfile.ZipFile(output) as w:
        untouched = [part for part in z.namelist() if part not in physical_changes]
        naming.need(all(z.read(part) == w.read(part) for part in untouched), "Unrelated package parts changed")
    naming.inventory(output.read_bytes())
    return {"status": "PREPARED_UNVERIFIED_NATIVE_INSERTION", "feature": feature,
            "source_sha256": source_sha256.upper(), "donor_sha256": donor_sha256.upper(),
            "output": str(output), "output_sha256": naming.sha(output.read_bytes()),
            "selector": selector, "inserted_model_ids": [x.get("id") for x in clones],
            "imported_physical_native_tags": sorted(physical_tags), "untouched_package_parts": len(untouched),
            "required_gates": ["official regeneration", "saved native consistency", "exact model and embedded data", "native physical semantics", "second changed-data update", "source and sibling preservation", "visual review"]}


def scatter_carrier(cache, chart, series_id, expected):
    ids = chart["doc"]["ids"]
    source = ids[ids[series_id].find("m_varsrc").get("idref")]
    group = source.findtext("m_varval")
    wanted = [(float(x), float(y)) for g, x, y in zip(expected["group_labels"], expected["x_values"], expected["y_values"]) if g == group]
    naming.need(len(wanted) >= 2 and len({x for x, _ in wanted}) >= 2, "Trendline needs at least two points with distinct X values")
    matches = []
    for series in cache.xpath('.//*[local-name()="scatterChart"]/*[local-name()="ser"]'):
        columns = []
        for kind in ("xVal", "yVal"):
            points = series.xpath('./*[local-name()="' + kind + '"]/*[local-name()="numRef"]/*[local-name()="numCache"]/*[local-name()="pt"]')
            indexed = {int(pt.get("idx")): float(pt.xpath('string(./*[local-name()="v"])')) for pt in points}
            naming.need(len(indexed) == len(points), "Physical scatter point indices are duplicated")
            columns.append(indexed)
        naming.need(set(columns[0]) == set(columns[1]), "Physical scatter X/Y index closure differs")
        actual = [(columns[0][i], columns[1][i]) for i in sorted(columns[0])]
        if sorted(actual) == sorted(wanted):
            matches.append(series)
    naming.need(len(matches) == 1, "Physical native scatter series binding is missing or ambiguous")
    return matches[0]


def indexed_numeric_values(node, field):
    """Read a full native numeric cache, rejecting stale or ambiguous indices."""
    import math
    caches = node.xpath('./*[local-name()="' + field + '"]/*[local-name()="numRef"]/*[local-name()="numCache"]')
    naming.need(len(caches) == 1, "Numeric cache is missing or ambiguous: " + field)
    cache = caches[0]
    counts = cache.xpath('./*[local-name()="ptCount"]/@val')
    naming.need(len(counts) == 1 and counts[0].isdigit(), "Numeric cache point count is invalid")
    count, values = int(counts[0]), {}
    naming.need(count > 0, "Numeric cache is empty")
    for point in cache.xpath('./*[local-name()="pt"]'):
        index = point.get("idx", "")
        raw = point.xpath('./*[local-name()="v"]/text()')
        naming.need(index.isdigit() and int(index) not in values and len(raw) == 1, "Numeric cache indices or values are ambiguous")
        value = float(raw[0])
        naming.need(math.isfinite(value), "Numeric cache contains nonfinite data")
        values[int(index)] = value
    naming.need(set(values) == set(range(count)), "Numeric cache indices do not cover its point count")
    return [values[i] for i in range(count)]


def physical_endpoint_values(series, direction):
    field = direction + "Val" if series.xpath('./*[local-name()="' + direction + 'Val"]') else "val"
    return indexed_numeric_values(series, field)


def verify_category_positions(series, direction, categories):
    """Bind every interval endpoint to its native category, including scatter positions."""
    if E.QName(series.getparent()).localname == "scatterChart":
        orthogonal = "xVal" if direction == "y" else "yVal"
        actual = indexed_numeric_values(series, orthogonal)
        naming.need(actual == [index + .5 for index in range(len(categories))],
                    "Error-bar orthogonal category positions differ from the bounded native profile")
        return
    if not series.xpath('./*[local-name()="cat"]'):
        # Native sequence caches omit labels and use implicit category ordinals.
        # A complete indexed value cache therefore binds every point position.
        naming.need(len(indexed_numeric_values(series, "val")) == len(categories),
                    "Implicit native category positions differ from requested count")
        return
    caches = series.xpath('./*[local-name()="cat"]/*[local-name()="strRef" or local-name()="numRef"]/*[local-name()="strCache" or local-name()="numCache"]')
    naming.need(len(caches) == 1, "Error-bar category cache is missing or ambiguous")
    counts = caches[0].xpath('./*[local-name()="ptCount"]/@val')
    naming.need(counts == [str(len(categories))], "Error-bar category point count differs")
    values = {}
    for point in caches[0].xpath('./*[local-name()="pt"]'):
        index = point.get("idx", "")
        text = point.xpath('./*[local-name()="v"]/text()')
        naming.need(index.isdigit() and int(index) not in values and len(text) == 1,
                    "Error-bar category indices are ambiguous")
        values[int(index)] = text[0]
    naming.need(set(values) == set(range(len(categories))), "Error-bar category indices are incomplete")
    naming.need([values[index] for index in range(len(categories))] == [str(value) for value in categories],
                "Error-bar physical categories differ from requested category order")


def clipped_linear_domain(cache, lower, upper, slope, intercept):
    """Native partitions keep full endpoints; OOXML forecasts stop at plot edges."""
    axes = cache.xpath('.//*[local-name()="valAx"]')
    bounds = {}
    for kind, positions in (("x", {"b", "t"}), ("y", {"l", "r"})):
        candidates = [a for a in axes if a.xpath('./*[local-name()="axPos"]/@val') and a.xpath('./*[local-name()="axPos"]/@val')[0] in positions]
        naming.need(len(candidates) == 1, "Trendline domain requires unique X/Y axes")
        axis = candidates[0]
        naming.need(not axis.xpath('./*[local-name()="scaling"]/*[local-name()="logBase"]'), "Logarithmic trendline clipping is outside this profile")
        values = [axis.xpath('./*[local-name()="scaling"]/*[local-name()="' + field + '"]/@val') for field in ("min", "max")]
        naming.need(all(len(v) == 1 for v in values), "Trendline axes need explicit native minimum and maximum")
        bounds[kind] = [float(values[0][0]), float(values[1][0])]
    lower, upper = max(lower, bounds["x"][0]), min(upper, bounds["x"][1])
    if abs(slope) > 1e-14:
        yminx, ymaxx = sorted((y-intercept)/slope for y in bounds["y"])
        lower, upper = max(lower, yminx), min(upper, ymaxx)
    else:
        naming.need(bounds["y"][0] <= intercept <= bounds["y"][1], "Constant trendline is outside the plot")
    naming.need(lower <= upper, "Native trendline does not intersect its plot")
    return lower, upper


def nonlinear_fit(data, group, kind):
    """Independent OLS: exponential log(Y) on X, logarithmic Y on log(X)."""
    import math
    naming.need(kind in {"exp", "log"}, "Unverified nonlinear fit type")
    pairs = [(x,y) for g,x,y in zip(data["group_labels"],data["x_values"],data["y_values"]) if g == group]
    naming.need(len(pairs) >= 2 and all(math.isfinite(x) and math.isfinite(y) for x,y in pairs),
                "Nonlinear fit requires finite X/Y data and at least two points")
    naming.need(all(y > 0 for x,y in pairs) if kind == "exp" else all(x > 0 for x,y in pairs),
                "Exponential Y or logarithmic X must be positive")
    x = [a if kind == "exp" else math.log(a) for a,bb in pairs]
    y = [math.log(bb) if kind == "exp" else bb for a,bb in pairs]
    mx,my = sum(x)/len(x),sum(y)/len(y)
    denominator = sum((v-mx)**2 for v in x)
    naming.need(denominator > 0, "Nonlinear fit requires distinct X values")
    slope = sum((a-mx)*(bb-my) for a,bb in zip(x,y))/denominator
    intercept = my-slope*mx
    constant = math.exp(intercept) if kind == "exp" else intercept
    naming.need(math.isfinite(constant) and math.isfinite(slope) and slope > 0,
                "Nonlinear forecast profile requires a finite increasing fitted curve")
    return constant,slope


def verify_nonlinear_forecast(cache, trend, xmax, constant, slope, kind, xmin=None):
    """Compare saved upper forecast with a fit-dependent plot intersection."""
    import math
    naming.need(kind in {"exp", "log"} and math.isfinite(slope) and slope > 0 and math.isfinite(constant),
                "Nonlinear forecast profile requires a finite increasing fitted curve")
    lower_domain = verify_lower_forecast(trend, xmin)
    bounds = {}
    for axis_kind,positions in (("x", {"b", "t"}), ("y", {"l", "r"})):
        axes = [a for a in cache.xpath('.//*[local-name()="valAx"]')
                if a.xpath('./*[local-name()="axPos"]/@val') and a.xpath('./*[local-name()="axPos"]/@val')[0] in positions]
        naming.need(len(axes) == 1 and not axes[0].xpath('./*[local-name()="scaling"]/*[local-name()="logBase"]'),
                    "Nonlinear forecast requires unique linear X/Y axes")
        maxima = axes[0].xpath('./*[local-name()="scaling"]/*[local-name()="max"]/@val')
        naming.need(len(maxima) == 1 and math.isfinite(float(maxima[0])), "Nonlinear forecast requires explicit finite axis bounds")
        bounds[axis_kind] = float(maxima[0])
        if axis_kind == "x":
            minima = axes[0].xpath('./*[local-name()="scaling"]/*[local-name()="min"]/@val')
            naming.need(len(minima) == 1, "Nonlinear forecast requires explicit X minimum")
            bounds["xmin"] = float(minima[0])
    if kind == "exp":
        naming.need(constant > 0 and bounds["y"] > 0, "Exponential forecast requires positive fitted constant and Y extent")
        intersection = math.log(bounds["y"]/constant)/slope
    else:
        intersection = math.exp((bounds["y"]-constant)/slope)
    wanted = min(bounds["x"], intersection)
    values = trend.xpath('./*[local-name()="forward"]/@val')
    naming.need(len(values) <= 1 and len(values) == len(trend.xpath('./*[local-name()="forward"]')),
                "Nonlinear forecast is ambiguous or missing its value")
    actual = xmax + (float(values[0]) if values else 0.0)
    naming.need(forecast_close(actual,wanted,bounds["xmin"],bounds["x"]),
                "Physical nonlinear forecast endpoint contradicts independent fitted curve and plot bounds")
    return {"actual_upper_x":actual,"reference_upper_x":wanted,"fit_limits_visible_endpoint":wanted < bounds["x"], **lower_domain}


def power_fit(data, group):
    """Independent log-OLS reference, y = coefficient * x ** exponent."""
    import math
    pairs = [(x, y) for g, x, y in zip(data["group_labels"], data["x_values"], data["y_values"]) if g == group]
    naming.need(len(pairs) >= 2 and all(math.isfinite(x) and math.isfinite(y) and x > 0 and y > 0 for x,y in pairs),
                "Power profile requires finite positive X/Y data")
    x = [math.log(p[0]) for p in pairs]; y = [math.log(p[1]) for p in pairs]
    mx, my = sum(x)/len(x), sum(y)/len(y)
    denominator = sum((v-mx)**2 for v in x)
    naming.need(denominator > 0, "Power profile requires distinct X values")
    exponent = sum((a-mx)*(b-my) for a,b in zip(x,y))/denominator
    coefficient = math.exp(my - exponent*mx)
    return coefficient, exponent


def verify_power_forecast(cache, trend, xmax, coefficient, exponent, xmin=None):
    """Observe fit-dependent upper clipping in the regenerated physical forecast."""
    import math
    naming.need(math.isfinite(exponent) and exponent > 0 and math.isfinite(coefficient) and coefficient > 0,
                "Power forecast proof currently requires a finite positive increasing fitted curve")
    lower_domain = verify_lower_forecast(trend, xmin)
    bounds = {}
    for kind, positions in (("x", {"b", "t"}), ("y", {"l", "r"})):
        axes = [axis for axis in cache.xpath('.//*[local-name()="valAx"]')
                if axis.xpath('./*[local-name()="axPos"]/@val') and
                axis.xpath('./*[local-name()="axPos"]/@val')[0] in positions]
        naming.need(len(axes) == 1 and not axes[0].xpath('./*[local-name()="scaling"]/*[local-name()="logBase"]'),
                    "Power forecast proof requires unique linear X/Y axes")
        maxima = axes[0].xpath('./*[local-name()="scaling"]/*[local-name()="max"]/@val')
        naming.need(len(maxima) == 1 and math.isfinite(float(maxima[0])), "Power forecast proof requires explicit finite native axis bounds")
        bounds[kind] = float(maxima[0])
        if kind == "x":
            minima = axes[0].xpath('./*[local-name()="scaling"]/*[local-name()="min"]/@val')
            naming.need(len(minima) == 1, "Power forecast requires explicit X minimum")
            bounds["xmin"] = float(minima[0])
    naming.need(bounds["y"] > 0, "Power plot has no positive Y extent")
    wanted = min(bounds["x"], (bounds["y"]/coefficient)**(1/exponent))
    values = trend.xpath('./*[local-name()="forward"]/@val')
    naming.need(len(values) <= 1 and len(values) == len(trend.xpath('./*[local-name()="forward"]')),
                "Power forecast is ambiguous or missing its value")
    actual = xmax + (float(values[0]) if values else 0.0)
    naming.need(forecast_close(actual,wanted,bounds["xmin"],bounds["x"]),
                "Physical power forecast endpoint contradicts independent fitted curve and plot bounds")
    return {"actual_upper_x": actual, "reference_upper_x": wanted, "fit_limits_visible_endpoint": wanted < bounds["x"], **lower_domain}


def verify_legend_swatch(cache, swatch, values):
    """Bounded bar profile: swatch color belongs to the uniquely bound series."""
    carriers = [node for node in cache.xpath('.//*[local-name()="barChart"]/*[local-name()="ser"]')
                if indexed_numeric_values(node, "val") == values]
    naming.need(len(carriers) == 1, "Legend swatch requires a unique native bar series value binding")
    def fill(node):
        fills = node.xpath('./*[local-name()="spPr"]/*[local-name()="solidFill"]')
        naming.need(len(fills) == 1, "Legend series or swatch fill is missing or ambiguous")
        return [(E.QName(child).localname, sorted(child.attrib.items())) for child in fills[0].iter()]
    naming.need(fill(swatch) == fill(carriers[0]), "Legend swatch color differs from its owning native series")


def inspect_inserted(path, selector, feature, expected_request):
    """Read saved output; reject physical-only and model-only insertion claims."""
    import thinkcell
    _, charts, _ = naming.inventory(Path(path).read_bytes())
    target = naming.choose(charts, **selector)
    naming.need(target["owner"].find("m_bConsistent").get("val") == "1", "Saved chart is inconsistent")
    actual = thinkcell.baseline_request(Path(path), target)
    naming.need(actual == expected_request, "Saved native model or embedded data differs from requested data")
    root, ids = target["doc"]["root"], target["doc"]["ids"]
    series = series_refs(target)
    with zipfile.ZipFile(path) as z:
        if feature == "legend":
            legends = [x for x in root.findall("CSequenceChartLegendSE") if x.find("m_cse").get("idref") == target["owner"].get("id")]
            naming.need(len(legends) == 1, "Inserted native legend is absent or ambiguous")
            legend = legends[0]
            naming.need(legend.find("m_bConsistent").get("val") == "1", "Inserted legend is inconsistent")
            refs = root.findall("CContainerSE/m_cse/elem")
            naming.need(sum(e.get("idref") == legend.get("id") for e in refs) == 1, "Legend container ownership missing")
            slide = naming.logical_slides(z)[target["doc"]["slide_number"] - 1]["part"]
            drawing, rels = naming.xml(z.read(slide)), naming.relationship_map(z, slide)
            cache = naming.xml(z.read(target["frames"][0]["native_chart_part"]))
            tagged = {}
            for shape in drawing.findall(".//p:sp", naming.NS) + drawing.findall(".//p:cxnSp", naming.NS):
                for tagref in shape.findall(".//p:tags", naming.NS):
                    for tag in naming.xml(z.read(rels[tagref.get("{" + naming.NS["r"] + "}id")]["resolved"])):
                        if tag.get("name", "").upper() == "THINKCELLSHAPEDONOTDELETE":
                            naming.need(tag.get("val") not in tagged, "Native physical tag is ambiguous")
                            tagged[tag.get("val")] = shape
            names, parents = [], []
            for ref in legend.findall("./m_clegendentry/elem"):
                entry = ids[ref.get("idref")]
                sid = entry.find("m_featParent").get("idref")
                parents.append(sid)
                naming.need(sid in series, "Legend entry binds unrelated series")
                src = ids[ids[sid].find("m_varsrc").get("idref")]
                name = src.findtext("m_varval")
                label = ids[entry.find("m_legendlabel").get("idref")]
                shape_tag = label.findtext("m_ppttb/m_bstrShapeName")
                naming.need(shape_tag in tagged, "Native physical legend label missing")
                fields = tagged[shape_tag].findall(".//a:fld", naming.NS)
                formats = {"datetime" + ids[v.get("idref")].findtext("m_bstrFormat") for v in src.findall("./m_ctextvar/elem")}
                naming.need(len(fields) == 1 and fields[0].get("type") in formats, "Legend field lacks matching native series source")
                naming.need(fields[0].findtext("a:t", namespaces=naming.NS) == name, "Legend visible name is stale")
                rectangle = ids[entry.find("m_pptrect").get("idref")]
                swatch_tag = rectangle.findtext("m_bstrShapeName")
                naming.need(swatch_tag in tagged and rectangle.find("m_bPlaced").get("val") == "1", "Native physical legend swatch missing")
                verify_legend_swatch(cache, tagged[swatch_tag], actual["expected_model"]["series_values"][series.index(sid)])
                names.append(name)
            naming.need(len(parents) == len(series) and set(parents) == set(series), "Legend series coverage differs")
            return {"status": "NATIVE_MODEL_AND_PHYSICAL_SEMANTICS_PASS", "feature": feature, "series_names": names}
        if feature in TRENDLINE_PROFILES:
            type_code, physical_type = TRENDLINE_PROFILES[feature]
            partitions = [ids[x.get("idref")] for x in target["table"].findall("./m_cscatpartition/elem")]
            inserted = [x for x in partitions if x.find("m_epartitiontype").get("val") == "1"]
            naming.need(len(inserted) == 1, "Inserted linear partition is missing or ambiguous")
            partition = inserted[0]
            naming.need(partition.find("m_etrendlinetype").get("val") == type_code, "Trendline type changed")
            sid = partition.find("m_scatdseries").get("idref")
            drawing = naming.xml(z.read(target["frames"][0]["native_chart_part"]))
            physical = scatter_carrier(drawing, target, sid, actual["expected_model"])
            trends = physical.xpath('./*[local-name()="trendline"]')
            naming.need(len(trends) == 1 and trends[0].xpath('./*[local-name()="trendlineType"]/@val') == [physical_type], "Native physical trendline type differs")
            naming.need(not trends[0].xpath('./*[local-name()="intercept"]'), "Forced physical trendline intercept contradicts calculated model")
            orders = trends[0].xpath('./*[local-name()="order"]/@val')
            naming.need(orders == ([str(POLYNOMIAL_ORDERS[feature])] if feature in POLYNOMIAL_ORDERS else []),
                        "Polynomial order differs from authenticated native profile")
            group = ids[ids[sid].find("m_varsrc").get("idref")].findtext("m_varval")
            data = actual["expected_model"]
            if feature in POLYNOMIAL_ORDERS:
                coefficients = polynomial_fit(data, group, POLYNOMIAL_ORDERS[feature])
                naming.need(partition.find("m_linefPartition") is None, "Polynomial partition retains a stale linear segment")
                native_curve = partition.find("m_pptfreeform/m_bstrShapeName")
                naming.need(native_curve is not None and bool(native_curve.text), "Polynomial native curve identity is missing")
                xmax = max(x for g,x in zip(data["group_labels"],data["x_values"]) if g == group)
                forecast = verify_polynomial_forecast(drawing, trends[0], xmax, coefficients, min(x for g,x in zip(data["group_labels"],data["x_values"]) if g == group))
                return {"status": "NATIVE_MODEL_AND_PHYSICAL_SEMANTICS_PASS", "feature": feature, "series": group,
                        "reference_coefficients_ascending": coefficients, "polynomial_order": POLYNOMIAL_ORDERS[feature],
                        "physical_forecast": forecast,
                        "fit_evidence": "independent polynomial least squares checked against observed physical forecast clipping; native model does not serialize fit coefficients"}
            if feature in {"exponential-trendline", "logarithmic-trendline"}:
                coefficient, slope = nonlinear_fit(data, group, physical_type)
                naming.need(partition.find("m_linefPartition") is None, "Nonlinear partition retains a stale linear segment")
                native_curve = partition.find("m_pptfreeform/m_bstrShapeName")
                naming.need(native_curve is not None and bool(native_curve.text), "Nonlinear native curve identity is missing")
                xmax = max(x for g,x in zip(data["group_labels"],data["x_values"]) if g == group)
                forecast = verify_nonlinear_forecast(drawing, trends[0], xmax, coefficient, slope, physical_type, min(x for g,x in zip(data["group_labels"],data["x_values"]) if g == group))
                return {"status": "NATIVE_MODEL_AND_PHYSICAL_SEMANTICS_PASS", "feature": feature, "series": group,
                        "reference_constant": coefficient, "reference_slope": slope,
                        "physical_forecast": forecast,
                        "fit_evidence": "independent OLS checked against observed physical forecast clipping; native model does not serialize fit coefficients"}
            if feature == "power-trendline":
                coefficient, exponent = power_fit(data, group)
                naming.need(partition.find("m_linefPartition") is None, "Power partition retains a stale linear segment")
                native_curve = partition.find("m_pptfreeform/m_bstrShapeName")
                naming.need(native_curve is not None and bool(native_curve.text), "Power native curve identity is missing")
                xmax = max(x for g,x in zip(data["group_labels"],data["x_values"]) if g == group)
                forecast = verify_power_forecast(drawing, trends[0], xmax, coefficient, exponent, min(x for g,x in zip(data["group_labels"],data["x_values"]) if g == group))
                return {"status": "NATIVE_MODEL_AND_PHYSICAL_SEMANTICS_PASS", "feature": feature, "series": group,
                        "reference_coefficient": coefficient, "reference_exponent": exponent,
                        "physical_forecast": forecast,
                        "fit_evidence": "independent log-OLS checked against observed physical forecast clipping; native model does not serialize fit coefficients"}
            pairs = [(x,y) for g,x,y in zip(data["group_labels"], data["x_values"], data["y_values"]) if g == group]
            avgx = sum(x for x,y in pairs)/len(pairs); avgy = sum(y for x,y in pairs)/len(pairs)
            slope = sum((x-avgx)*(y-avgy) for x,y in pairs)/sum((x-avgx)**2 for x,y in pairs)
            intercept = avgy - slope*avgx
            import math
            endpoints = partition.findall("m_linefPartition/m_ptFrom") + partition.findall("m_linefPartition/m_ptTo")
            naming.need(len(endpoints) == 2 and all(math.isclose(float(p.get("y")), slope*float(p.get("x"))+intercept, abs_tol=1e-8) for p in endpoints), "Native partition endpoints do not reflect requested regression")
            # The native model may extend to the visible axis range. Physical
            # forward/backward values must describe precisely the same domain.
            xmin, xmax = min(x for x,y in pairs), max(x for x,y in pairs)
            model_xmin, model_xmax = sorted(float(p.get("x")) for p in endpoints)
            model_xmin, model_xmax = clipped_linear_domain(drawing, model_xmin, model_xmax, slope, intercept)
            extensions = {}
            for field in ("forward", "backward"):
                vals = trends[0].xpath('./*[local-name()="' + field + '"]/@val')
                naming.need(len(vals) <= 1, "Duplicate physical trendline extension")
                extensions[field] = float(vals[0]) if vals else 0.0
            naming.need(math.isclose(xmax + extensions["forward"], model_xmax, abs_tol=1e-8) and
                        math.isclose(xmin - extensions["backward"], model_xmin, abs_tol=1e-8), "Physical trendline domain contradicts native model endpoints")
            return {"status": "NATIVE_MODEL_AND_PHYSICAL_SEMANTICS_PASS", "feature": feature, "series": group, "slope": slope, "intercept": intercept}
        naming.need(feature in {"errorbar-range", "errorbar-caps"}, "Unknown inspection feature")
        range_refs = target["table"].findall("./m_cscserrange/elem")
        naming.need(len(range_refs) == 1, "Inserted range is absent or ambiguous")
        interval = ids[range_refs[0].get("idref")]
        paired = [x.get("idref") for x in interval.findall("./m_cscdseries/elem")]
        naming.need(len(paired) == 2 and len(set(paired)) == 2 and set(paired) <= set(series), "Native range ownership is invalid")
        part = target["frames"][0]["native_chart_part"]
        naming.need(part, "Selected chart has no physical native chart cache")
        drawing = naming.xml(z.read(part))
        physical_series = drawing.xpath('.//*[local-name()="plotArea"]/*[local-name()="lineChart" or local-name()="scatterChart"]/*[local-name()="ser"]')
        errors = [(s, e) for s in physical_series for e in s.xpath('./*[local-name()="errBars"]')]
        naming.need(len(errors) == 1, "Native custom error-bar cache missing or ambiguous")
        carrier, error = errors[0]
        direction = "y" if target["owner"].find("m_eorient").get("val") == "0" else "x"
        for physical in physical_series:
            verify_category_positions(physical, direction, actual["expected_model"]["categories"])
        naming.need(error.xpath('./*[local-name()="errDir"]/@val') == [direction], "Error-bar direction differs")
        naming.need(error.xpath('./*[local-name()="errValType"]/@val') == ["cust"], "Error-bar is not a native custom interval")
        naming.need(error.xpath('./*[local-name()="errBarType"]/@val') == ["plus"] and not error.xpath('./*[local-name()="minus"]'), "Error-bar must use the native signed plus-only profile")
        values = actual["expected_model"]["series_values"]
        left, right = [values[series.index(s)] for s in paired]
        expected_signed = [[b-a for a,b in zip(left,right)], [a-b for a,b in zip(left,right)]]
        extents = indexed_numeric_values(error, "plus")
        owner_values = physical_endpoint_values(carrier, direction)
        naming.need(owner_values in (left, right) and left != right, "Error-bar physical carrier does not bind a unique endpoint series")
        signed = expected_signed[0] if owner_values == left else expected_signed[1]
        naming.need(extents == signed, "Native signed interval extents do not match their owning endpoint data")
        if feature == "errorbar-caps":
            for vector_ref in target["table"].findall("./ocol/elem"):
                scalar_refs = ids[vector_ref.get("idref")].findall("./ocol/elem")
                for sid in paired:
                    scalar = ids[scalar_refs[series.index(sid)].get("idref")]
                    marker = scalar.find("m_markerprops/m_emarkerstyle")
                    naming.need(marker is not None and marker.get("val") == "dash", "Native endpoint cap marker missing")
            # Actual native dash caps are per-point dPt marker overrides in
            # current donors, even when the series default symbol is none.
            for endpoint in (left, right):
                endpoint_carriers = []
                for s in physical_series:
                    if physical_endpoint_values(s, direction) == endpoint:
                        endpoint_carriers.append(s)
                naming.need(len(endpoint_carriers) == 1, "Endpoint cap carrier is missing or ambiguous")
                physical = endpoint_carriers[0]
                default_dash = physical.xpath('./*[local-name()="marker"]/*[local-name()="symbol"]/@val') == ["dash"]
                overrides = physical.xpath('./*[local-name()="dPt"]')
                cap_indices = {int(point.xpath('string(./*[local-name()="idx"]/@val)')) for point in overrides
                               if point.xpath('./*[local-name()="marker"]/*[local-name()="symbol"]/@val') == ["dash"]}
                naming.need(default_dash or cap_indices == set(range(len(endpoint))), "Physical native dash caps do not cover every endpoint")
        return {"status": "NATIVE_MODEL_AND_PHYSICAL_SEMANTICS_PASS", "feature": feature, "signed_extents": extents}


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "verify":
        ap = argparse.ArgumentParser(description="Read-only native model, data and physical insertion gates")
        ap.add_argument("--input", type=Path, required=True)
        ap.add_argument("--feature", choices=sorted(INSERTION_FEATURES), required=True)
        ap.add_argument("--selector", required=True)
        ap.add_argument("--request", type=Path, required=True)
        ap.add_argument("--expected-sha256", required=True)
        args = ap.parse_args(sys.argv[2:])
        naming.need(naming.sha(args.input.read_bytes()) == args.expected_sha256.upper(), "Readback artifact hash changed")
        print(json.dumps(inspect_inserted(args.input, json.loads(args.selector), args.feature,
                                         json.loads(args.request.read_text(encoding="utf-8-sig"))), indent=2))
        naming.need(naming.sha(args.input.read_bytes()) == args.expected_sha256.upper(), "Readback artifact changed during inspection")
        return
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--source", type=Path, required=True)
    ap.add_argument("--donor", type=Path, required=True)
    ap.add_argument("--output", type=Path, required=True)
    ap.add_argument("--feature", choices=sorted(INSERTION_FEATURES), required=True)
    ap.add_argument("--selector", required=True, help="JSON exact chart selectors")
    ap.add_argument("--donor-selector", required=True, help="JSON exact donor chart selectors")
    ap.add_argument("--source-sha256", required=True)
    ap.add_argument("--donor-sha256", required=True)
    ap.add_argument("--target-series-id")
    ap.add_argument("--donor-feature-id")
    args = ap.parse_args()
    print(json.dumps(prepare(args.source, args.output, args.donor, args.feature,
                             json.loads(args.selector), json.loads(args.donor_selector),
                             args.source_sha256, args.donor_sha256, target_series_id=args.target_series_id,
                             donor_feature_id=args.donor_feature_id), indent=2))


if __name__ == "__main__":
    main()
