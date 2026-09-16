"""Prepare coherent native think-cell label edits on a copy.

The adapter is deliberately model-first.  It targets an existing native label
by semantic type/index, edits its active precision/font/placement fields, and
keeps the native CFB carrier intact.  It can also insert a scalar label by
cloning an unplaced scalar-label model from the same chart.  Official JSON
regeneration, native reopen and changed-data checks remain required.

No PowerPoint or ppttc access occurs here.
"""
from __future__ import annotations
import argparse, base64, copy, hashlib, io, json, os, posixpath, re, subprocess, sys, tempfile, zipfile
from pathlib import Path
from lxml import etree as E

HERE = Path(__file__).resolve().parent


def discover_script_dir() -> Path:
    """Find the helper bundle beside a relocated adapter.

    The sentinel is ``chart_geometry.py``.  The explicit environment path is
    checked first; adjacent bundle layouts are then searched before the lane
    checkout fallback used during development.
    """
    override = os.environ.get("THINKCELL_LABEL_SCRIPTS_ROOT")
    if override:
        return Path(override).expanduser().resolve()
    candidates = []
    for ancestor in (HERE, *HERE.parents):
        candidates.extend((ancestor,
                           ancestor / "scripts",
                           ancestor / "staged-plugin" / "skills" / "thinkcell-edit" / "scripts",
                           ancestor / "skills" / "thinkcell-edit" / "scripts"))
    for candidate in candidates:
        if (candidate / "chart_geometry.py").is_file():
            return candidate.resolve()
    # Development checkout fallback.  Import below provides the actionable
    # missing-dependency error if this path has also been removed.
    return (HERE.parents[1] / "staged-plugin" / "skills" / "thinkcell-edit" / "scripts").resolve()


SCRIPTS = discover_script_dir()
sys.path.insert(0, str(SCRIPTS))
from chart_geometry import inventory, need, sha, streams, xml
from runtime import powershell, powershell_env

LABEL_TYPES = {
    "scalar": "CSequenceChartDataScalarLabel",
    "sum": "CSequenceChartDataSumLabel",
    "category": "CSequenceChartDataCategoryLabel",
    "series": "CSequenceChartDataSeriesLabel",
}
POSITION_ENUMS = {-2, 0, 1, 3, 13, 14, 103, 105}
HEX = re.compile(r"^[0-9A-Fa-f]{6}$")
PPT_NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def canonical(b: bytes) -> bytes:
    return E.tostring(xml(b), method="c14n")


def quoted_format_key(display: str) -> str:
    """Encode a baseline display using think-cell's value-key grammar."""
    need(isinstance(display, str) and display and "\x00" not in display,
         "physical_field.baseline_text must be nonempty text")
    return "".join("'" + char + "'" for char in display)


def tagged_physical_shape(slide_root, rel_root, zip_file, shape_tag: str):
    """Find one physical shape by its authored think-cell tag value."""
    relmap = {r.get("Id"): r.get("Target") for r in rel_root}
    matches = []
    for shape in slide_root.xpath(".//p:sp", namespaces=PPT_NS):
        refs = shape.xpath("./p:nvSpPr/p:nvPr/p:custDataLst/p:tags/@r:id", namespaces=PPT_NS)
        if not refs or refs[0] not in relmap:
            continue
        part = posixpath.normpath(posixpath.join("ppt/slides", relmap[refs[0]]))
        try:
            tag_xml = zip_file.read(part)
        except KeyError:
            continue
        if shape_tag.encode("utf-8") in tag_xml:
            matches.append((shape, part))
    need(len(matches) == 1, f"physical shape tag must resolve exactly once; found {len(matches)}")
    return matches[0]


def apply_physical_field(shape, controls: dict):
    """Synchronize one existing dynamic field with a bound model variable.

    The field is selected by its caller-supplied baseline text, never by a
    shape or field id.  All other fields/runs remain byte-for-byte unchanged.
    """
    spec = controls.get("physical_field")
    if spec is None:
        return {"changed": False}
    need(isinstance(spec, dict), "physical_field must be an object")
    baseline = spec.get("baseline_text")
    display = spec.get("baseline_display", baseline)
    key = spec.get("baseline_format_key") or quoted_format_key(baseline)
    need(isinstance(key, str) and key and "\x00" not in key,
         "physical_field.baseline_format_key must be nonempty text")
    fields = shape.xpath(".//a:fld", namespaces=PPT_NS)
    matches = [f for f in fields if f.findtext("a:t", namespaces=PPT_NS) == baseline]
    need(len(matches) == 1,
         "physical_field.baseline_text must identify exactly one existing field")
    field = matches[0]
    run = field.find("a:rPr", namespaces=PPT_NS)
    need(run is not None, "selected physical field lacks a run style")
    old = {"type": field.get("type"), "text": field.findtext("a:t", namespaces=PPT_NS),
           "size": run.get("sz"), "bold": run.get("b")}
    field.set("type", "datetime" + key)
    field.find("a:t", namespaces=PPT_NS).text = display
    sync = spec.get("sync_model_style", False)
    if sync:
        label_style = controls
        if "font_size" in label_style:
            run.set("sz", str(round(float(label_style["font_size"]) * 100)))
        if "bold" in label_style:
            run.set("b", "1" if label_style["bold"] else "0")
    return {"changed": True, "field_index": fields.index(field),
            "baseline_text": baseline, "display": display,
            "format_key": key, "old": old,
            "sync_model_style": bool(sync)}


def active_nodes(node, name: str):
    return [x for x in node.findall(name) if int(x.get("reqver", "0")) <= 38764 < int(x.get("endver", "999999999"))]


def all_labels(root, label_type: str):
    tag = LABEL_TYPES.get(label_type, label_type)
    return [n for n in root if n.tag == tag and n.get("id")]


def id_map(root):
    """Return the chart-local native object registry.

    A think-cell model is a flat object graph.  The order in which labels are
    serialized is therefore an implementation detail and can change after a
    native save.  Semantic resolution below follows the chart's explicit
    vector/series references instead of depending on that order.
    """
    return {n.get("id"): n for n in root if n.get("id")}


def source_text(source):
    """Read a native text variable source without decoding its display key."""
    if source is None:
        return None
    value = source.find("m_varval")
    if value is None or value.get("type") not in (None, "5"):
        return None
    return value.text or ""


def chart_series(root):
    """Return series names in the same order used by each vector's ``ocol``.

    ``CSequenceChartDataSeries`` objects are commonly serialized in reverse
    order.  The chart-level ``m_cscdser`` sequence is the authoritative order
    for scalar columns, so it is preferred over document order.
    """
    ids = id_map(root)
    sequences = root.findall(".//m_cscdser")
    for sequence in sequences:
        refs = [e.get("idref") for e in sequence.findall("elem") if e.get("idref")]
        result = []
        for ref in refs:
            node = ids.get(ref)
            if node is None or node.tag != "CSequenceChartDataSeries":
                continue
            src_ref = node.find("m_varsrc")
            src = ids.get(src_ref.get("idref")) if src_ref is not None else None
            result.append((node, source_text(src)))
        if result:
            return result
    return [(node, source_text(ids.get(node.find("m_varsrc").get("idref"))))
            for node in root if node.tag == "CSequenceChartDataSeries"]


def scalar_context(root, scalar):
    """Return ``(category, series)`` for a scalar from explicit model links."""
    ids = id_map(root)
    category = None
    series_index = None
    for vector in root.iter("CSequenceChartDataVector"):
        elems = vector.findall("ocol/elem")
        for i, elem in enumerate(elems):
            if elem.get("idref") == scalar.get("id"):
                category = source_text(vector.find("m_varsrcCategory"))
                series_index = i
                break
        if series_index is not None:
            break
    series = None
    ordered = chart_series(root)
    if series_index is not None and series_index < len(ordered):
        series = ordered[series_index][1]
    return category, series


def scalar_for_semantics(root, spec):
    """Resolve a scalar by category/series, with strict ambiguity checks."""
    category = spec.get("category")
    series = spec.get("series")
    for key in ("label_index", "label_id", "scalar_index", "scalar_id"):
        need(spec.get(key) is None,
             f"{key} cannot be combined with category/series selectors")
    if category is not None:
        need(isinstance(category, str) and category, "category must be nonempty text")
    if series is not None:
        need(isinstance(series, str) and series, "series must be nonempty text")
    scalars = list(root.iter("CSequenceChartDataScalar"))
    if category is None and series is None:
        return None
    matches = []
    for scalar in scalars:
        actual_category, actual_series = scalar_context(root, scalar)
        if category is not None and actual_category != category:
            continue
        if series is not None and actual_series != series:
            continue
        matches.append(scalar)
    need(len(matches) == 1,
         "Semantic scalar selector must resolve exactly one native scalar "
         f"(category={category!r}, series={series!r}); found {len(matches)}")
    return matches[0]


def scalar_label_for_semantics(root, spec):
    scalar = scalar_for_semantics(root, spec)
    if scalar is None:
        return None, None
    link = scalar.find("m_scdlabel")
    need(link is not None and link.get("idref") not in (None, "0"),
         "Selected semantic scalar has no native scalar label")
    labels = [x for x in all_labels(root, "scalar") if x.get("id") == link.get("idref")]
    need(len(labels) == 1, "Selected semantic scalar label is not unique")
    return labels[0], scalar


def require_numeric_scalar(root, scalar):
    """Guard semantic label edits against text-only or stale scalar owners."""
    ids = id_map(root)
    ref = scalar.find("m_varsrcAbsolute")
    need(ref is not None and ref.get("idref") in ids,
         "Selected scalar has no native absolute numeric source")
    source = ids[ref.get("idref")]
    value = source.find("m_varval")
    need(value is not None and value.get("type") == "1" and value.get("val") is not None,
         "Selected scalar absolute source is not a live numeric field")
    return value.get("type")


def label_owner(root, label):
    candidates = []
    for n in root.iter():
        for child in n:
            if child.get("idref") == label.get("id"):
                if n is not root:
                    candidates.append((n, child.tag))
    if not candidates:
        return None, None
    # Prefer the semantic owning object over the root registry entry.
    for preferred in ("CSequenceChartDataScalar", "CSequenceChartDataScalarGroup",
                      "CSequenceChartDataVector", "CSequenceChartLegendEntry",
                      "CScatterChartDataVector"):
        for n, field in candidates:
            if n.tag == preferred:
                return n, field
    return candidates[0]


def find_chart(raw: bytes, target: dict):
    _, charts, _ = inventory(raw)
    matches = []
    for c in charts:
        if target.get("slide_number") is not None and c["doc"]["slide_number"] != target["slide_number"]:
            continue
        if target.get("slide_id") is not None and c["doc"]["slide_id"] != target["slide_id"]:
            continue
        if target.get("shape_tag") and target["shape_tag"] not in c["tags"]:
            continue
        matches.append(c)
    need(len(matches) == 1, "Target chart selector is not unique")
    return matches[0]


def resolve_label(root, spec: dict):
    # Category/series selectors follow the native scalar owner and remain
    # stable when object/label serialization order changes after save/reopen.
    if spec.get("category") is not None or spec.get("series") is not None:
        need(spec.get("label_type", "scalar") == "scalar",
             "category/series selectors currently resolve scalar labels only")
        label, scalar = scalar_label_for_semantics(root, spec)
        require_numeric_scalar(root, scalar)
        owner, field = label_owner(root, label)
        need(owner is scalar and field == "m_scdlabel",
             "Semantic scalar label owner is not the selected numeric scalar")
        return label, owner, field
    labels = all_labels(root, spec.get("label_type", "scalar"))
    if spec.get("label_id") is not None:
        labels = [x for x in labels if x.get("id") == str(spec["label_id"])]
    elif spec.get("label_index") is not None:
        i = spec["label_index"]
        need(isinstance(i, int) and not isinstance(i, bool) and i >= 0, "label_index must be a nonnegative integer")
        labels = labels[i:i + 1]
    need(len(labels) == 1, "Selected native label is not unique")
    owner, field = label_owner(root, labels[0])
    return labels[0], owner, field


def resolve_scalar(root, spec: dict):
    semantic = scalar_for_semantics(root, spec)
    if semantic is not None:
        return semantic
    scalars = list(root.iter("CSequenceChartDataScalar"))
    if spec.get("scalar_id") is not None:
        scalars = [x for x in scalars if x.get("id") == str(spec["scalar_id"])]
    elif spec.get("scalar_index") is not None:
        i = spec["scalar_index"]
        need(isinstance(i, int) and not isinstance(i, bool) and i >= 0, "scalar_index must be a nonnegative integer")
        scalars = scalars[i:i + 1]
    need(len(scalars) == 1, "Selected scalar is not unique")
    return scalars[0]


def set_text(node, child: str, value: str):
    x = node.find(child)
    need(x is not None, f"Native node lacks {child}")
    x.text = value


def set_val(node, child: str, value: str):
    x = node.find(child)
    need(x is not None, f"Native node lacks {child}")
    x.set("val", value)


def apply_font(label, spec):
    changed = []
    size = spec.get("font_size")
    if size is not None:
        need(isinstance(size, (int, float)) and not isinstance(size, bool) and 1 <= float(size) <= 72, "font_size must be 1..72 pt")
        value = f"{float(size):.20E}"
        for font in [label.find("m_font"), label.find("m_ppttb/m_font")]:
            if font is not None and font.find("m_nSize") is not None:
                font.find("m_nSize").set("val", value); changed.append("m_font/m_nSize")
    for key, child in (("bold", "m_bBold"), ("italic", "m_bItalic")):
        if key in spec:
            need(isinstance(spec[key], bool), f"{key} must be boolean")
            for font in [label.find("m_font"), label.find("m_ppttb/m_font")]:
                if font is not None and font.find(child) is not None:
                    font.find(child).set("val", "1" if spec[key] else "0"); changed.append(child)
    if "theme_color_index" in spec:
        i = spec["theme_color_index"]
        need(isinstance(i, int) and not isinstance(i, bool) and 0 <= i <= 63, "theme_color_index must be 0..63")
        found = 0
        for col in [label.find("m_colFont"), label.find("m_ppttb/m_colFont")]:
            if col is not None and col.find("m_msothmcolidx") is not None:
                col.find("m_msothmcolidx").set("val", str(i)); changed.append("m_colFont/m_msothmcolidx"); found += 1
        need(found > 0, "Label has no native theme font-color field")
    need("rgb_color" not in spec, "Explicit RGB label color remains untested; use theme_color_index")
    return changed


def apply_precision(label, owner, spec):
    requested = any(k in spec for k in ("prefix", "suffix", "decimal_digits", "magnitude", "number_format_key"))
    if not requested:
        return []
    p = label.find("m_prec")
    changed = []
    if p is not None and "prefix" in spec:
        need(isinstance(spec["prefix"], str) and "\x00" not in spec["prefix"], "prefix must be text")
        set_text(p, "m_strPrefix", spec["prefix"]); changed.append("m_prec/m_strPrefix")
    if p is not None and "suffix" in spec:
        need(isinstance(spec["suffix"], str) and "\x00" not in spec["suffix"], "suffix must be text")
        set_text(p, "m_strSuffix17909", spec["suffix"]); changed.append("m_prec/m_strSuffix17909")
    if p is not None and "decimal_digits" in spec:
        d = spec["decimal_digits"]
        need(isinstance(d, int) and not isinstance(d, bool) and 0 <= d <= 15, "decimal_digits must be 0..15")
        set_val(p, "m_nDecimalDigits17909", str(d)); changed.append("m_prec/m_nDecimalDigits17909")
    if p is not None and "magnitude" in spec:
        m = spec["magnitude"]
        need(isinstance(m, int) and not isinstance(m, bool) and -15 <= m <= 15, "magnitude must be -15..15")
        set_val(p, "m_nMagnitude17909", str(m)); changed.append("m_prec/m_nMagnitude17909")
    # A bound CTextVariable is a second native owner.  Keep its precision and
    # key synchronized; never synthesize a key from a display string.
    if owner is not None:
        for ref in owner.findall("m_varsrcAbsolute") + owner.findall("m_varsrcAbsoluteSum"):
            rid = ref.get("idref")
            if not rid: continue
            root = label.getroottree().getroot(); src = root.find(f"./CVariableSource[@id='{rid}']")
            if src is None: continue
            tvref = src.find("m_ctextvar/elem")
            if tvref is None: continue
            tv = root.find(f"./CTextVariable[@id='{tvref.get('idref')}']")
            if tv is None: continue
            tp = tv.find("m_prec17834")
            need(tp is not None, "Bound text variable lacks active precision")
            for child in ("m_strPrefix", "m_strSuffix17909", "m_nDecimalDigits17909", "m_nMagnitude17909"):
                dst_node = tp.find(child)
                if dst_node is not None and child in ("m_strPrefix", "m_strSuffix17909", "m_nDecimalDigits17909", "m_nMagnitude17909"):
                    if child == "m_nDecimalDigits17909" and "decimal_digits" not in spec: continue
                    if child == "m_nMagnitude17909" and "magnitude" not in spec: continue
                    if child == "m_strPrefix" and "prefix" not in spec: continue
                    if child == "m_strSuffix17909" and "suffix" not in spec: continue
                    if child == "m_nDecimalDigits17909": dst_node.set("val", str(spec["decimal_digits"]))
                    elif child == "m_nMagnitude17909": dst_node.set("val", str(spec["magnitude"]))
                    elif child == "m_strPrefix": dst_node.text = spec["prefix"]
                    else: dst_node.text = spec["suffix"]
            if "number_format_key" in spec:
                key = spec["number_format_key"]
                need(isinstance(key, str) and key and "\x00" not in key, "number_format_key must be nonempty text")
                fmt = tv.find("m_bstrFormat"); need(fmt is not None, "Bound text variable lacks m_bstrFormat")
                fmt.text = key; changed.append("CTextVariable/m_bstrFormat")
            elif any(k in spec for k in ("prefix", "suffix", "decimal_digits", "magnitude")):
                need("number_format_key" in spec, "Bound label format changes require exact number_format_key")
    elif "number_format_key" in spec:
        need(False, "number_format_key requires a bound CTextVariable")
    if p is None and not changed:
        need(False, "Selected label is not a numeric precision owner")
    return changed


def apply_position(label, spec):
    if "label_place" not in spec:
        return []
    place = spec["label_place"]
    need(isinstance(place, int) and not isinstance(place, bool) and place in POSITION_ENUMS, "label_place is not a known native enum")
    nodes = label.findall("m_elabelplace")
    need(nodes, "Selected label lacks native m_elabelplace")
    for x in nodes: x.set("val", str(place))
    if "auto_placed" in spec:
        need(isinstance(spec["auto_placed"], bool), "auto_placed must be boolean")
        for x in label.findall("m_bAutoPlaced"): x.set("val", "1" if spec["auto_placed"] else "0")
    return ["m_elabelplace/@val"]


def fresh_id(root):
    nums=[int(x.get("id")) for x in root if x.get("id", "").isdigit()]
    return str(max(nums, default=0)+1)


def fresh_tag(root, seed):
    used={x.text for x in root.findall(".//m_bstrShapeName") if x.text}
    i=0
    while True:
        tag="t"+base64.urlsafe_b64encode(hashlib.sha256(f"{seed}|{i}".encode()).digest()).decode().rstrip("=")[:22]
        if tag not in used: return tag
        i+=1


def insert_scalar_label(root, spec):
    target = resolve_scalar(root, spec)
    if spec.get("category") is not None or spec.get("series") is not None:
        require_numeric_scalar(root, target)
    old = target.find("m_scdlabel")
    need(old is not None and old.get("idref") in (None, "0"), "Target scalar already owns a scalar label")
    templates = [x for x in root if x.tag == "CSequenceChartDataScalarLabel" and x.find("m_ppttb/m_bPlaced") is not None and x.find("m_ppttb/m_bPlaced").get("val") == "0"]
    need(templates, "No unplaced native scalar-label template exists on this chart")
    template = templates[0]
    clone=copy.deepcopy(template); nid=fresh_id(root); clone.set("id",nid)
    for x in clone.iter():
        if x is not clone and x.get("idref") not in (None,"0"):
            need(False, "Template has owned references; use a same-chart closure donor")
        need("guid" not in x.tag.lower() and not x.get("guid"), "Template has GUID-owned children")
    tag = fresh_tag(root, template.findtext("m_ppttb/m_bstrShapeName", "scalar"))
    clone.find("m_ppttb/m_bstrShapeName").text=tag
    target.find("m_scdlabel").set("idref",nid)
    root.append(clone)
    return clone, {"new_label_id":nid,"new_shape_tag":tag,"target_scalar_id":target.get("id")}


def mutate(root, spec):
    report={"changed_paths":[],"mode":"existing"}
    if spec.get("insert_scalar"):
        label, extra=insert_scalar_label(root,spec); report.update(extra,mode="insert_scalar")
        owner, _ = label_owner(root,label)
    else:
        label, owner, field=resolve_label(root,spec); report.update(label_id=label.get("id"),label_type=label.tag,owner=owner.tag if owner is not None else None)
        if spec.get("category") is not None or spec.get("series") is not None:
            report["semantic_selector"] = {
                key: spec[key] for key in ("category", "series") if spec.get(key) is not None
            }
    report["changed_paths"] += apply_precision(label,owner,spec)
    report["changed_paths"] += apply_font(label,spec)
    report["changed_paths"] += apply_position(label,spec)
    return report


def make_plan(path: Path, target: dict, controls: dict, allow_experimental_insertion: bool = False):
    raw=path.read_bytes(); chart=find_chart(raw,target); root=chart["doc"]["root"]
    need(controls.get("kind", "native_label_controls") == "native_label_controls", "Use native_label_controls kind")
    need(not controls.get("percent_conversion"), "Percent-only conversion belongs to the percent lane")
    if controls.get("insert_scalar"):
        need(allow_experimental_insertion,
             "insert_scalar is experimental; pass --experimental-insertion explicitly")
    preview=copy.deepcopy(root); summary=mutate(preview,dict(controls));
    need(summary["changed_paths"], "Specify at least one label control")
    physical_summary = None
    if controls.get("physical_field") is not None:
        slide_number = target.get("slide_number")
        label, _owner, _field = resolve_label(root, controls)
        derived_tag = label.findtext("m_ppttb/m_bstrShapeName")
        shape_tag = controls["physical_field"].get("shape_tag") or derived_tag
        need(isinstance(slide_number, int) and isinstance(shape_tag, str) and shape_tag,
             "physical_field requires a placed label shape tag")
        slide_part = f"ppt/slides/slide{slide_number}.xml"
        rel_part = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            slide = E.fromstring(z.read(slide_part)); rels = E.fromstring(z.read(rel_part))
            shape, _ = tagged_physical_shape(slide, rels, z, shape_tag)
            physical_summary = apply_physical_field(copy.deepcopy(shape), controls)
        need(controls.get("number_format_key") == physical_summary.get("format_key"),
             "physical_field format key must equal the bound model number_format_key")
    return {"schema_version":1,"kind":"native_label_controls","source_sha256":sha(raw),"target":target,"controls":controls,"experimental_insertion":bool(controls.get("insert_scalar")),"preview":summary,"physical_preview":physical_summary,"carrier":chart["doc"]["part"]}


def replace_carrier(raw: bytes, chart, after: bytes, output: Path,
                    slide_part: str | None = None, slide_after: bytes | None = None):
    with tempfile.TemporaryDirectory(prefix="native_label_",dir=output.parent) as td:
        td=Path(td); carrier=td/"carrier.bin"; payload=td/"model.xml"; carrier.write_bytes(chart["doc"]["ole"]); payload.write_bytes(after)
        cmd=[powershell(),"-NoProfile","-ExecutionPolicy","RemoteSigned","-File",str(SCRIPTS/"thinkcell_no_click"/"implementation"/"replace_ole_stream.ps1"),"-StoragePath",str(carrier),"-StreamBytesPath",str(payload)]
        p=subprocess.run(cmd,capture_output=True,text=True,env=powershell_env(),timeout=60)
        need(p.returncode==0,p.stderr[-500:]); changed=carrier.read_bytes(); updated=streams(changed)
        old=chart["doc"]["streams"]; need(set(updated)==set(old) and all(updated[k]==v for k,v in old.items() if k!=('think-cellXML',)),"Non-model CFB stream drift")
        with zipfile.ZipFile(io.BytesIO(raw)) as zin, zipfile.ZipFile(output,"x") as zout:
            zout.comment=zin.comment
            for item in zin.infolist():
                if item.filename == chart["doc"]["part"]:
                    data = changed
                elif slide_part is not None and item.filename == slide_part:
                    need(slide_after is not None, "slide_after is required with slide_part")
                    data = slide_after
                else:
                    data = zin.read(item.filename)
                zout.writestr(copy.copy(item), data)


def prepare(path: Path, plan_path: Path, output: Path, report_path: Path,
            allow_experimental_insertion: bool = False):
    raw=path.read_bytes(); plan=json.loads(plan_path.read_text(encoding="utf-8-sig"))
    if plan.get("controls", {}).get("insert_scalar"):
        need(allow_experimental_insertion and plan.get("experimental_insertion") is True,
             "insert_scalar is experimental; pass --experimental-insertion explicitly")
    need(sha(raw)==plan["source_sha256"].upper(),"Source hash changed; inspect again")
    need(len({path.resolve(),plan_path.resolve(),output.resolve(),report_path.resolve()})==4 and not output.exists() and not report_path.exists(),"Use distinct new paths")
    chart=find_chart(raw,plan["target"]); before=chart["doc"]["streams"][("think-cellXML",)]; root=xml(before); original=canonical(before)
    summary=mutate(root,dict(plan["controls"])); after=E.tostring(root,encoding="utf-8")
    # Reversing the exact lane operation catches accidental sibling/model drift.
    check=xml(after); need(summary.get("mode")=="insert_scalar" or summary.get("label_id"),"No selected label")
    # Restore every pre-existing identified node, then remove newly-created
    # nodes.  This proves that the operation's inverse covers bound text
    # variables as well as the selected label itself.
    before_root = xml(before)
    before_ids = {x.get("id"): x for x in before_root if x.get("id")}
    check_ids = {x.get("id"): x for x in check if x.get("id")}
    for ident, old_node in before_ids.items():
        new_node = check_ids.get(ident)
        if new_node is not None:
            new_node.getparent().replace(new_node, copy.deepcopy(old_node))
    for ident, new_node in list(check_ids.items()):
        if ident not in before_ids:
            check.remove(new_node)
    need(canonical(E.tostring(check,encoding="utf-8"))==original,"Whole-model reversal failed")
    slide_part = None
    slide_after = None
    physical_summary = None
    if plan["controls"].get("physical_field") is not None:
        slide_number = plan["target"].get("slide_number")
        label, _owner, _field = resolve_label(root, plan["controls"])
        derived_tag = label.findtext("m_ppttb/m_bstrShapeName")
        shape_tag = plan["controls"]["physical_field"].get("shape_tag") or derived_tag
        need(isinstance(slide_number, int) and isinstance(shape_tag, str) and shape_tag,
             "physical_field requires a placed label shape tag")
        slide_part = f"ppt/slides/slide{slide_number}.xml"
        rel_part = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            slide = E.fromstring(z.read(slide_part)); rels = E.fromstring(z.read(rel_part))
            shape, _ = tagged_physical_shape(slide, rels, z, shape_tag)
            before_fields = [E.tostring(node, encoding="utf-8")
                             for node in shape.xpath(".//a:fld", namespaces=PPT_NS)]
            physical_summary = apply_physical_field(shape, plan["controls"])
            after_fields = [E.tostring(node, encoding="utf-8")
                            for node in shape.xpath(".//a:fld", namespaces=PPT_NS)]
            need(len(before_fields) == len(after_fields), "physical field count changed")
            selected_index = physical_summary["field_index"]
            need(all(before_fields[i] == after_fields[i] for i in range(len(before_fields))
                     if i != selected_index),
                 "non-target physical label fields changed")
            physical_summary["preserved_field_indices"] = [i for i in range(len(after_fields))
                                                            if i != selected_index]
            need(plan["controls"].get("number_format_key") == physical_summary.get("format_key"),
                 "physical_field format key must equal the bound model number_format_key")
            slide_after = E.tostring(slide, xml_declaration=True, encoding="UTF-8", standalone=True)
    replace_carrier(raw,chart,after,output,slide_part,slide_after)
    if slide_part is not None:
        with zipfile.ZipFile(io.BytesIO(raw)) as zin, zipfile.ZipFile(output) as zout:
            changed_entries = [name for name in zout.namelist() if zin.read(name) != zout.read(name)]
        need(set(changed_entries) == {chart["doc"]["part"], slide_part},
             "physical field preparation changed unexpected package entries: " + str(changed_entries))
    report={"status":"NATIVE_LABEL_CONTROLS_PREPARED_REGEN_REQUIRED","source_sha256":sha(raw),"output_sha256":sha(output.read_bytes()),"summary":summary,"physical_summary":physical_summary,"changed_entries":[chart["doc"]["part"]] + ([slide_part] if slide_part else []),"source_unchanged":sha(path.read_bytes())==sha(raw),"native_gates_remaining":["official regeneration","native reopen/render","changed-data verification"]}
    report_path.write_text(json.dumps(report,indent=2),encoding="utf-8"); return report


if __name__ == "__main__":
    ap=argparse.ArgumentParser(description=__doc__); sub=ap.add_subparsers(dest="cmd",required=True)
    m=sub.add_parser("make-plan");m.add_argument("--input",type=Path,required=True);m.add_argument("--target-json",type=Path,required=True);m.add_argument("--controls-json",type=Path,required=True);m.add_argument("--plan-out",type=Path,required=True);m.add_argument("--experimental-insertion",action="store_true",help="enable the unverified scalar-label insertion mode")
    p=sub.add_parser("prepare");p.add_argument("--input",type=Path,required=True);p.add_argument("--plan",type=Path,required=True);p.add_argument("--output",type=Path,required=True);p.add_argument("--report",type=Path,required=True);p.add_argument("--experimental-insertion",action="store_true",help="enable the unverified scalar-label insertion mode")
    a=ap.parse_args()
    try:
        if a.cmd=="make-plan": a.plan_out.write_text(json.dumps(make_plan(a.input,json.loads(a.target_json.read_text()),json.loads(a.controls_json.read_text()),a.experimental_insertion),indent=2),encoding="utf-8"); print(json.dumps({"status":"PLAN_CREATED"}))
        else: print(json.dumps(prepare(a.input,a.plan,a.output,a.report,a.experimental_insertion),indent=2))
    except Exception as e:
        print(json.dumps({"status":"REJECTED","error":str(e)})); raise
