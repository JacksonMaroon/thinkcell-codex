"""Portable, data-derived preflight for native multi-break ownership.

This adapter does not assume donor IDs or a fixed two-break closure.  It
selects the target by shape tag, resolves every model and physical reference,
derives crossing category groups from each normalized omitted interval, and
records the closure expected after official regeneration.
"""
from __future__ import annotations
import argparse, json, pathlib, sys, zipfile
from lxml import etree as E

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import run_variant_control as variant

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"p": P, "r": R}
Q = lambda n: "{" + R + "}" + n


def need(ok, msg):
    if not ok:
        raise RuntimeError(msg)


def physical_relationships(path: pathlib.Path, names: set[str]):
    """Return each requested tag's resolved relationship target exactly once."""
    found = {}
    with zipfile.ZipFile(path) as z:
        slide = E.fromstring(z.read("ppt/slides/slide1.xml"))
        relroot = E.fromstring(z.read("ppt/slides/_rels/slide1.xml.rels"))
        import posixpath
        rels = {x.get("Id"): posixpath.normpath(posixpath.join("ppt/slides", x.get("Target"))) for x in relroot}
        for tags in slide.findall(".//p:tags", NS):
            rid = tags.get(Q("id")); target = rels.get(rid)
            need(target is not None and target in z.namelist(), "physical tag relationship target is missing")
            values = [x.get("val") for x in E.fromstring(z.read(target)) if (x.get("name") or "").upper() == "THINKCELLSHAPEDONOTDELETE"]
            for name in names.intersection(values):
                need(name not in found, "physical tag is duplicated")
                found[name] = {"relationship_id": rid, "target": target}
    need(set(found) == names, "model physical tags do not all resolve to slide relationships")
    return found


def physical_x_positions(path: pathlib.Path, names: set[str]):
    """Resolve each physical tag to its slide-shape x origin."""
    found = {}
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    with zipfile.ZipFile(path) as z:
        slide = E.fromstring(z.read("ppt/slides/slide1.xml"))
        relroot = E.fromstring(z.read("ppt/slides/_rels/slide1.xml.rels"))
        import posixpath
        rels = {x.get("Id"): posixpath.normpath(posixpath.join("ppt/slides", x.get("Target"))) for x in relroot}
        for tags in slide.findall(".//p:tags", NS):
            rid = tags.get(Q("id")); target = rels.get(rid)
            if target is None or target not in z.namelist():
                continue
            values = {x.get("val") for x in E.fromstring(z.read(target)) if (x.get("name") or "").upper() == "THINKCELLSHAPEDONOTDELETE"}
            hit = names.intersection(values)
            if not hit:
                continue
            shape = tags.xpath("ancestor::*[self::p:sp or self::p:graphicFrame][1]", namespaces=NS)
            off = shape[0].find(f".//{{{A}}}xfrm/{{{A}}}off") if shape else None
            need(off is not None and off.get("x") is not None, "physical tag shape x origin is missing")
            for name in hit:
                need(name not in found, "physical tag has multiple slide shapes")
                found[name] = int(off.get("x"))
    need(set(found) == names, "physical tag x origins do not resolve")
    return found


def crossing_categories(values, low):
    """Positive baseline-zero span rule: any bar endpoint above the gap low."""
    return [j for j in range(len(values[0])) if any(v is not None and v > low for row in values for v in [row[j]])]


def preflight(path: pathlib.Path, plan: dict, phase: str = "seed"):
    data = plan["targets"][0]["data"]
    chart, _ = variant.target_any(path, plan)
    actual = variant.runner.model_of(chart); expected = data["expected_model"]
    for key in ("categories", "series_names", "series_values", "category_extents"):
        need(variant.runner.equal(actual.get(key), expected[key]), "exact model mismatch: " + key)
    sheets = variant.runner.read_sheet(path, 1, chart["owner_name"])["sheets"]
    cells = {(x["row"], x["column"]): x["value"] for x in sheets[0]["nonempty_cells"]}
    wanted = {(r + 1, c + 1): v for r, row in enumerate(data["matrix"]) for c, v in enumerate(row) if v not in (None, "")}
    need(cells == wanted, "exact datasource cells mismatch")
    ids = chart["doc"]["ids"]; axis = ids[chart["owner"].find("m_daxisPrimaryValue").get("idref")]
    refs = axis.findall("m_cdaxisbreak/elem"); need(refs, "native breaks are missing")
    ordered = []
    all_names = []
    for ref in refs:
        need(ref.get("idref") in ids, "break owner reference is unresolved")
        br = ids[ref.get("idref")]; ordinal = br.find("m_nOrdinal")
        need(br.tag == "CDataAxisBreak" and ordinal is not None and br.find("m_bUser") is not None and br.find("m_bUser").get("val") == "1", "break owner fields are incomplete")
        shapes = []
        shape_tag_groups = []
        for sr in br.findall("m_cpptbreakshp/elem"):
            need(sr.get("idref") in ids, "break shape reference is unresolved")
            shape = ids[sr.get("idref")]; need(shape.tag == "CPPTBreakShape", "break shape type changed")
            names = []
            for field in variant.runner.FIELDS:
                node = shape.find(field); tag = node.find("m_bstrShapeName") if node is not None else None
                need(tag is not None and tag.text, "break shape physical tag is missing"); names.append(tag.text)
            need(len(set(names)) == len(names), "break shape tags are duplicated"); all_names.extend(names); shapes.append(shape); shape_tag_groups.append(names)
        ordered.append({"ordinal": int(ordinal.get("val")), "fraction": float(br.find("m_fFraction").get("val")), "shape_count": len(shapes), "shape_ids": [x.get("id") for x in shapes], "physical_tag_count": len(shapes) * 3, "shape_tag_groups": shape_tag_groups})
    need(sorted(x["ordinal"] for x in ordered) == list(range(len(ordered))), "break ordinals are not contiguous")
    need(len(all_names) == len(set(all_names)), "break physical tags are not globally unique")
    relationships = physical_relationships(path, set(all_names))
    x_positions = physical_x_positions(path, set(all_names))
    category_x = sorted(set(x_positions.values()))
    need(category_x, "break physical category positions are missing")
    vector = axis.findall("m_vecintvlOrdinal"); need(len(vector) == 1, "ordered interval vector is missing")
    intervals = [[float(e.find("begin").get("val")), float(e.find("end").get("val"))] for e in vector[0].findall("elem")]
    requested = [x for x in sorted(data["expected_breaks"], key=lambda x: x["ordinal"])]
    expected_gaps = data.get("expected_generated_gaps") if phase == "native" else [x["gap"] for x in requested]
    expected_counts = data.get("expected_generated_shape_counts") if phase == "native" else [1] * len(requested)
    need(len(expected_gaps) == len(ordered) and len(expected_counts) == len(ordered), "expected closure does not cover every break")
    normalized = []
    values = actual["series_values"]
    for ordinal, gap in enumerate(expected_gaps):
        need(ordinal < len(intervals), "break ordinal exceeds interval vector")
        low, high = intervals[ordinal]
        need(abs(low - gap[0]) <= 1e-8 and abs(high - gap[1]) <= 1e-8, "expected normalized omitted interval is missing at its ordinal slot")
        cats = crossing_categories(values, low)
        shape_categories = []
        for shape_names in ordered[ordinal]["shape_tag_groups"]:
            xs = [x_positions[name] for name in shape_names]
            need(len(set(xs)) == 1, "break shape physical tags do not share a category origin")
            shape_categories.append(category_x.index(xs[0]))
        unique_category_ownership = sorted(shape_categories) == sorted(set(shape_categories))
        if phase == "native":
            need(unique_category_ownership, "native break shapes duplicate category ownership")
            need(sorted(shape_categories) == sorted(cats), "native break shape ownership differs from data-derived crossings")
        normalized.append({"ordinal": ordinal, "gap": [low, high], "crossing_categories": cats, "derived_shape_count": len(cats), "actual_shape_count": ordered[ordinal]["shape_count"], "shape_category_indices": sorted(shape_categories)})
    actual_counts = [x["actual_shape_count"] for x in normalized]
    need(actual_counts == list(expected_counts), "native shape closure differs from expected counts")
    # The cloned seed deliberately has one physical shape for each break.  A
    # seed is therefore structurally valid even when its category closure is
    # incomplete; official native regeneration must expand it to the data
    # derived crossing count.
    closure_matches_derived = [x["actual_shape_count"] == x["derived_shape_count"] for x in normalized]
    if phase == "native":
        need(all(closure_matches_derived), "native shape closure does not match data-derived crossings")
        status = "PORTABLE_MULTI_BREAK_PREFLIGHT_PASS"
    else:
        status = "PORTABLE_MULTI_BREAK_SEED_PREFLIGHT_PASS"
    return {"status": status, "phase": phase, "breaks": sorted(ordered, key=lambda x: x["ordinal"]), "normalized_intervals": normalized, "physical_relationships": relationships, "physical_tag_count": len(all_names), "interval_count": len(intervals), "model_datasheet_exact": True, "closure_matches_derived": closure_matches_derived}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(); parser.add_argument("--input", type=pathlib.Path, required=True); parser.add_argument("--plan", type=pathlib.Path, required=True); parser.add_argument("--phase", choices=("seed", "native"), default="seed"); parser.add_argument("--output", type=pathlib.Path)
    args = parser.parse_args(); result = preflight(args.input.resolve(), json.loads(args.plan.read_text(encoding="utf-8")), args.phase); text = json.dumps(result, indent=2)
    if args.output: args.output.write_text(text, encoding="utf-8")
    print(text)
