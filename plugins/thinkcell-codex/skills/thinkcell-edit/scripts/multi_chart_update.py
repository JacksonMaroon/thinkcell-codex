"""Experimental multi-chart runner with official JSON and native verification.

Every chart on the one-slide copy must be explicitly targeted. Supports internal
ordinary sequence charts and at most one fixed-topology waterfall. Ordinary
sequence category/series counts can change with opt-in and a verified simple
datasheet, including its optional 100%= row. Waterfall slot counts, equals cells,
connectors and grounds stay fixed. External links, Mekko and other families need
their separate contracts. Outputs are released only after data, native reopen,
package and applicable cache checks; visual review remains required.
"""
from __future__ import annotations
import argparse, hashlib, io, json, math, shutil, subprocess, sys, zipfile
from pathlib import Path
from lxml import etree as E

PLUGIN = Path(__file__).resolve().parent
if not (PLUGIN / "thinkcell_no_click" / "implementation").is_dir():
    raise RuntimeError("Install this file in the private think-cell scripts directory.")
IMPL = PLUGIN / "thinkcell_no_click" / "implementation"
sys.path[:0] = [str(IMPL), str(PLUGIN)]

from office_operation_lock import serialized_office, run_locked_subprocess
from prepare_thinkcell_name import (choose, inventory, link_contract, logical_slides,
                                    need, prepare, sha, xml)
from update_thinkcell_json import (equal, model_of, payload, read_blob,
                                   wrap_biff_workbook_stream)
from chart_semantics import kind, details, identity_invariants
from runtime import find_ppttc, powershell, powershell_env
from audit_thinkcell_integrity import inspect_presentation
from chart_semantics import audit_scope
from read_named_datasheet import read as read_datasheet
from datasheet_colors import decorate_table, validate as validate_appearance
from series_fill_matching import cache_value_vectors, match_unique_series


def load(path): return json.loads(Path(path).read_text(encoding="utf-8-sig"))
def dump(path, value): Path(path).write_text(json.dumps(value, indent=2), encoding="utf-8")
def labels_equal(left, right):
    """Datasheets display numeric category labels as strings; model readers do not."""
    def one(a,b):
        if isinstance(a,(int,float)) and not isinstance(a,bool) and isinstance(b,(int,float)) and not isinstance(b,bool): return equal(a,b)
        try:
            if isinstance(a,str) and isinstance(b,(int,float)) and not isinstance(b,bool): return equal(float(a),b)
            if isinstance(b,str) and isinstance(a,(int,float)) and not isinstance(a,bool): return equal(a,float(b))
        except ValueError: pass
        return a==b
    return len(left)==len(right) and all(one(a,b) for a,b in zip(left,right))

def cache_semantics(path, candidate):
    """Read the visible native chart subtype without trusting stale model enums."""
    part=candidate["frames"][0]["native_chart_part"]
    need(part, "Target has no native chart cache part")
    with zipfile.ZipFile(path) as z: root=E.fromstring(z.read(part))
    ns={"c":"http://schemas.openxmlformats.org/drawingml/2006/chart"}
    bar=root.find(".//c:barChart",ns); need(bar is not None, "Expected native bar/column cache")
    def val(node): return None if node is None else node.get("val")
    return {"bar_dir":val(bar.find("c:barDir",ns)),"grouping":val(bar.find("c:grouping",ns)),
            "vary_colors":val(bar.find("c:varyColors",ns)),"series_count":len(bar.findall("c:ser",ns))}

def cache_semantics_or_none(path, candidate):
    """Keep legacy non-bar data updates out of the appearance-only mapper."""
    try:
        return cache_semantics(path, candidate)
    except ValueError as error:
        if str(error) != "Expected native bar/column cache":
            raise
        return None

def _cache_root(path, candidate):
    part=candidate["frames"][0]["native_chart_part"]
    with zipfile.ZipFile(path) as z: return E.fromstring(z.read(part))

def _ordinary_fill_route(root, candidate):
    """Prove a regular bar/column cache's series order from its values.

    think-cell's 100%-column donor has a separately certified reverse order.
    All other charts use this stricter route: a single ordinary bar cache, one
    unbroken primary value axis, and a one-to-one full-vector match between
    cache series and the private model.  There is deliberately no ordinal
    fallback for duplicate values, reordering, or a cache we cannot classify.
    """
    ns={"c":"http://schemas.openxmlformats.org/drawingml/2006/chart"}
    bars=root.findall(".//c:barChart",ns)
    need(len(bars)==1, "Appearance fills require exactly one native bar/column cache")
    bar=bars[0]
    def val(node): return None if node is None else node.get("val")
    need(val(bar.find("c:barDir",ns)) in {"bar","col"}, "Appearance fills require a native bar/column cache")
    need(val(bar.find("c:grouping",ns)) in {"stacked","clustered"}, "Appearance fills require stacked or clustered bars/columns")
    need(len(root.findall(".//c:valAx",ns))==1, "Appearance fills require one ordinary value axis")
    owner=candidate["owner"]
    primary=owner.findall("m_daxisPrimaryValue")
    secondary=owner.findall("m_daxisSecondaryValue")
    need(len(primary)==1 and primary[0].get("idref") not in {None,"0"}, "Appearance fills require one primary value axis")
    need(len(secondary)==1 and secondary[0].get("idref")=="0", "Appearance fills exclude secondary axes")
    axis=candidate["doc"]["ids"].get(primary[0].get("idref"))
    need(axis is not None and axis.tag=="CSequenceChartDataAxis" and not axis.findall("m_cdaxisbreak/elem"), "Appearance fills exclude axis breaks")
    model=model_of(candidate)
    return match_unique_series(cache_value_vectors(root), model["series_names"], model["series_values"])

def _fill_series_names(path, candidate, contract=None):
    root=_cache_root(path,candidate)
    visible=cache_semantics(path,candidate)
    # This topology was independently native-certified.  It has no category
    # vector identity because every series is expressed as the 100% scale.
    contract=contract or {}
    model=model_of(candidate)
    if contract.get("only_optional_row")=="100%=" and visible["bar_dir"]=="col" and visible["grouping"]=="stacked":
        return root, list(reversed(model["series_names"])), "percent_reverse"
    return root, _ordinary_fill_route(root,candidate), "unique_vectors"

def cache_fills(path, candidate, contract=None):
    """Return exact visible RGB fill values keyed by native series name.

    This is a verification readback, never a cache write.  A fill plan is
    accepted only when the official regeneration produced this same mapping.
    """
    root,names,_=_fill_series_names(path,candidate,contract)
    ns={"c":"http://schemas.openxmlformats.org/drawingml/2006/chart","a":"http://schemas.openxmlformats.org/drawingml/2006/main"}
    result={}
    series_nodes=root.findall(".//c:barChart/c:ser",ns)
    # think-cell's visible 100%-column cache omits c:tx and emits series in
    # reverse model order.  Tie this explicit cache order back to the native
    # model names rather than relying on ordinal names supplied by a caller.
    need(len(series_nodes)==len(names), "Visible series count differs from native model")
    for name,series in zip(names,series_nodes):
        fills=series.xpath("./c:spPr/a:solidFill/a:srgbClr/@val",namespaces=ns)
        need(len(fills)==1, "Visible series has no exact RGB fill")
        points={int(x.find("c:idx",ns).get("val")): color.upper() for x in series.findall("c:dPt",ns)
                if x.find("c:idx",ns) is not None
                for color in x.xpath("./c:spPr/a:solidFill/a:srgbClr/@val",namespaces=ns)}
        result[name]={"series":fills[0].upper(),"points":points}
    return result

def verify_appearance(path, candidate, contract, appearance):
    actual=cache_fills(path,candidate,contract)
    for name,color in appearance["series_fills"].items():
        need(actual[name]["series"]==color.removeprefix("#").upper(), "Generated series RGB fill differs: "+name)
    for name,colors in appearance["point_fills"].items():
        for index,color in enumerate(colors):
            if color is not None:
                need(actual[name]["points"].get(index)==color.removeprefix("#").upper(), "Generated point RGB fill differs: "+name+"["+str(index)+"]")
    return actual

def datasheet_fill_enabled(candidate):
    node=candidate["table"].find("m_bExcelOnTop")
    return node is not None and node.get("val")=="1"

def _color_fingerprint_root(root):
    """Conservative native-cache style fingerprint for data-only legacy routes."""
    payload=b"\0".join(E.tostring(node, method="c14n") for node in root.iter()
                       if E.QName(node).localname == "spPr" and
                       any(E.QName(descendant).localname in {"srgbClr", "schemeClr", "sysClr", "prstClr", "hslClr"}
                           for descendant in node.iter()))
    return hashlib.sha256(payload).hexdigest().upper()

def raw_color_fingerprint(path, candidate):
    return _color_fingerprint_root(_cache_root(path, candidate))

def _category_index(source_categories, target_categories):
    """Map preserved point overrides by category identity, never ordinal."""
    need(len(source_categories)==len(set(source_categories)), "Cannot preserve point fills: source categories are ambiguous")
    need(len(target_categories)==len(set(target_categories)), "Cannot preserve point fills: requested categories are ambiguous")
    need(set(source_categories)==set(target_categories), "Cannot preserve point fills when category identities change")
    return {category:index for index,category in enumerate(target_categories)}

def merge_enabled_appearance(current, source_categories, request, requested):
    """Merge a partial explicit request with enabled-donor fills.

    Point overrides belong to categories rather than their current ordinal.  A
    whole-series request intentionally clears old point overrides for that
    series, then applies only the explicitly requested point colors.
    """
    target_categories=request["expected_model"]["categories"]
    names=request["expected_model"]["series_names"]
    requested_series=requested["series_fills"]
    requested_points=requested["point_fills"]
    source_index=None
    if any(value["points"] for value in current.values()):
        source_index=_category_index(source_categories,target_categories)
    series={name:"#"+current[name]["series"] for name in names if name in current}
    series.update(requested_series)
    points={}
    for name in names:
        prior=current.get(name, {"points":{}})["points"]
        if name in requested_series:
            merged=[None]*len(target_categories)
        else:
            merged=[None]*len(target_categories)
            if prior:
                for old_index,color in prior.items():
                    need(isinstance(old_index,int) and 0 <= old_index < len(source_categories), "Existing point fill index is invalid")
                    merged[source_index[source_categories[old_index]]]= "#"+color
        for index,color in enumerate(requested_points.get(name, [])):
            if color is not None:
                merged[index]=color
        if any(color is not None for color in merged):
            points[name]=merged
    return {"schema":"tc.datasheet-fill.v1","series_fills":series,"point_fills":points}

def effective_appearance(prepared, candidate, spec, contract):
    """Use explicit requests, or preserve an enabled donor's existing fills.

    When the donor is configured to consume datasheet fills, omitting fill
    attributes would otherwise silently return it to the fallback palette.
    Existing named series therefore receive their current exact RGB fills;
    newly added series retain normal donor fallback and are reported later.
    """
    if "appearance" in spec:
        if not datasheet_fill_enabled(candidate):
            return spec["appearance"], "requested"
        current=cache_fills(prepared,candidate,contract)
        merged=merge_enabled_appearance(current, model_of(candidate)["categories"], spec["data"], spec["appearance"])
        validate_appearance(merged,spec["data"],contract)
        return merged, "requested_merged"
    if not datasheet_fill_enabled(candidate): return None, "not_enabled"
    # `_fill_series_names` checks either the certified 100% donor path or the
    # ordinary unique-vector path before we copy any colors into JSON cells.
    current=cache_fills(prepared,candidate,contract)
    appearance=merge_enabled_appearance(
        current, model_of(candidate)["categories"], spec["data"],
        {"schema":"tc.datasheet-fill.v1","series_fills":{},"point_fills":{}},
    )
    validate_appearance(appearance,spec["data"],contract)
    return appearance, "preserved"

def owner_semantics(candidate):
    """Stable private-model chart grammar; excludes data/name/cache references."""
    ignored={"m_dtable","m_strName","m_pptseqchart"}
    # Official JSON legitimately regenerates object IDs for plot/axis anchors.
    # Their fields and presence remain semantic; their numeric idref values are not.
    return [(node.tag,tuple(sorted((k,"<id>") if k=="idref" else (k,v) for k,v in node.attrib.items())),node.text or "") for node in candidate["owner"] if node.tag not in ignored]

def cells_matrix(c):
    """Return canonical (series rows/category columns) source matrix and orientation."""
    ds = link_contract(c); streams = c["doc"]["streams"]; store = ds["storage"]
    meta = streams.get((store, "think-cellXML")); need(meta is not None, "Datasource orientation metadata missing")
    orient = xml(meta).find("PersistentType/m_eorient")
    need(orient is not None and orient.get("val") in {"0", "1"}, "Unknown datasource orientation")
    if (store, "Package") in streams: sheets = read_blob(streams[(store, "Package")], "xlsb_package")
    elif (store, "Workbook") in streams: sheets = read_blob(wrap_biff_workbook_stream(streams[(store, "Workbook")]), "legacy_biff_cfb")
    else: raise ValueError("Unsupported embedded datasource storage")
    need(len(sheets) == 1 and sheets[0]["nonempty_cells"], "Expected one nonempty embedded datasource")
    cells = sheets[0]["nonempty_cells"]; rows=max(x["row"] for x in cells); cols=max(x["column"] for x in cells)
    raw=[[None]*cols for _ in range(rows)]
    for x in cells: raw[x["row"]-1][x["column"]-1]=x["value"]
    canonical=[list(x) for x in zip(*raw)] if orient.get("val") == "1" else raw
    return canonical, orient.get("val"), ds

def simple_sequence_contract(c, request, flexible):
    """Validate the only layout safe for an experimental count-change canary."""
    need(c["owner"].tag == "CSequenceChartSE" and kind(c) == "CSequenceChartSE", "This experiment excludes waterfall and Mekko")
    baseline=model_of(c); source, orientation, ds=cells_matrix(c); expected=request["expected_model"]; matrix=request["matrix"]
    required={"categories", "series_names", "series_values"}
    optional={"category_extents"}
    need(set(expected) in (required, required|optional), "Expected model fields are invalid")
    need(isinstance(matrix,list) and len(matrix)>=2 and all(isinstance(r,list) for r in matrix), "Matrix must have header plus series")
    width=len(matrix[0]); need(width>=2 and all(len(r)==width for r in matrix), "Matrix must be rectangular with one or more categories")
    cats=expected["categories"]; names=expected["series_names"]; values=expected["series_values"]
    need(all(isinstance(x,(str,int,float)) and not isinstance(x,bool) for x in cats) and all(isinstance(x,str) for x in names) and len(cats)==width-1 and len(set(names))==len(names), "Categories/series names invalid")
    need(len(names)>0 and len(values)==len(names) and all(isinstance(r,list) and len(r)==len(cats) for r in values), "Series dimensions invalid")
    need(all(v is None or (isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(v)) for r in values for v in r), "Series values must be finite numbers/null")
    need(labels_equal(matrix[0][1:],cats), "Matrix category row differs from expected model")
    need(labels_equal(source[0][1:],baseline["categories"]), "Source category row disagrees with model")
    # Explicit special 100%= row is the only permitted non-series row.  Detect it
    # from the source rather than relying on its label, which varies by donor.
    start=1; special=None
    if len(source)>1 and source[1][0] not in baseline["series_names"] and all(isinstance(v,(int,float)) and not isinstance(v,bool) for v in source[1][1:]):
        special=source[1]; start=2
        ext=expected.get("category_extents")
        need(isinstance(ext,list) and len(ext)==len(cats), "100%= donor requires category_extents")
        need(matrix[1][0]==special[0] and equal(matrix[1][1:],ext), "100%= row must be explicit and match category_extents")
        need(all(isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(v) and v>0 for v in ext), "category_extents must be positive")
    elif len(source)>1 and source[1][0] not in baseline["series_names"] and all(v is None for v in source[1][1:]) and "category_extents" in baseline:
        # think-cell's 100%-axis marker is a deliberately blank reserved row in
        # the vendor paired-column donor.  The model derives its extents from
        # the data, so preserve blanks and independently require the declared
        # extents to equal the new column totals.
        special=source[1]; start=2; ext=expected.get("category_extents")
        need(isinstance(ext,list) and len(ext)==len(cats), "Blank 100%= donor requires category_extents")
        need(matrix[1][0]==special[0], "Blank 100%= row label changed")
        totals=[sum((r[i] or 0) for r in values) for i in range(len(cats))]
        supplied=matrix[1][1:]
        if all(v is None for v in supplied):
            need(equal(ext,totals), "Blank 100%= category_extents must equal series totals")
        else:
            need(all(isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(v) and v>0 for v in supplied), "100%= denominators must be positive finite numbers")
            need(equal(ext,supplied), "100%= denominators must match category_extents")
    else:
        need("category_extents" not in expected, "category_extents supplied for a donor without 100%= row")
    source_names=[r[0] for r in source[start:]]
    need(all(isinstance(n,str) and n for n in source_names) and len(set(source_names))==len(source_names), "Source has an unsupported optional/reserved datasource row")
    need(len(matrix)==start+len(names), "Matrix has unsupported optional/reserved rows")
    for i,name in enumerate(names,start):
        need(matrix[i][0]==name and equal(matrix[i][1:],values[i-start]), "Matrix series row disagrees with expected model")
    if not flexible:
        need(len(cats)==len(baseline["categories"]) and len(names)==len(baseline["series_names"]), "Set allow_sequence_count_change for a count-changing canary")
    return {"family":"sequence", "datasheet_orientation":orientation, "source_dimensions":[len(source),len(source[0])],
            "requested_dimensions":[len(matrix),len(matrix[0])], "count_change": [len(source),len(source[0])] != [len(matrix),len(matrix[0])],
            "only_optional_row":"100%=" if special else None, "datasource":ds}

def target_from_spec(candidates, spec):
    s=spec.get("selector", {}); need(set(s) <= {"slide_id","slide_number","shape_id","shape_tag"}, "Invalid selector keys")
    return choose(candidates, s.get("slide_id"), s.get("slide_number"), s.get("shape_id"), s.get("shape_tag"))

def waterfall_contract(c, request):
    """Fixed slots, equals cells and optional rows; data uses official JSON only."""
    base=model_of(c); source,orientation,ds=cells_matrix(c)
    matrix=request['matrix']; exp=request['expected_model']
    need(set(exp)=={'categories','series_names','series_values'}, 'Waterfall expected model fields invalid')
    need(len(matrix)==len(source) and all(isinstance(r,list) and len(r)==len(source[0]) for r in matrix), 'Waterfall slot dimensions must remain fixed')
    need(len(exp['categories'])==len(base['categories']) and labels_equal(matrix[0][1:],exp['categories']), 'Waterfall categories invalid')
    need(len(exp['series_names'])==len(base['series_names']) and len(set(exp['series_names']))==len(exp['series_names']) and len(exp['series_values'])==len(base['series_values']), 'Waterfall series dimensions invalid')
    seen=[]
    for index,row in enumerate(source[1:],1):
        if row[0] not in base['series_names']:
            need(equal(row,matrix[index]), 'Waterfall reserved row changed'); continue
        si=base['series_names'].index(row[0]); seen.append(si); values=exp['series_values'][si]
        need(len(values)==len(exp['categories']) and matrix[index][0]==exp['series_names'][si], 'Waterfall series row invalid')
        for old,new,wanted in zip(row[1:],matrix[index][1:],values):
            need((old=='e')==(new=='e'), 'Waterfall equals slots changed')
            need(wanted is None or (not isinstance(wanted,bool) and isinstance(wanted,(int,float)) and math.isfinite(wanted)), 'Waterfall expected values invalid')
            need((new=='e' and isinstance(wanted,(int,float))) or equal(new,wanted), 'Waterfall cells differ from expected values')
    need(sorted(seen)==list(range(len(base['series_names']))), 'Waterfall source series unresolved')
    return {'family':'waterfall','semantics':details(c),'fixed_slot_dimensions':[len(source),len(source[0])],'datasource':ds}

def validate_plan(data, plan):
    need(isinstance(plan,dict) and set(plan) <= {"targets","allow_sequence_count_change"}, "Invalid plan keys")
    targets=plan.get("targets"); need(isinstance(targets,list) and targets, "Plan requires targets")
    docs, candidates, _ = inventory(data)
    identity_invariants(docs)
    need(len(logical_slides(zipfile.ZipFile(io.BytesIO(data))))==1, "Canary input must be one slide")
    chosen=[]; names=[]
    for spec in targets:
        need(isinstance(spec,dict) and set(spec) in ({"selector","name","data"},{"selector","name","data","appearance"}), "Each target needs selector, name, data, and optional appearance")
        c=target_from_spec(candidates,spec); link_contract(c); chosen.append((c,spec)); names.append(spec["name"])
        need(isinstance(spec["name"],str) and spec["name"].isascii() and spec["name"].strip(), "Name must be printable ASCII")
    need(len(set(x.casefold() for x in names))==len(names), "Automation names must be unique")
    need(len({x[0]["frames"][0]["shape_tag"] for x in chosen})==len(chosen), "Duplicate target chart")
    need(all(c["owner"].tag=="CSequenceChartSE" and kind(c) in {"CSequenceChartSE","waterfall"} and c["exact"] for c in candidates), "All slide charts must be exact sequence or fixed-topology waterfall elements")
    need(sum(kind(c)=='waterfall' for c in candidates)<=1, 'Mixed route currently supports at most one waterfall')
    need(len(chosen)==len(candidates), "Plan must update every native chart on the slide, preventing silent sibling changes")
    flex=bool(plan.get("allow_sequence_count_change",False))
    contracts=[waterfall_contract(c,s['data']) if kind(c)=='waterfall' else simple_sequence_contract(c,s["data"],flex) for c,s in chosen]
    for (_, spec), contract in zip(chosen,contracts):
        if "appearance" in spec:
            need(contract["family"]=="sequence", "Appearance fills support ordinary sequence charts only")
            validate_appearance(spec["appearance"], spec["data"], contract)
    return chosen, contracts

def prepare_all(src, digest, stage, chosen):
    """Apply name-only copies sequentially; every subsequent hash is read back."""
    current=src; manifests=[]
    for index,(old,spec) in enumerate(chosen):
        docs,candidates,_=inventory(current.read_bytes())
        # Select by stable tag, retained by copy-only naming.
        tag=old["frames"][0]["shape_tag"]; now=choose(candidates, slide_number=1, shape_tag=tag)
        nxt=stage/f"named-{index+1}.pptx"
        result=prepare(current, sha(current.read_bytes()), nxt, slide_number=1, shape_tag=tag, name=spec["name"], execute=True)
        need(result["only_name_fields_changed"], "Name preparation changed non-name fields")
        manifests.append(result); current=nxt
    return current, manifests

def validate_generated(prepared, generated, plan, effective_appearances=None):
    before=prepared.read_bytes(); after=generated.read_bytes(); selected, contracts=validate_plan(before,plan)
    docs, candidates, _=inventory(after); identity_invariants(docs); need(len({c["owner_name"] for c in candidates})==len(candidates), "Generated automation names are not unique"); by_name={c["owner_name"]:c for c in candidates}
    need(set(by_name)=={s["name"] for _,s in selected}, "Generated chart names differ from explicit target set")
    need(len(candidates)==len(selected), "Generated native chart count differs from explicit target count")
    results=[]
    for (old,spec),contract in zip(selected,contracts):
        c=by_name[spec["name"]]; need(c["frames"][0]["shape_tag"]==old["frames"][0]["shape_tag"], "Target retargeted")
        actual=model_of(c); exp=spec["data"]["expected_model"]
        need(labels_equal(actual["categories"],exp["categories"]), "Generated categories differ")
        actual_series=dict(zip(actual["series_names"],actual["series_values"])); expected_series=dict(zip(exp["series_names"],exp["series_values"]))
        need(len(actual_series)==len(actual["series_names"]) and equal(actual_series,expected_series), "Generated series differ")
        if "category_extents" in exp: need(equal(actual.get("category_extents"),exp["category_extents"]), "Generated category extents differ")
        need(kind(c)==kind(old), "Chart kind changed")
        if kind(c)=='waterfall':need(details(c)==contract['semantics'], 'Waterfall equals slots, connectors or grounds changed')
        need(owner_semantics(old)==owner_semantics(c), "Native model chart grammar changed")
        before_cache=cache_semantics_or_none(prepared,old); after_cache=cache_semantics_or_none(generated,c)
        need((before_cache is None)==(after_cache is None), "Native bar/column cache availability changed")
        base_subtype={k:before_cache[k] for k in ("bar_dir","grouping","vary_colors")} if before_cache else None
        after_subtype={k:after_cache[k] for k in ("bar_dir","grouping","vary_colors")} if after_cache else None
        if before_cache:
            degenerate_clustered=(base_subtype["bar_dir"]=="col" and base_subtype["grouping"]=="clustered" and after_subtype=={"bar_dir":"col","grouping":"stacked","vary_colors":base_subtype["vary_colors"]} and after_cache["series_count"]==1 and len(exp["series_names"])==1)
            restored_clustered=(before_cache["series_count"]==1 and base_subtype["bar_dir"]=="col" and base_subtype["grouping"]=="stacked" and after_subtype=={"bar_dir":"col","grouping":"clustered","vary_colors":base_subtype["vary_colors"]} and after_cache["series_count"]>1)
            need(base_subtype==after_subtype or degenerate_clustered or restored_clustered, "Native cache subtype/direction changed")
        else:
            degenerate_clustered=restored_clustered=False
        sheet=read_datasheet(generated,1,spec["name"])["sheets"]
        need(len(sheet)==1, "Expected one generated datasheet")
        actual_cells={(x["row"],x["column"]):x["value"] for x in sheet[0]["nonempty_cells"]}
        # JSON matrices are canonical sequence orientation, while embedded
        # datasheets may be transposed. Normalize after generation from its
        # own metadata, never from the source donor.
        canonical, orientation, _ = cells_matrix(c)
        if orientation=="1": actual_cells={(col,row):v for (row,col),v in actual_cells.items()}
        expected_cells={(r+1,col+1):v for r,row in enumerate(spec["data"]["matrix"]) for col,v in enumerate(row) if v not in (None,"")}
        need(set(actual_cells)==set(expected_cells) and all(equal(actual_cells[k],v) for k,v in expected_cells.items()), "Generated embedded datasheet cells differ")
        requested=(effective_appearances or {}).get(spec["name"], spec.get("appearance"))
        if requested is not None:
            colors=verify_appearance(generated,c,contract,requested)
            preserved=None
        elif datasheet_fill_enabled(old):
            before_colors=cache_fills(prepared,old,contract); colors=cache_fills(generated,c,contract)
            shared=set(before_colors)&set(colors)
            need(all(before_colors[name]==colors[name] for name in shared), "Existing series/point RGB fill changed without an appearance request")
            preserved=sorted(shared)
        else:
            if contract.get("count_change"):
                # Count changes can legitimately add/remove cached style nodes.
                # The legacy data path remains governed by its existing model,
                # datasheet, subtype and integrity checks.
                before_colors=colors={"mode":"legacy_unmapped_count_change"}
            else:
                before_colors=raw_color_fingerprint(prepared,old); colors=raw_color_fingerprint(generated,c)
                need(before_colors==colors, "Native color formatting changed without an appearance request")
            preserved=None
        results.append({"name":spec["name"],"shape_tag":c["frames"][0]["shape_tag"],"contract":contract,"exact_datasheet_cells":True,"category_extents":actual.get("category_extents"),"native_cache_semantics":after_cache,"appearance_fills":colors,"preserved_existing_fill_series":preserved,"degenerate_clustered_to_stacked":degenerate_clustered})
    before_brands=[]; after_brands=[]
    from update_thinkcell_json import snapshot
    for _,spec in selected:
        before_brands.append(snapshot(prepared,spec["name"])["branding"]); after_brands.append(snapshot(generated,spec["name"])["branding"])
    need(before_brands==after_brands, "Theme or notes changed")
    audit=inspect_presentation(generated,True)
    if any(kind(c)=='waterfall' for c in candidates):
        from mixed_integrity import mixed_integrity_scope
        integrity=mixed_integrity_scope(generated,audit)
    else:integrity=audit_scope(generated,audit)
    need(prepared.read_bytes()==before and generated.read_bytes()==after, "Input/output changed during validation")
    return {"targets":results,"theme_and_notes_preserved":True,"integrity_scope":integrity,
            "native_chart_count":len(candidates),"native_tags":sorted(c["frames"][0]["shape_tag"] for c in candidates),"unique_names":sorted(by_name)}

@serialized_office
def run(args):
    src=args.input.resolve(); out=args.output.resolve(); plan=load(args.plan); digest=sha(src.read_bytes())
    need(digest==args.expected_sha256.upper(), "Input hash mismatch"); need(src!=out and out.suffix.lower()==".pptx", "Output must be a distinct PPTX")
    chosen,contracts=validate_plan(src.read_bytes(),plan)
    report={"status":"DRY_RUN_PASS", "source":str(src), "source_sha256":digest, "targets":[{"name":s["name"],"shape_tag":c["frames"][0]["shape_tag"]} for c,s in chosen], "contracts":contracts,
            "native_reopen_render_required":True, "no_external_links":True, "all_native_charts_explicitly_targeted":True}
    if not args.prepare and not args.execute:return report
    need(not out.exists(), "Output already exists"); need(args.report.resolve() not in {src,out,args.plan.resolve()}, "Report overlaps an input or output"); stage=out.parent/(out.stem+"_multi_tc_work"); need(not stage.exists(), "Staging directory exists")
    stage.mkdir(); prepared,manifests=prepare_all(src,digest,stage,chosen)
    report.update(status="PREPARED_FOR_OFFICIAL_JSON", prepared=str(prepared), prepared_sha256=sha(prepared.read_bytes()), name_preparations=manifests, staging_directory=str(stage), output_withheld=True)
    if not args.execute:return report
    ppttc=Path(args.ppttc) if args.ppttc else find_ppttc(); need(ppttc and ppttc.is_file(), "Official ppttc.exe not found")
    # Resolve contracts after naming; names are the only permitted preparation delta.
    selected,_=validate_plan(prepared.read_bytes(),plan)
    job=stage/"multi-update.ppttc"; generated=stage/"generated.pptx"
    effective=[effective_appearance(prepared,c,s,ct) for (c,s),ct in zip(selected,contracts)]
    job.write_text(json.dumps([{ "template":str(prepared), "data":[{"name":s["name"],"table":payload(decorate_table(s["data"],ct,appearance) if appearance else s["data"]["matrix"],ct)} for ((c,s),ct,(appearance,mode)) in zip(selected,contracts,effective)]}]),encoding="utf-8")
    report["appearance_modes"]={s["name"]:mode for (_,s),(_,mode) in zip(selected,effective)}
    with (stage/"generator.stdout.txt").open("w") as stdout,(stage/"generator.stderr.txt").open("w") as stderr:
        code=run_locked_subprocess([str(ppttc),str(job),'-o',str(generated)],operation='ppttc-multi-json',timeout_seconds=240,stdout=stdout,stderr=stderr,creationflags=subprocess.CREATE_NO_WINDOW).returncode
    need(code==0 and generated.exists(), "Official JSON generator failed; inspect staging logs")
    effective_by_name={s["name"]:appearance for (_,s), (appearance,_) in zip(selected,effective)}
    results=validate_generated(prepared,generated,plan,effective_by_name)
    native=stage/"native-reopened.pptx"; native_report=stage/"native-report.json"; render=stage/"native.png"
    verify=IMPL/"native_verify_scoped.ps1"
    with (stage/"native.stdout.txt").open("w") as stdout,(stage/"native.stderr.txt").open("w") as stderr:
        code2=run_locked_subprocess([powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(verify),'-InputFile',str(generated),'-OutputFile',str(native),'-ReportFile',str(native_report),'-RenderFile',str(render)],operation='native-multi-verify',timeout_seconds=240,stdout=stdout,stderr=stderr,env=powershell_env(),creationflags=subprocess.CREATE_NO_WINDOW).returncode
    need(code2==0 and native.exists() and native_report.exists() and render.exists(), "Native reopen/render gate failed; inspect staging logs")
    native_state=load(native_report); need(native_state.get("native_reopen_pass") and native_state.get("other_presentations_unchanged") and native_state.get("source_unchanged"), "Native verification report failed")
    final_validation=validate_generated(prepared,native,plan,effective_by_name)
    need(sha(src.read_bytes())==digest, "Source changed during run")
    with out.open("xb") as f: f.write(native.read_bytes())
    report.update(status="ALL_GATES_PASS", output=str(out), output_sha256=sha(out.read_bytes()), validation=final_validation, generator_exit_code=code, native_verification=native_state, render=str(render), source_unchanged=True, output_withheld=False)
    return report

def main():
    p=argparse.ArgumentParser(description=__doc__); p.add_argument("--input",type=Path,required=True); p.add_argument("--expected-sha256",required=True); p.add_argument("--plan",type=Path,required=True); p.add_argument("--output",type=Path,required=True); p.add_argument("--report",type=Path,required=True); p.add_argument("--prepare",action="store_true"); p.add_argument("--execute",action="store_true"); p.add_argument("--ppttc")
    a=p.parse_args()
    try:
        need(not a.report.exists(), "Report already exists"); need(a.report.resolve() not in {a.input.resolve(),a.plan.resolve(),a.output.resolve()}, "Report overlaps an input or output"); report=run(a); dump(a.report,report); print(json.dumps({"status":report["status"],"report":str(a.report.resolve())}))
    except Exception as e:
        if 'a' in locals() and a.input.exists() and sha(a.input.read_bytes())!=a.expected_sha256.upper(): print(json.dumps({"status":"REJECTED","error":"Source changed during failure handling"}),file=sys.stderr)
        print(json.dumps({"status":"REJECTED","error":str(e)}),file=sys.stderr); raise SystemExit(1)
if __name__=="__main__": main()
