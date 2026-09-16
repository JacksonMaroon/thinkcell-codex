"""Prepare a native CAGR closure imported into a chart with no CAGR.

This is a separate research adapter from ``parameterized_cagr.py``.  It uses
the proved donor's complete native CAGR graph, ellipse label, numeric field
and two arrow lines, then remaps every model identity and attaches fresh
anchors to selected target scalars.  Both donor and target charts are located
semantically by ``m_strName`` and carrier discovery; no OLE part or model ID
is hardcoded.

The target must be an ordinary internal CSequenceChartSE with zero existing
CAGR graphs.  The production route requires one scalar in each selected
category.  Multi-series total-owner hypotheses are deliberately disabled until
they have native proof.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import io
import json
import uuid
import zipfile
from pathlib import Path

from lxml import etree as E

P = "http://schemas.openxmlformats.org/presentationml/2006/main"; A = "http://schemas.openxmlformats.org/drawingml/2006/main"; R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"; PR = "http://schemas.openxmlformats.org/package/2006/relationships"; CT = "http://schemas.openxmlformats.org/package/2006/content-types"; NS = {"p": P, "a": A, "r": R}


def need(ok: bool, msg: str) -> None:
    if not ok: raise RuntimeError(msg)


def allocate_rid(used: set[str], serial: int) -> tuple[str, int]:
    """Allocate and register one slide relationship ID atomically."""
    while "rId" + str(serial) in used:
        serial += 1
    rid = "rId" + str(serial)
    used.add(rid)
    return rid, serial + 1


def sha(path: Path) -> str: return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def load_parameterized() -> object:
    path = Path(__file__).with_name("parameterized_cagr.py"); spec = importlib.util.spec_from_file_location("parameterized_cagr_lane", path); need(spec and spec.loader, "parameterized adapter is missing"); module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module); return module


def chart(raw: bytes, pc: object, skill_dir: Path, name: str):
    read_streams, read_xml, *_ = pc.init_helpers(skill_dir)
    found = [x for x in pc.chart_candidates(raw, read_streams, read_xml) if x["chart_name"] == name]
    need(len(found) == 1, f"chart name {name!r} must resolve exactly one chart")
    return found[0], read_streams, read_xml


def tag_map(z: zipfile.ZipFile, rels: E._Element) -> dict[str, str]:
    out = {}
    for rel in rels:
        if not rel.get("Type", "").endswith("/tags"): continue
        part = "ppt/tags/" + rel.get("Target", "").rsplit("/", 1)[-1]; need(part in z.namelist(), "tag part missing")
        vals = [x.get("val") for x in E.fromstring(z.read(part)).findall("{%s}tag" % P) if x.get("name", "").upper() == "THINKCELLSHAPEDONOTDELETE"]
        need(len(vals) == 1 and vals[0], "invalid tag part"); out[rel.get("Id")] = vals[0]
    return out


def physical(slide: E._Element, tags: dict[str, str], value: str, kind: str) -> E._Element:
    path = ".//*[p:nvSpPr/p:cNvPr]" if kind == "sp" else ".//*[p:nvCxnSpPr/p:cNvPr]"; matches=[]
    for node in slide.xpath(path, namespaces=NS):
        refs=node.xpath(".//p:nvPr/p:custDataLst/p:tags/@r:id", namespaces=NS)
        if len(refs)==1 and tags.get(refs[0])==value: matches.append(node)
    need(len(matches)==1, f"donor tag does not resolve to one {kind}"); return matches[0]


def prepare(target_path: Path, donor_path: Path, target_sha: str, output: Path, report: Path, workdir: Path, skill_dir: Path, target_name: str, donor_name: str, source_index: int, sink_index: int, total_owner: str | None = None) -> dict:
    pc = load_parameterized(); target_path=target_path.resolve(); donor_path=donor_path.resolve(); output=output.resolve(); report=report.resolve(); workdir=workdir.resolve()
    need(target_path.is_file() and donor_path.is_file() and sha(target_path)==target_sha.upper(), "target hash mismatch"); need(output.parent==workdir and report.parent==workdir and not output.exists() and not report.exists(), "outputs must be new direct children")
    target_raw=target_path.read_bytes(); donor_raw=donor_path.read_bytes(); tc, read_streams, read_xml=chart(target_raw,pc,skill_dir,target_name); dc,_,_=chart(donor_raw,pc,skill_dir,donor_name)
    with zipfile.ZipFile(io.BytesIO(target_raw)) as tz, zipfile.ZipFile(io.BytesIO(donor_raw)) as dz:
        target_root=tc["root"]; donor_root=dc["root"]; tids=pc.ids_of(target_root); dids=pc.ids_of(donor_root); tchart=tids[tc["chart_id"]]; dchart=dids[dc["chart_id"]]; ttable=tids[pc.one(tchart,"m_dtable").get("idref")]; dtable=dids[pc.one(dchart,"m_dtable").get("idref")]
        tmembers=ttable.find("m_cscdveccon"); need(tmembers is not None and not tmembers.findall("elem"), "target already contains a native connector feature")
        dmembers=dtable.find("m_cscdveccon"); need(dmembers is not None and dmembers.findall("elem"), "donor has no CAGR feature"); dgraph=dids[dmembers.find("elem").get("idref")]; need(dgraph.tag=="CSequenceChartDataVectorCAGR", "donor feature is not CAGR")
        # Resolve target scalar endpoints by category index. A multi-series
        # target is refused until a native total/extent scalar is specified.
        scalar_map=pc.endpoint_scalars(target_root,ttable); need(source_index in scalar_map and sink_index in scalar_map and sink_index>source_index, "target endpoint index is invalid")
        need(total_owner is None, "total-owner hypotheses are unverified and disabled in the production importer")
        need(len(scalar_map[source_index])==1 and len(scalar_map[sink_index])==1, "target has multiple series; select a proved total/extent scalar first")
        tslide=E.fromstring(tz.read("ppt/slides/slide1.xml")); trels=E.fromstring(tz.read("ppt/slides/_rels/slide1.xml.rels")); tct=E.fromstring(tz.read("[Content_Types].xml")); dslide=E.fromstring(dz.read("ppt/slides/slide1.xml")); drels=E.fromstring(dz.read("ppt/slides/_rels/slide1.xml.rels")); dtags=tag_map(dz,drels); dlabel=dids[pc.one(dgraph,"m_scdvecconlabel").get("idref")]; dlabel_tag=dlabel.findtext(".//m_bstrShapeName"); dlabel_shape=physical(dslide,dtags,dlabel_tag,"sp"); dline_refs=pc.one(pc.one(dgraph,"m_pptgenlineArrow"),"m_cpptline").findall("elem"); dline_shapes=[physical(dslide,dtags,pc.one(dids[x.get("idref")],"m_bstrShapeName").text,"cxnSp") for x in dline_refs]
        next_id=max([int(x) for x in tids if x.isdigit()] + [0])+1; mapping={dgraph.get("id"):str(next_id)}; next_id+=1; fields=("m_varsrcRelative","m_varsrcAbsolute","m_varsrcRelativeGroup","m_varsrcCAGR","m_scdvecconlabel")
        for f in fields: mapping[pc.one(dgraph,f).get("idref")]=str(next_id); next_id+=1
        for x in dline_refs: mapping[x.get("idref")]=str(next_id); next_id+=1
        source_anchor,sink_anchor=str(next_id),str(next_id+1); next_id+=2; textvar_id=str(next_id); clone=copy.deepcopy(dgraph); clone.set("id",mapping[dgraph.get("id")]);
        for node in clone.iter():
            if node.get("idref") in mapping: node.set("idref",mapping[node.get("idref")])
        clone.find("m_anchorSource").set("idref",source_anchor); clone.find("m_anchorSink").set("idref",sink_anchor); cloned=[]
        for old_id in [pc.one(dgraph,f).get("idref") for f in fields]+[x.get("idref") for x in dline_refs]:
            node=copy.deepcopy(dids[old_id]); node.set("id",mapping[old_id]);
            if node.tag=="CVariableSource": pc.one(node,"m_guid").set("val",str(uuid.uuid4()))
            if node.tag=="CSequenceChartDataVectorConnectorLabel": pc.one(node,".//m_bstrShapeName").text=pc.unique_tag("t")
            if node.tag=="CPPTLine": pc.one(node,"m_bstrShapeName").text=pc.unique_tag("t")
            cloned.append(node)
        new_cagr=next(x for x in cloned if x.get("id")==mapping[pc.one(dgraph,"m_varsrcCAGR").get("idref")]); cache,key=pc.literal_percent(float(pc.one(new_cagr,"m_varval").get("val"))); tv=E.fromstring(pc.NUMERIC_PERCENT_XML); tv.set("id",textvar_id); tv.find("m_bstrFormat").text=key; E.SubElement(pc.one(new_cagr,"m_ctextvar"),"elem",idref=textvar_id)
        for idx,anchor_id in ((source_index,source_anchor),(sink_index,sink_anchor)):
            owner=scalar_map[idx][0]
            need(owner.find("m_anchorVectorConnector") is None,"target endpoint already owns a native feature"); E.SubElement(owner,"m_anchorVectorConnector",idref=anchor_id); anchor=E.Element("CSequenceChartAnchor",id=anchor_id); mem=E.SubElement(anchor,"m_cfeature",length="1"); E.SubElement(mem,"elem",idref=clone.get("id")); target_root.append(anchor)
        E.SubElement(tmembers,"elem",idref=clone.get("id")); tmembers.set("length",str(len(tmembers.findall("elem"))));
        # Keep chart/table adjacency and insert the feature at the same model
        # boundary used by the native donor: after chart gridline/gap objects,
        # immediately before the data-series group.  The donor's ordering is
        # graph -> arrows -> variable sources -> label -> text variable.
        insert_at = next((i for i, node in enumerate(target_root)
                          if i > target_root.index(ttable)
                          and node.tag == "CSequenceChartDataSeriesGroup"), len(target_root))
        line_nodes = [x for x in cloned if x.tag == "CPPTLine"]
        field_nodes = [x for x in cloned if x.tag == "CVariableSource"]
        label_nodes = [x for x in cloned if x.tag == "CSequenceChartDataVectorConnectorLabel"]
        target_root[insert_at:insert_at] = [clone, *line_nodes, *field_nodes, *label_nodes, tv]
        changed=pc.init_helpers(skill_dir)[4](tz.read(tc["carrier"]),E.tostring(target_root,encoding="utf-8"),workdir); target_tag=pc.one(next(x for x in cloned if x.get("id")==mapping[pc.one(dgraph,"m_scdvecconlabel").get("idref")]),".//m_bstrShapeName").text; used=[int(x.get("id")) for x in tslide.xpath(".//p:cNvPr",namespaces=NS) if (x.get("id") or "").isdigit()]; usedr={x.get("Id") for x in trels}; serial=1; parts={}
        def add(shape:E._Element, tag:str, kind:str):
            nonlocal serial
            cp=copy.deepcopy(shape); sid=str(max(used)+1); used.append(int(sid)); pc.one(cp,".//p:cNvPr").set("id",sid); pc.one(cp,".//p:cNvPr").set("name","Imported CAGR "+kind+" "+sid)
            rid, serial = allocate_rid(usedr, serial); part="cagr-import-"+str(serial-1)+".xml"; pc.one(cp,".//p:nvPr/p:custDataLst/p:tags").set("{%s}id"%R,rid); pc.one(tslide,".//p:spTree").append(cp); E.SubElement(trels,"{%s}Relationship"%PR,Id=rid,Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags",Target="../tags/"+part); E.SubElement(tct,"{%s}Override"%CT,PartName="/ppt/tags/"+part,ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"); parts["ppt/tags/"+part]=(f'<p:tagLst xmlns:p="{P}"><p:tag name="THINKCELLSHAPEDONOTDELETE" val="{tag}"/></p:tagLst>').encode()
        label=copy.deepcopy(dlabel_shape); sid=str(max(used)+1); used.append(int(sid)); pc.one(label,".//p:cNvPr").set("id",sid); pc.one(label,".//p:cNvPr").set("name","Imported CAGR label "+sid)
        para=pc.one(label,".//a:p")
        for child in list(para):
            if E.QName(child).localname in {"r", "fld"}:
                para.remove(child)
        ppr=pc.one(para,"a:pPr")
        source_para=pc.one(dlabel_shape,".//a:p")
        run=next((x for x in source_para if E.QName(x).localname=="r"), None)
        need(run is not None, "donor label has no run properties")
        rpr=copy.deepcopy(pc.one(run,"a:rPr")); fld=E.Element("{%s}fld"%A,id="{"+str(uuid.uuid4()).upper()+"}",type="datetime"+key); fld.append(rpr); E.SubElement(fld,"{%s}t"%A).text=cache; para.insert(para.index(ppr)+1,fld);
        rid, serial = allocate_rid(usedr, serial); part="cagr-import-label.xml"; pc.one(label,".//p:nvPr/p:custDataLst/p:tags").set("{%s}id"%R,rid); pc.one(tslide,".//p:spTree").append(label); E.SubElement(trels,"{%s}Relationship"%PR,Id=rid,Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags",Target="../tags/"+part); E.SubElement(tct,"{%s}Override"%CT,PartName="/ppt/tags/"+part,ContentType="application/vnd.openxmlformats-officedocument.presentationml.tags+xml"); parts["ppt/tags/"+part]=(f'<p:tagLst xmlns:p="{P}"><p:tag name="THINKCELLSHAPEDONOTDELETE" val="{target_tag}"/></p:tagLst>').encode()
        for shape,node in zip(dline_shapes,[x for x in cloned if x.tag=="CPPTLine"]): add(shape,pc.one(node,"m_bstrShapeName").text,"arrow")
        modified={tc["carrier"]:changed,"ppt/slides/slide1.xml":E.tostring(tslide,xml_declaration=True,encoding="UTF-8",standalone=True),"ppt/slides/_rels/slide1.xml.rels":E.tostring(trels,xml_declaration=True,encoding="UTF-8",standalone=True),"[Content_Types].xml":E.tostring(tct,xml_declaration=True,encoding="UTF-8",standalone=True),**parts}
        with zipfile.ZipFile(output,"x") as out:
            for item in tz.infolist():
                if item.filename not in modified: out.writestr(copy.copy(item),tz.read(item.filename))
            for name,data in modified.items(): out.writestr(name,data)
    result={"status":"PREPARED_CAGR_DONOR_IMPORT_REGENERATION_REQUIRED","target_sha256":target_sha.upper(),"candidate_sha256":sha(output),"target_chart":{k:tc[k] for k in ("carrier","chart_id","chart_name","model_version","series_count","category_count","cagr_count")},"donor_chart":{k:dc[k] for k in ("carrier","chart_id","chart_name","model_version","series_count","category_count","cagr_count")},"endpoints":{"source_index":source_index,"sink_index":sink_index,"series_mode":"single-series-scalar","total_owner":None},"inserted":{"feature_model_id":clone.get("id"),"label_tag":target_tag,"text_variable_id":textvar_id,"source_anchor_id":source_anchor,"sink_anchor_id":sink_anchor},"target_unchanged":sha(target_path)==target_sha.upper()}; report.write_text(json.dumps(result,indent=2),encoding="utf-8"); return result


if __name__=="__main__":
    p=argparse.ArgumentParser(description=__doc__); p.add_argument("--target",type=Path,required=True); p.add_argument("--donor",type=Path,required=True); p.add_argument("--target-sha256",required=True); p.add_argument("--output",type=Path,required=True); p.add_argument("--report",type=Path,required=True); p.add_argument("--workdir",type=Path,required=True); p.add_argument("--skill-dir",type=Path,required=True); p.add_argument("--target-name",required=True); p.add_argument("--donor-name",required=True); p.add_argument("--source-index",type=int,required=True); p.add_argument("--sink-index",type=int,required=True); a=p.parse_args(); print(json.dumps(prepare(a.target,a.donor,a.target_sha256,a.output,a.report,a.workdir,a.skill_dir.resolve(),a.target_name,a.donor_name,a.source_index,a.sink_index,None),indent=2))
