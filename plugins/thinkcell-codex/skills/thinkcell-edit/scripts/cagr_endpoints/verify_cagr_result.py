"""Independent model/datasheet/field verifier for a native CAGR result."""
from __future__ import annotations

import argparse
import datetime as dt
import importlib.util
import json
import math
import sys
import zipfile
from pathlib import Path

from lxml import etree as E

P="http://schemas.openxmlformats.org/presentationml/2006/main"; A="http://schemas.openxmlformats.org/drawingml/2006/main"; R="http://schemas.openxmlformats.org/officeDocument/2006/relationships"; NS={"p":P,"a":A,"r":R}


def need(ok: bool, msg: str) -> None:
    if not ok: raise RuntimeError(msg)


def thirty360(a: dt.date, b: dt.date) -> float:
    days=360*(b.year-a.year)+30*(b.month-a.month)+(min(b.day,30)-min(a.day,30)); need(days>0,"non-positive official30/360 period"); return days/360


def main(a: argparse.Namespace) -> dict:
    prep=json.loads(a.prep_report.read_text(encoding="utf-8-sig")); data=json.loads(a.data.read_text(encoding="utf-8-sig")); dates=[dt.date.fromisoformat(x) for x in data["dates"]]; values=[float(x) for x in data["values"]]; need(len(dates)==len(values),"data length mismatch")
    prepared_chart=prep.get("chart") or prep.get("target_chart")
    need(isinstance(prepared_chart,dict) and prepared_chart.get("chart_name")==data["name"],"prepared semantic chart does not match data name")
    prepared_endpoints=prep.get("endpoints") or {}
    need(int(prepared_endpoints.get("source_index",-1))==a.source_index and int(prepared_endpoints.get("sink_index",-1))==a.sink_index,"prepared endpoints do not match verifier endpoints")
    need(bool(prep.get("source_unchanged",prep.get("target_unchanged",False))),"prepared source/target integrity flag is false")
    skill=a.skill_dir.resolve(); scripts=skill/"scripts"; impl=scripts/"thinkcell_no_click"/"implementation"; sys.path[:0]=[str(scripts),str(impl)]
    from chart_geometry import streams,xml  # type: ignore
    from read_named_datasheet import read as read_sheet  # type: ignore
    with zipfile.ZipFile(a.input) as z:
        parts=[x for x in z.namelist() if x.startswith("ppt/embeddings/oleObject") and x.endswith(".bin")]; charts=[]
        for part in parts:
            payload=streams(z.read(part)).get(("think-cellXML",));
            if not payload: continue
            root=xml(payload); ids={n.get("id"):n for n in root if n.get("id")}
            for chart in root:
                if chart.tag=="CSequenceChartSE" and chart.findtext("m_strName")==data["name"]: charts.append((part,root,ids,chart))
        need(len(charts)==1,"semantic chart selection is ambiguous")
        carrier,root,ids,chart=charts[0]; table=ids[chart.find("m_dtable").get("idref")]; members=table.find("m_cscdveccon").findall("elem"); graphs=[ids[x.get("idref")] for x in members]; need(len(graphs)>=a.expected_graphs,"CAGR graph count is below expected")
        target_tag=prep.get("inserted",{}).get("label_tag") or prep.get("inserted",{}).get("tag"); need(target_tag,"prepared report has no inserted label tag"); graph=next((g for g in graphs if ids.get(g.find("m_scdvecconlabel").get("idref"),E.Element("x")).findtext(".//m_bstrShapeName")==target_tag),None); need(graph is not None,"inserted graph label tag missing")
        def endpoint(anchor_id:str)->dict:
            out=[]
            for scalar in root.iter("CSequenceChartDataScalar"):
                ar=scalar.find("m_anchorVectorConnector");
                if ar is None or ar.get("idref")!=anchor_id: continue
                vector=next((v for v in root.iter("CSequenceChartDataVector") if scalar.get("id") in [x.get("idref") for x in v.findall("ocol/elem")]),None); need(vector is not None,"endpoint vector missing"); var=ids[scalar.find("m_varsrcAbsolute").get("idref")]; out.append({"index":int(vector.find("m_nIndexInDataSheet").get("val")),"value":float(var.find("m_varval").get("val")),"date":vector.findtext("m_varsrcCategory/m_varval/m_datetime") or vector.find("m_varsrcCategory/m_varval/m_datetime").get("val") if vector.find("m_varsrcCategory/m_varval/m_datetime") is not None else None})
            need(len(out)==1,"endpoint anchor is missing or ambiguous"); return out[0]
        source=endpoint(graph.find("m_anchorSource").get("idref")); sink=endpoint(graph.find("m_anchorSink").get("idref")); need((source["index"],sink["index"])==(a.source_index,a.sink_index),"endpoint indices changed")
        expected=(values[a.sink_index]/values[a.source_index])**(1/thirty360(dates[a.source_index],dates[a.sink_index]))-1; cagr=ids[graph.find("m_varsrcCAGR").get("idref")]; actual=float(cagr.find("m_varval").get("val")); need(math.isclose(actual,expected,rel_tol=0,abs_tol=2e-12),f"CAGR mismatch: {actual} vs {expected}")
        slide=E.fromstring(z.read("ppt/slides/slide1.xml")); rels=E.fromstring(z.read("ppt/slides/_rels/slide1.xml.rels")); tags={}
        for rel in rels:
            if rel.get("Type","").endswith("/tags"):
                part="ppt/tags/"+rel.get("Target").rsplit("/",1)[-1]; tags[rel.get("Id")]=E.fromstring(z.read(part)).find("{%s}tag"%P).get("val")
        shapes=[s for s in slide.xpath(".//*[p:nvSpPr/p:cNvPr]",namespaces=NS) if (r:=s.xpath(".//p:nvPr/p:custDataLst/p:tags/@r:id",namespaces=NS)) and tags.get(r[0])==target_tag]; need(len(shapes)==1,"inserted label physical shape missing or duplicated"); fields=shapes[0].xpath(".//a:fld",namespaces=NS); need(len(fields)==1,"inserted label numeric field missing"); visible=fields[0].findtext("a:t",namespaces=NS); need(visible==f"{expected*100:.1f}%",f"visible cache {visible!r} disagrees with expected")
        sheet=read_sheet(a.input,1,data["name"]); need(sheet.get("source_unchanged") and len(sheet.get("sheets",[]))==1,"datasheet readback failed"); matrix=sheet["sheets"][0]["matrix"]; need(len(matrix)>=3 and len(matrix[0])>=len(values)+1,"datasheet matrix too small"); cells={(c["row"],c["column"]):c for c in sheet["sheets"][0]["nonempty_cells"]}; epoch=dt.date(1899,12,30); actual_dates=[]
        for i in range(len(values)):
            date_cell=cells[(1,i+2)]; value_cell=cells[(3,i+2)]; need(date_cell.get("record_id")==5 and isinstance(date_cell.get("value"),(int,float)),"datasheet date is not typed"); actual_dates.append((epoch+dt.timedelta(days=int(date_cell["value"]))).isoformat()); need(math.isclose(float(value_cell["value"]),values[i],rel_tol=0,abs_tol=1e-9),"datasheet value mismatch")
        need(actual_dates==[x.isoformat() for x in dates],"datasheet typed dates mismatch")
    result={"status":"NATIVE_CAGR_INDEPENDENT_READBACK_PASS","input":str(a.input),"carrier":carrier,"chart_name":data["name"],"graph_count":len(graphs),"endpoints":{"source":source,"sink":sink},"expected_cagr":expected,"actual_cagr":actual,"visible_cache":visible,"typed_dates":actual_dates,"datasheet_values":values}
    a.report.write_text(json.dumps(result,indent=2),encoding="utf-8"); print(json.dumps(result,indent=2)); return result


if __name__=="__main__":
    p=argparse.ArgumentParser(description=__doc__); p.add_argument("--input",type=Path,required=True); p.add_argument("--prep-report",type=Path,required=True); p.add_argument("--data",type=Path,required=True); p.add_argument("--skill-dir",type=Path,required=True); p.add_argument("--report",type=Path,required=True); p.add_argument("--source-index",type=int,required=True); p.add_argument("--sink-index",type=int,required=True); p.add_argument("--expected-graphs",type=int,default=2); main(p.parse_args())
