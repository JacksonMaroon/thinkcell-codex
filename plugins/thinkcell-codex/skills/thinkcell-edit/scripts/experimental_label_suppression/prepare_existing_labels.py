"""Prepare a guarded removal of existing think-cell scalar and sum labels.

Experimental: this removes only labels already owned by a selected sequence
chart.  It never inserts or fabricates labels, annotations, or data.
"""
from __future__ import annotations
import argparse, copy, io, json, posixpath, subprocess, sys, tempfile, zipfile
from pathlib import Path
from lxml import etree as E

HERE=Path(__file__).resolve().parent; SCRIPTS=HERE.parent
sys.path[:0]=[str(SCRIPTS),str(SCRIPTS/'thinkcell_no_click'/'implementation')]
from chart_geometry import inventory, need, sha, streams, xml
from runtime import powershell, powershell_env

NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','pr':'http://schemas.openxmlformats.org/package/2006/relationships'}

def refs(root, wanted):
 out=[]
 for n in root.iter():
  if n.get('idref') in wanted: out.append((n.getparent().tag if n.getparent() is not None else '',n.tag,n.get('idref')))
 return out
def all_refs_resolve(root):
 ids={n.get('id') for n in root if n.get('id')}
 missing=[]
 for n in root.iter():
  r=n.get('idref')
  if r and r!='0' and r not in ids: missing.append((n.tag,r))
 need(not missing,'dangling model idrefs: '+str(missing[:5]))
def val(ids,node,field):
 r=node.find(field); need(r is not None,'missing '+field)
 v0=r.find('m_varval')
 if v0 is not None:
  return (v0.get('val') if v0.get('val') is not None else (v0.text or '')).strip(),r
 rr=r if r.get('idref') else r.find('.//*[@idref]'); need(rr is not None,'missing variable reference '+field)
 s=ids.get(rr.get('idref')); need(s is not None and s.tag=='CVariableSource','bad '+field)
 v=s.find('m_varval'); need(v is not None,'source has no value')
 return (v.get('val') if v.get('val') is not None else (v.text or '')).strip(),s
def resolve_model(raw, chart_name, category, series_names, suppress_sums, all_carrier_sums, hide_category_axis):
 _,charts,_=inventory(raw)
 matches=[c for c in charts if c['owner'].tag=='CSequenceChartSE' and (c['owner'].findtext('m_strName') or '').strip()==chart_name]
 need(len(matches)==1,'target chart must be unique by automation name')
 c=matches[0]; root=c['doc']['root']; ids=c['doc']['ids']; owner=ids[c['owner'].get('id')]; table=ids[owner.find('m_dtable').get('idref')]
 series=[ids[e.get('idref')] for e in table.find('m_cscdser')]
 names=[val(ids,s,'m_varsrc')[0] for s in series]; need(len(names)==len(set(names)),'series names ambiguous')
 vecs=[ids[e.get('idref')] for e in table.findall('./ocol/elem')]
 cats=[val(ids,v,'m_varsrcCategory')[0] for v in vecs]; need(cats.count(category)==1,'requested category ambiguous or missing')
 vec=vecs[cats.index(category)]; scalars=[ids[e.get('idref')] for e in vec.findall('./ocol/elem')]
 need(len(scalars)==len(series),'USPS scalar/series map invalid')
 selections=[]
 need(len(series_names)==len(set(series_names)),'requested series repeated')
 for series_name in series_names:
  need(series_name in names,'requested series missing: '+series_name)
  scalar=scalars[names.index(series_name)]; sid=scalar.get('id'); labref=scalar.find('m_scdlabel'); need(labref is not None and labref.get('idref')!='0','selected scalar has no existing label')
  label=ids.get(labref.get('idref')); need(label is not None and label.tag=='CSequenceChartDataScalarLabel','bad label reference')
  numeric,src=val(ids,scalar,'m_varsrcAbsolute'); textref=src.find('m_ctextvar')
  tv=None
  textnode=textref.find('.//*[@idref]') if textref is not None else None
  if textref is not None and (textref.get('idref') or textnode is not None):
   textid=textref.get('idref') or textnode.get('idref'); tv=ids.get(textid); need(tv is not None and tv.tag=='CTextVariable','bad text variable')
  inbound=refs(root,{label.get('id')}); need(inbound==[(scalar.tag,'m_scdlabel',label.get('id'))],'label not uniquely owned: '+str(inbound))
  conn=label.find('./m_pptgenlineConnector/m_cpptline'); cid=(conn.get('idref') or (conn.find('.//*[@idref]').get('idref') if conn is not None and conn.find('.//*[@idref]') is not None else None)) if conn is not None else None
  if cid and cid!='0':
   cin=refs(root,{cid}); need(cin==[(conn.tag,'elem',cid)],'connector not uniquely owned: '+str(cin))
  tag=label.findtext('./m_ppttb/m_bstrShapeName')
  direct_precision = tv is None
  if direct_precision:
   placed=label.find('./m_ppttb/m_bPlaced'); need(label.find('m_prec') is not None and placed is not None and placed.get('val')=='0','direct-precision label profile differs')
  else: need(tag,'bound label has no physical tag')
  selections.append({'kind':'scalar_label','series':series_name,'owner_id':sid,'ref_field':'m_scdlabel','scalar_id':sid,'label_id':label.get('id'),'source_id':src.get('id'),'source_value':numeric,'text_variable_id':tv.get('id') if tv is not None else None,'connector_id':cid if cid and cid!='0' else None,'tag':tag,'physical':not direct_precision,'direct_precision':direct_precision})
 if suppress_sums:
  group_ids=[]
  for container in [table]+vecs:
   group_ids.extend(e.get('idref') for e in container.findall('./m_cscdscgrp/elem') if e.get('idref') and e.get('idref')!='0')
  need(group_ids and len(group_ids)==len(set(group_ids)),'sum-label groups are missing or ambiguous')
  sums=[]
  for gid in group_ids:
   host=ids.get(gid); need(host is not None and host.tag=='CSequenceChartDataScalarGroup','unexpected sum-label group')
   ref=host.find('m_scdsumlabel'); need(ref is not None and ref.get('idref')!='0','group has no existing sum label')
   sums.append((host,ids.get(ref.get('idref'))))
  if all_carrier_sums:
   sums=[]
   for label in [n for n in root if n.tag=='CSequenceChartDataSumLabel']:
    hosts=[n for n in root if any(x.tag=='m_scdsumlabel' and x.get('idref')==label.get('id') for x in n)]
    need(len(hosts)==1,'carrier sum label host is ambiguous')
    sums.append((hosts[0],label))
  for hostnode,label in sums:
   need(label is not None and label.tag=='CSequenceChartDataSumLabel','bad sum label reference')
   inbound=refs(root,{label.get('id')}); need(len(inbound)==1 and inbound[0][1]=='m_scdsumlabel','sum label not uniquely owned: '+str(inbound))
   tag=label.findtext('./m_ppttb/m_bstrShapeName'); need(tag,'sum label has no physical tag')
   conn=label.find('./m_pptgenlineConnector/m_cpptline'); cid=(conn.get('idref') or (conn.find('.//*[@idref]').get('idref') if conn is not None and conn.find('.//*[@idref]') is not None else None)) if conn is not None else None
   if cid and cid!='0': need(refs(root,{cid})==[(conn.tag,'elem',cid)],'sum connector not uniquely owned')
   selections.append({'kind':'sum_label','owner_id':hostnode.get('id'),'ref_field':'m_scdsumlabel','label_id':label.get('id'),'connector_id':cid if cid and cid!='0' else None,'tag':tag})
 axis=None
 if hide_category_axis:
  axis_ref=owner.find('m_daxisCategory'); need(axis_ref is not None and axis_ref.get('idref')!='0','category axis missing')
  axis=ids.get(axis_ref.get('idref')); need(axis is not None and axis.tag=='CSequenceChartDataAxis','category axis identity changed')
  vis=axis.find('./m_linestyle/m_bVisible'); need(vis is not None and vis.get('val')=='1','category axis visibility must be 1')
 all_refs_resolve(root)
 return c,root,ids,owner,axis,selections
def tag_shapes(raw, selections):
 with zipfile.ZipFile(io.BytesIO(raw)) as z:
  slidepart='ppt/slides/slide1.xml'; relpart='ppt/slides/_rels/slide1.xml.rels'; slide=E.fromstring(z.read(slidepart)); rels=E.fromstring(z.read(relpart))
  relmap={x.get('Id'):x for x in rels}
  found=[]
  for sel in selections:
   if not sel.get('physical',True):
    matches=[]
    if sel.get('tag'):
     for sp in slide.xpath('.//p:sp',namespaces=NS):
      rid=sp.xpath('./p:nvSpPr/p:nvPr/p:custDataLst/p:tags/@r:id',namespaces=NS)
      if not rid: continue
      rel=relmap.get(rid[0]); need(rel is not None,'tag relationship missing')
      part=posixpath.normpath(posixpath.join('ppt/slides',rel.get('Target')))
      if sel['tag'] in z.read(part).decode('utf-8',errors='strict'): matches.append(sp)
    need(not matches,'direct-precision label unexpectedly has a physical shape')
    sel['physical_shape_absent']=True
    continue
   matches=[]
   for sp in slide.xpath('.//p:sp',namespaces=NS):
    rid=sp.xpath('./p:nvSpPr/p:nvPr/p:custDataLst/p:tags/@r:id',namespaces=NS)
    if not rid: continue
    rel=relmap.get(rid[0]); need(rel is not None,'tag relationship missing')
    part=posixpath.normpath(posixpath.join('ppt/slides',rel.get('Target')))
    if sel['tag'] in z.read(part).decode('utf-8',errors='strict'): matches.append((sp,rid[0],part))
   need(len(matches)==1,'physical tag must identify one shape: '+sel['tag'])
   sp,rid,part=matches[0]; usage=slide.xpath('.//p:tags[@r:id=$x]',namespaces=NS,x=rid); need(len(usage)==1,'tag relationship shared')
   need(sum(1 for x in rels if x.get('Target')==relmap[rid].get('Target'))==1,'tag part relationship shared')
   # Tag parts are not targets of any other package relationship.
   basename=part.split('/')[-1]; uses=sum(z.read(n).count(basename.encode()) for n in z.namelist() if n.endswith('.rels'))
   need(uses==1,'tag part has external relationship use')
   sel.update({'shape_id':sp.xpath('./p:nvSpPr/p:cNvPr/@id',namespaces=NS)[0],'relationship_id':rid,'tag_part':part})
   sel['visible_text']=''.join(sp.xpath('.//a:t/text()',namespaces={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}))
   found.append((sp,rid,part))
 return slide,rels,found
def apply_model(root, ids, axis, sels, reverse=False):
 for s in sels:
  scalar=ids[s['owner_id']]; ref=scalar.find(s['ref_field']); need(ref is not None,'label host reference missing')
  ref.set('idref',s['label_id'] if reverse else '0')
 if reverse:
  raise RuntimeError('reverse requires original clone, not reinsertion')
 for s in sels:
  node=ids[s['label_id']]; root.remove(node)
  if s['connector_id']: root.remove(ids[s['connector_id']])
 if axis is not None: axis.find('./m_linestyle/m_bVisible').set('val','0')
def prepare(source, output, expected_source_sha256, chart_name, category, series_names, suppress_sums, all_carrier_sums, hide_category_axis, report_path=None):
 source,output=Path(source),Path(output); need(source.resolve()!=output.resolve() and not output.exists(),'output must be a distinct new file')
 if report_path:
  report_path=Path(report_path); need(report_path.resolve() not in {source.resolve(),output.resolve()} and not report_path.exists(),'report must be a distinct new file')
 raw=source.read_bytes(); source_hash=sha(raw); need(source_hash==expected_source_sha256.upper(),'source SHA-256 mismatch before mutation')
 c,root,ids,owner,axis,sels=resolve_model(raw,chart_name,category,series_names,suppress_sums,all_carrier_sums,hide_category_axis); slide,rels,phys=tag_shapes(raw,sels)
 before=c['doc']['streams'][('think-cellXML',)]; datasheet={ '/'.join(k):sha(v) for k,v in c['doc']['streams'].items() if len(k)>1 and k[0].startswith('think-cellChild') }
 need(datasheet,'embedded datasheet not found')
 source_values={s['source_id']:s['source_value'] for s in sels if s['kind']=='scalar_label'}
 baseline=E.tostring(root,method='c14n'); apply_model(root,ids,axis,sels); all_refs_resolve(root)
 for source_id, old_value in source_values.items():
  value=ids[source_id].find('m_varval'); need(value is not None and value.get('val')==old_value,'target source value changed: '+source_id)
 # Reverse from an independently preserved tree proves all changes are known and bounded.
 expected=xml(before); eids={n.get('id'):n for n in expected if n.get('id')}; eaxis=eids[axis.get('id')] if axis is not None else None; apply_model(expected,eids,eaxis,sels)
 after=E.tostring(root,encoding='utf-8'); need(E.tostring(xml(after),method='c14n')==E.tostring(expected,method='c14n'),'model differs from exact planned mutation')
 with tempfile.TemporaryDirectory(dir=output.parent,prefix='suppression_') as td:
  td=Path(td); carrier=td/'carrier.bin'; payload=td/'model.xml'; carrier.write_bytes(c['doc']['ole']); payload.write_bytes(after)
  q=subprocess.run([powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(SCRIPTS/'thinkcell_no_click'/'implementation'/'replace_ole_stream.ps1'),'-StoragePath',str(carrier),'-StreamBytesPath',str(payload)],capture_output=True,text=True,env=powershell_env(),timeout=60); need(q.returncode==0,'OLE patch failed: '+q.stderr[-400:]); changed=carrier.read_bytes(); st=streams(changed)
  need(set(st)==set(c['doc']['streams']) and all(st[k]==v for k,v in c['doc']['streams'].items() if k!=('think-cellXML',)),'untouched CFB stream drift')
  for sp,rid,part in phys: sp.getparent().remove(sp); rel=next(x for x in rels if x.get('Id')==rid); rels.remove(rel)
  slide_bytes=E.tostring(slide,xml_declaration=True,encoding='UTF-8',standalone=True); rel_bytes=E.tostring(rels,xml_declaration=True,encoding='UTF-8',standalone=True); removed={x[2] for x in phys}
  with zipfile.ZipFile(io.BytesIO(raw)) as zin,zipfile.ZipFile(output,'x') as zout:
   zout.comment=zin.comment
   for e in zin.infolist():
    if e.filename in removed: continue
    data=changed if e.filename==c['doc']['part'] else slide_bytes if phys and e.filename=='ppt/slides/slide1.xml' else rel_bytes if phys and e.filename=='ppt/slides/_rels/slide1.xml.rels' else zin.read(e.filename)
    zout.writestr(copy.copy(e),data)
 with zipfile.ZipFile(io.BytesIO(raw)) as a,zipfile.ZipFile(output) as b:
  changed_entries=[n for n in b.namelist() if a.read(n)!=b.read(n)]; expected_changed={c['doc']['part']} | ({'ppt/slides/slide1.xml','ppt/slides/_rels/slide1.xml.rels'} if phys else set()); need(set(changed_entries)==expected_changed,'unexpected OOXML drift: '+str(changed_entries)); need(set(a.namelist())-set(b.namelist())==removed,'unexpected entry removal')
 need(sha(source.read_bytes())==source_hash,'source changed')
 report={'status':'PREPARED_STATIC_CANDIDATE_NATIVE_PROOF_REQUIRED','source_path':str(source.resolve()),'output_path':str(output.resolve()),'report_path':str(report_path.resolve()) if report_path else None,'expected_source_sha256':expected_source_sha256.upper(),'source_sha256':source_hash,'output_sha256':sha(output.read_bytes()),'target_chart':{'automation_name':chart_name,'owner_id':owner.get('id')},'selectors':{'category':category,'series':series_names,'suppress_existing_sum_labels':suppress_sums,'suppress_all_existing_carrier_sum_labels':all_carrier_sums,'hide_existing_category_axis_line':hide_category_axis},'selected':sels,'axis':({'id':axis.get('id'),'field':'m_linestyle/m_bVisible','from':'1','to':'0'} if axis is not None else None),'target_source_values_unchanged':source_values,'datasheet_stream_sha256':datasheet,'changed_zip_entries':sorted(expected_changed),'removed_zip_entries':sorted(removed),'source_unchanged':True,'gates_passed':['expected source SHA-256 matched before mutation','unique chart name/category/series resolution','unique label and connector ownership','direct-precision labels proven untagged and model-only','physical tags unique when present','no dangling model idrefs','target source values unchanged','non-model CFB streams unchanged','embedded datasheet streams unchanged','no OOXML drift outside selected carrier/slide/rels and selected tag parts'],'native_gates_remaining':['official regeneration','native reopen and render','changed-data proof']}
 if report_path: Path(report_path).write_text(json.dumps(report,indent=2),encoding='utf-8')
 return report
if __name__=='__main__':
 p=argparse.ArgumentParser(description=__doc__); p.add_argument('--input',required=True); p.add_argument('--output',required=True); p.add_argument('--expected-source-sha256',required=True); p.add_argument('--chart-name',required=True); p.add_argument('--category',required=True); p.add_argument('--series',action='append',required=True); p.add_argument('--suppress-existing-sum-labels',action='store_true'); p.add_argument('--suppress-all-existing-carrier-sum-labels',action='store_true'); p.add_argument('--hide-category-axis-line',action='store_true'); p.add_argument('--report')
 a=p.parse_args(); need(not a.suppress_all_existing_carrier_sum_labels or a.suppress_existing_sum_labels,'carrier sum scope requires --suppress-existing-sum-labels'); print(json.dumps(prepare(a.input,a.output,a.expected_source_sha256,a.chart_name,a.category,a.series,a.suppress_existing_sum_labels,a.suppress_all_existing_carrier_sum_labels,a.hide_category_axis_line,a.report),indent=2))
