"""Prepare one tested native axis-break insertion candidate on a copy.

No synthetic nodes or drawing substitutes are used.  The target is a native
ordinary clustered-column donor with an existing outlier data contract.  The
break graph and six drawing parts come from the installed native break donor.
"""
from __future__ import annotations
import argparse,base64,collections,copy,hashlib,io,json,posixpath,subprocess,sys,tempfile,uuid,zipfile
from pathlib import Path
from lxml import etree as E
HERE=Path(__file__).resolve().parent
# This module is installed under <bundled skill>/scripts.  It deliberately
# resolves helpers from that bundle, never from a workspace or installed skill.
SKILL_ROOT=HERE.parent
sys.path[:0]=[str(HERE),str(HERE/'thinkcell_no_click'/'implementation')]
from chart_geometry import inventory,streams,xml
from runtime import powershell,powershell_env
P='http://schemas.openxmlformats.org/presentationml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';NS={'p':P,'r':R};Q=lambda n:'{'+R+'}'+n; FIELDS=('m_pptshpLow','m_pptshpHigh','m_pptshpBody')
def need(x,m):
 if not x:raise RuntimeError(m)
def sha(b):return hashlib.sha256(b).hexdigest()
def tag():return 't'+base64.urlsafe_b64encode(uuid.uuid4().bytes).decode().rstrip('=')
def one(path):
 raw=Path(path).read_bytes();_,cs,_=inventory(raw);cs=[c for c in cs if c['exact'] and c['owner'].tag=='CSequenceChartSE'];need(len(cs)==1,'Require one exact ordinary sequence chart');c=cs[0];ids=c['doc']['ids'];axis=ids[c['owner'].find('m_daxisPrimaryValue').get('idref')];return raw,c,axis
def rels(z,part):
 rp=str(Path(part).parent/'_rels'/(Path(part).name+'.rels')).replace('\\','/');root=E.fromstring(z.read(rp));m={r.get('Id'):posixpath.normpath(posixpath.join(str(Path(part).parent).replace('\\','/'),r.get('Target'))) for r in root};return rp,root,m
def parts(z,slide,names):
 _,_,rm=rels(z,slide);s=E.fromstring(z.read(slide));out=[]
 for n in list(s.find('p:cSld/p:spTree',NS)):
  for t in n.findall('.//p:tags',NS):
   p=rm.get(t.get(Q('id')));vals=[] if p not in z.namelist() else [x.get('val') for x in E.fromstring(z.read(p)) if x.get('name','').upper()=='THINKCELLSHAPEDONOTDELETE']
   hit=set(vals)&set(names)
   if hit:need(len(hit)==1,'physical shape carries multiple break tags');out.append((n,next(iter(hit)),t.get(Q('id')),p))
 need(len(out)==6 and {x[1] for x in out}==set(names),'Require six donor physical parts');return s,out
def replace(ole,model):
 h=HERE/'thinkcell_no_click'/'implementation'/'replace_ole_stream.ps1'
 with tempfile.TemporaryDirectory(prefix='tc_insert_real_') as d:
  d=Path(d);c=d/'c.bin';m=d/'m.xml';c.write_bytes(ole);m.write_bytes(model);p=subprocess.run([powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(h),'-StoragePath',str(c),'-StreamBytesPath',str(m)],capture_output=True,text=True,env=powershell_env(),timeout=60);need(p.returncode==0,p.stderr[-500:]);return c.read_bytes()
def put(srcraw,out,repl):
 with zipfile.ZipFile(io.BytesIO(srcraw)) as z,zipfile.ZipFile(out,'x') as d:
  d.comment=z.comment
  for x in z.infolist():d.writestr(copy.copy(x),repl.get(x.filename,z.read(x.filename)))
def build(ordinary,broken,expected_ordinary_sha256,expected_donor_sha256,out,report):
 traw,tc,taxis=one(ordinary);braw,bc,baxis=one(broken)
 need(sha(traw).upper()==expected_ordinary_sha256.upper(),'ordinary source hash mismatch before mutation')
 need(sha(braw).upper()==expected_donor_sha256.upper(),'break donor hash mismatch before mutation')
 need(tc['doc']['root'].find('version').get('val')==bc['doc']['root'].find('version').get('val'),'ordinary and donor native model versions differ')
 prior_interval_cardinality=[x.get('length') for x in taxis.findall('m_vecintvlOrdinal')]
 need(len(taxis.findall('m_cdaxisbreak/elem'))==0,'target already has a native axis break')
 need(len(taxis.findall('m_vecintvlOrdinal'))==1 and prior_interval_cardinality==['12'],'unsupported ordinary interval topology')
 need(not taxis.findall('m_cdaxistickmark/elem'),'unsupported target tickmark topology')
 brref=baxis.findall('m_cdaxisbreak/elem');need(len(brref)==1,'Donor lacks one break');bids=bc['doc']['ids'];br=bids[brref[0].get('idref')];srefs=br.findall('m_cpptbreakshp/elem');need(len(srefs)==2,'Donor must supply the tested two-shape seed graph');shapes=[bids[x.get('idref')] for x in srefs];old=[x.find(f+'/m_bstrShapeName').text for x in shapes for f in FIELDS];need(len(old)==6 and len(set(old))==6,'Donor tags invalid')
 slide='ppt/slides/slide1.xml'
 with zipfile.ZipFile(io.BytesIO(braw)) as bz: bslide,bparts=parts(bz,slide,old)
 root=xml(tc['doc']['streams'][('think-cellXML',)]);ids={x.get('id'):x for x in root if x.get('id')};taxis=ids[taxis.get('id')];newids=[str(max(int(x) for x in ids if x.isdigit())+i) for i in (1,2,3)];clones=[copy.deepcopy(br),*(copy.deepcopy(x) for x in shapes)]
 for x,i in zip(clones,newids):x.set('id',i)
 clones[0].find('m_cpptbreakshp/elem[1]').set('idref',newids[1]);clones[0].find('m_cpptbreakshp/elem[2]').set('idref',newids[2]);new=[tag() for _ in range(6)]
 for node,names in zip(clones[1:],(new[:3],new[3:])):
  for f,n in zip(FIELDS,names):node.find(f+'/m_bstrShapeName').text=n
# Existing ordinary axis has its validated no-break intervals.  Replace its
# vector *in place* with the observed current-version three-interval topology,
# scaled to this axis's actual min/max.  Think-cell's XML model is ordered: a
# former experiment inserted these fields at child index zero and the template
# parser rejected it before JSON generation.
 lo=float(taxis.find('m_fMinValue').get('val'));hi=float(taxis.find('m_fMaxValue').get('val'));blo=float(baxis.find('m_fMinValue').get('val'));bhi=float(baxis.find('m_fMaxValue').get('val'));source_intervals=baxis.findall('m_vecintvlOrdinal')[0].findall('elem');source_low=float(source_intervals[1].find('end').get('val'));source_high=float(source_intervals[0].find('begin').get('val'));a=lo+(source_low-blo)/(bhi-blo)*(hi-lo);b=lo+(source_high-blo)/(bhi-blo)*(hi-lo);need(lo<a<b<hi,'Scaled valid interval topology is out of bounds')
 vectors=[copy.deepcopy(x) for x in baxis.findall('m_vecintvlOrdinal')]
 vals=[(b,hi),(lo,a),(a,b),(-hi,-b),(-a,-lo),(-b,-a)]
 for v,triples in zip(vectors,(vals[:3],vals[3:])):
  for e,(x,y) in zip(v.findall('elem'),triples):e.find('begin').set('val',format(x,'.20E'));e.find('end').set('val',format(y,'.20E'))
 original_axis_order=[x.tag for x in taxis]
 prior_vectors=taxis.findall('m_vecintvlOrdinal');need(len(prior_vectors)==1,'ordinary axis must own one no-break interval vector')
 prior_breaks=taxis.findall('m_cdaxisbreak');need(len(prior_breaks)==1 and not prior_breaks[0].findall('elem'),'ordinary axis must own one empty break collection')
 vector_slot=list(taxis).index(prior_vectors[0]);break_slot=list(taxis).index(prior_breaks[0]);need(vector_slot < break_slot,'ordinary donor axis order is unexpected')
 taxis.replace(prior_vectors[0],vectors[0])
 ownerref=E.Element('m_cdaxisbreak',length='1');E.SubElement(ownerref,'elem',idref=newids[0])
 taxis.replace(prior_breaks[0],ownerref)
 expected_axis_order=original_axis_order
 need([x.tag for x in taxis]==expected_axis_order,'axis field order changed during replacement')
 for x in clones:root.append(x)
 ole=replace(tc['doc']['ole'],E.tostring(root,encoding='utf-8'));before=tc['doc']['streams'];after=streams(ole);need(set(before)==set(after) and all(before[k]==after[k] for k in before if k!=('think-cellXML',)),'Non-model OLE stream changed')
 with zipfile.ZipFile(io.BytesIO(traw)) as tz,zipfile.ZipFile(io.BytesIO(braw)) as bz:
  tslide=E.fromstring(tz.read(slide));tree=tslide.find('p:cSld/p:spTree',NS);rp,rr,rm=rels(tz,slide);brp,brr,brm=rels(bz,slide);updates={};nextid=max([int(x[3:]) for x in rm if x.startswith('rId') and x[3:].isdigit()]+[0])+1
  # Bind physical copies by their existing opaque tag, never slide traversal order.
  new_by_old=dict(zip(old,new));need(len(new_by_old)==6,'model shape tags are not unique')
  source_cnv_counts=collections.Counter(x.get('id') for x in tslide.findall('.//p:cNvPr',NS));existing_shape_ids={int(x) for x in source_cnv_counts if (x or '').isdigit()};next_shape_id=max(existing_shape_ids|{0})+1
  donor_rel_by_id={x.get('Id'):x for x in brr}
  for (node,oldname,oldrid,oldpart) in bparts:
   newname=new_by_old[oldname];clone=copy.deepcopy(node)
   # Every copied physical shape gets fresh cNvPr IDs in the target slide.
   for nv in clone.findall('.//p:cNvPr',NS):
    while next_shape_id in existing_shape_ids:next_shape_id+=1
    nv.set('id',str(next_shape_id));existing_shape_ids.add(next_shape_id);next_shape_id+=1
   # Copy every slide relationship referenced by this physical node.  The
   # known tags relation receives a fresh tag-part target; other dependencies
   # retain their donor relation metadata instead of becoming dangling rIds.
   for attr in clone.iter():
    for key,oldrid2 in list(attr.attrib.items()):
     if key!=Q('id') or oldrid2 not in donor_rel_by_id:continue
     donor_rel=donor_rel_by_id[oldrid2];nrid='rId'+str(nextid);nextid+=1
     if oldrid2==oldrid:
      newpart='ppt/tags/break-insertion-'+str(nextid)+'.xml';srcxml=E.fromstring(bz.read(oldpart));[x.set('val',newname) for x in srcxml if x.get('name','').upper()=='THINKCELLSHAPEDONOTDELETE'];updates[newpart]=E.tostring(srcxml,xml_declaration=True,encoding='UTF-8',standalone=True);target='../tags/'+Path(newpart).name
     else:
      target=donor_rel.get('Target')
     relattrs={k:v for k,v in donor_rel.attrib.items() if k!='Id'};relattrs['Id']=nrid;relattrs['Target']=target;E.SubElement(rr,'{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',**relattrs);attr.set(key,nrid)
   tree.append(clone)
  updates[slide]=E.tostring(tslide,xml_declaration=True,encoding='UTF-8',standalone=True);updates[rp]=E.tostring(rr,xml_declaration=True,encoding='UTF-8',standalone=True)
  # Append the copied tag parts.  If this package lacks an XML default, retain
  # the donor's actual declared content type instead of inventing one.
  ct='[Content_Types].xml';ctr=E.fromstring(tz.read(ct));donorct=E.fromstring(bz.read(ct));default=any(x.get('Extension')=='xml' for x in ctr)
  if not default:
   for p in updates:
    if p.startswith('ppt/tags/'):
     donor_override=next((x.get('ContentType') for x in donorct if x.get('PartName')=='/'+oldpart),None);donor_default=next((x.get('ContentType') for x in donorct if x.get('Extension')=='xml'),None);need(donor_override or donor_default,'donor tag content type is unavailable');E.SubElement(ctr,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',PartName='/'+p,ContentType=donor_override or donor_default)
   updates[ct]=E.tostring(ctr,xml_declaration=True,encoding='UTF-8',standalone=True)
  # custom ZIP write adds six new tag parts after copying all target entries
  with zipfile.ZipFile(out,'x') as d:
   d.comment=tz.comment
   for x in tz.infolist():d.writestr(copy.copy(x),ole if x.filename==tc['doc']['part'] else updates.get(x.filename,tz.read(x.filename)))
   for p,v in updates.items():
    if p not in tz.namelist():d.writestr(p,v)
 # Static ownership/reference readback.
 raw=out.read_bytes();_,cc,axis=one(out);ii=cc['doc']['ids'];ar=axis.findall('m_cdaxisbreak/elem');need(len(ar)==1 and ar[0].get('idref')==newids[0],'axis break ownership failed');nb=ii[newids[0]];need([x.get('idref') for x in nb.findall('m_cpptbreakshp/elem')]==newids[1:],'break shape references failed');got=[ii[i].find(f+'/m_bstrShapeName').text for i in newids[1:] for f in FIELDS];need(got==new,'model physical tags failed')
 with zipfile.ZipFile(io.BytesIO(raw)) as z:
  _,pp=parts(z,slide,new);need(len(pp)==6,'six physical parts missing')
  _,candidate_rels,candidate_relmap=rels(z,slide);candidate_slide=E.fromstring(z.read(slide));candidate_cnv_counts=collections.Counter(x.get('id') for x in candidate_slide.findall('.//p:cNvPr',NS));need(all(candidate_cnv_counts[k]==v for k,v in source_cnv_counts.items()) and len(candidate_cnv_counts)==len(source_cnv_counts)+6,'copied physical shapes collided with target cNvPr IDs')
  referenced_rids=[v for node in candidate_slide.iter() for k,v in node.attrib.items() if k==Q('id')];need(all(v in candidate_relmap for v in referenced_rids),'copied physical shape has a dangling slide relationship')
  for rel in candidate_rels:
   if rel.get('TargetMode')!='External':need(candidate_relmap[rel.get('Id')] in z.namelist(),'slide relationship target is missing from package')
 actual_axis=[x.tag for x in axis]
 need(actual_axis==expected_axis_order,'candidate axis child sequence differs from native ordinary donor')
 need(sha(Path(ordinary).read_bytes()).upper()==expected_ordinary_sha256.upper(),'ordinary source changed during preparation')
 need(sha(Path(broken).read_bytes()).upper()==expected_donor_sha256.upper(),'break donor changed during preparation')
 result={'status':'STATIC_REAL_GRAPH_ORDERED_INSERTION_PREPARED_REGENERATION_REQUIRED','ordinary_source':str(Path(ordinary).resolve()),'ordinary_source_sha256':sha(traw),'break_donor':str(Path(broken).resolve()),'break_donor_sha256':sha(braw),'candidate_sha256':sha(raw),'native_graph':{'break_id':newids[0],'shape_ids':newids[1:],'physical_tags':new},'axis_invariants':{'target_prior_break_count':0,'target_prior_interval_cardinality':prior_interval_cardinality,'inserted_break_count':1,'inserted_interval_versions':[{'reqver':x.get('reqver'),'endver':x.get('endver'),'length':x.get('length')} for x in vectors],'scaled_interval_endpoints':[a,b],'axis_range':[lo,hi],'axis_child_tag_sequence_preserved':True,'axis_child_tag_sequence':actual_axis,'replaced_slots':{'m_vecintvlOrdinal':vector_slot,'m_cdaxisbreak':break_slot}},'data_preserved_before_native':True,'physical_part_count':6,'source_and_donor_unchanged':True,'required_followup':['official JSON regeneration','native reopen/render','break graph and exact datasource readback','changed-data repeat']};report.write_text(json.dumps(result,indent=2));print(json.dumps(result,indent=2))
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--ordinary',type=Path,required=True);p.add_argument('--expected-ordinary-sha256',required=True);p.add_argument('--break-donor',type=Path,required=True);p.add_argument('--expected-donor-sha256',required=True);p.add_argument('--output',type=Path,required=True);p.add_argument('--report',type=Path,required=True);a=p.parse_args();need(not a.output.exists() and not a.report.exists(),'new outputs required');build(a.ordinary,a.break_donor,a.expected_ordinary_sha256,a.expected_donor_sha256,a.output,a.report)
