"""Internal graph composition adapter; use compose_sequence_charts.py for verified output.

This low-level module writes an intermediate candidate, never a delivery artifact.
"""
from __future__ import annotations
import argparse, copy, hashlib, io, json, shutil, subprocess, sys, tempfile, zipfile
import uuid
from pathlib import Path
from lxml import etree as E

ROOT=Path(__file__).resolve().parents[1]
IMPL=ROOT/'thinkcell_no_click'/'implementation'
sys.path[:0]=[str(IMPL),str(IMPL.parent.parent)]
from prepare_thinkcell_name import inventory, link_contract, logical_slides, need, sha, streams, xml, NS, relationship_map
from runtime import powershell, powershell_env

def fail(x): raise ValueError(x)
def ids(root):
 d={x.get('id'):x for x in root if x.get('id')}
 if len(d)!=sum(x.get('id') is not None for x in root): fail('duplicate or absent top-level IDs')
 return d
def child(root,tag):
 a=root.findall(tag)
 if len(a)!=1: fail('expected one '+tag)
 return a[0]
def model_keys(root, field):
 """Return the model-wide identity values for a concrete serialized field.

 The graph has several identity domains in addition to top-level ``id``.  A
 donor made from the same seed can carry equal UUIDs and hidden shape names
 even when every copied PowerPoint tag has been made unique.
 """
 out=[]
 for node in root.iter():
  if node.tag!=field: continue
  value=(node.text or node.get('val') or '').strip()
  if value: out.append(value)
 return out
def unique_model_keys(root, field, label):
 values=model_keys(root,field)
 duplicates=sorted(k for k in set(values) if values.count(k)!=1)
 if duplicates: fail(label+' is already non-unique: '+','.join(duplicates[:3]))
 return set(values)
def fresh_guid(old, occupied):
 """Stable, valid UUID replacement with an explicit collision retry."""
 serial=0
 while True:
  value=str(uuid.uuid5(uuid.NAMESPACE_URL,'think-cell-merge|'+old+'|'+str(serial)))
  if value not in occupied:return value
  serial+=1
def identity_maps(receiver_root, donor_root, receiver_tags):
 """Create one-to-one remaps for every known global model identity domain."""
 receiver_shapes=unique_model_keys(receiver_root,'m_bstrShapeName','receiver shape-name identity')
 donor_shapes=unique_model_keys(donor_root,'m_bstrShapeName','donor shape-name identity')
 receiver_guids=unique_model_keys(receiver_root,'m_guid','receiver GUID identity')
 donor_guids=unique_model_keys(donor_root,'m_guid','donor GUID identity')
 if any(str(uuid.UUID(v))!=v.lower() for v in receiver_guids|donor_guids): fail('model GUID identity is not a canonical UUID')
 shapes={}; occupied=set(receiver_tags)|receiver_shapes
 for old in sorted(donor_shapes & occupied):
  value=fresh_tag(old,occupied|set(shapes.values()))
  shapes[old]=value;occupied.add(value)
 guids={}; occupied=set(receiver_guids)
 for old in sorted(donor_guids & occupied):
  value=fresh_guid(old,occupied|set(guids.values()))
  guids[old]=value;occupied.add(value)
 if len(set(shapes.values()))!=len(shapes) or len(set(guids.values()))!=len(guids): fail('identity remap is not one-to-one')
 return shapes,guids
def substitute_identities(node, maps):
 """Rewrite exact serialized identity references, including attributes."""
 for x in node.iter():
  for mapping in maps:
   if (x.text or '') in mapping:x.text=mapping[x.text]
   for key,value in list(x.attrib.items()):
    if value in mapping:x.set(key,mapping[value])
def require_identity_closure(doc, donor_parts, maps):
 """Fail closed if a remapped identity occurs outside model/XML closure.

 Datasheets are copied byte-for-byte, so an identity occurring there would
 require a format-aware writer.  The observed donors contain GUIDs only in
 think-cellXML; enforce that instead of silently breaking a reference.
 """
 values=set().union(*[set(m) for m in maps])
 for path,data in doc['streams'].items():
  if path==('think-cellXML',): continue
  if any(v.encode('utf-8') in data for v in values): fail('identity reference outside model stream: '+'/'.join(path))
 for part,data in donor_parts.items():
  # The donor carrier is deliberately not cloned. Its child storage is copied
  # separately and has already been checked above.
  if part==doc['part']: continue
  if part.endswith(('.xml','.rels')): continue
  if any(v.encode('utf-8') in data for v in values): fail('identity reference in opaque OOXML part: '+part)
def active(c, donor=False):
 if c['owner'].tag!='CSequenceChartSE' or not c['exact']: fail('only exact sequence donors')
 link_contract(c)
 d=c['doc']; r=d['root']; tab=c['table']; sn=tab.find('m_bstrRangeName'); storage='' if sn is None else (sn.text or sn.get('val') or '')
 if d['root'].get('reqver')!='32687': fail('unexpected think-cell model version')
 if c['owner_name']!=c['table_name'] or not c['owner_name']: fail('donor needs one consistent existing automation name')
 if donor and storage!='think-cellChild0': fail('this experiment expects a single-child donor')
 if any(x.get('idref') not in (None,'0') and x.get('idref') not in ids(r) for x in r.iter()): fail('unresolved model idref')
 return d
def remap_foreign(receiver, donor, child_name, identity_maps, new_name):
 rr=receiver['root']; dr=donor['root']; ri=ids(rr); di=ids(dr)
 if child(rr,'CSmartGrid').get('id')!='1' or child(rr,'CContainerSE').get('id')!='2': fail('receiver root unsupported')
 if child(dr,'CSmartGrid').get('id')!='1' or child(dr,'CContainerSE').get('id')!='2': fail('donor root unsupported')
 # The only allowed global nodes are root smartgrid/container. Everything else is a complete foreign graph.
 if any(x.tag in {'CSmartGrid','CContainerSE'} for x in dr if x.get('id') not in {'1','2'}): fail('unknown global model node')
 offset=max(int(k) for k in ri)+1; mapping={str(k):str(int(k)+offset) for k in di if k not in {'1','2'}}
 foreign=[]
 for n in dr:
  if n.get('id') is None or n.get('id') in {'1','2'}: continue
  q=copy.deepcopy(n); q.set('id',mapping[n.get('id')])
  for x in q.iter():
   if x.get('idref') not in (None,'0'):
    old=x.get('idref')
    if old not in mapping: fail('foreign idref reaches excluded global id '+old)
    x.set('idref',mapping[old])
  if q.tag=='CSequenceChartDataTable':
   rng=q.find('m_bstrRangeName')
   if rng is None or (rng.text or rng.get('val') or '')!='think-cellChild0': fail('foreign datasource storage is not child0')
   if rng.get('val') is not None:rng.set('val',child_name)
   else:rng.text=child_name
   n=q.find('m_strName')
   if n is not None and (n.text or n.get('val') or '')==donor['name']:
    if n.get('val') is not None:n.set('val',new_name)
    else:n.text=new_name
  if q.tag=='CSequenceChartSE':
   n=q.find('m_strName')
   if n is not None and (n.text or n.get('val') or '')==donor['name']:
    if n.get('val') is not None:n.set('val',new_name)
    else:n.text=new_name
  substitute_identities(q,identity_maps)
  foreign.append(q)
 # foreign chart/legend slots are the only allowed container contribution.
 fc=child(dr,'CContainerSE'); slots=fc.find('m_cse')
 rc=child(rr,'CContainerSE'); rslots=rc.find('m_cse')
 if slots is None or rslots is None: fail('missing container cse slots')
 refs=[x.get('idref') for x in slots.findall('elem') if x.get('idref') not in (None,'0')]
 if not refs or any(x not in mapping for x in refs): fail('container has unknown foreign cse refs')
 # Preserve receiver slots and append only foreign owned chart/legend objects.
 for ref in refs:
  q=E.Element('elem');q.set('idref',mapping[ref]);rslots.append(q)
 rslots.set('length',str(len(rslots.findall('elem'))))
 for q in foreign: rr.append(q)
 # Exact closure / uniqueness after graft.
 now=ids(rr)
 if len(now)!=len(set(now)): fail('merge produced duplicate model IDs')
 unresolved=sorted({x.get('idref') for x in rr.iter() if x.get('idref') not in (None,'0') and x.get('idref') not in now})
 if unresolved: fail('unresolved refs after merge: '+','.join(unresolved))
 return mapping, E.tostring(rr,encoding='utf-8')
def relname(part):
 p=Path(part); return str(p.parent/'_rels'/(p.name+'.rels')).replace('\\','/')
def resolve(source,target):
 import posixpath
 return posixpath.normpath(posixpath.join(posixpath.dirname(source),target)).lstrip('/')
def relative(source,target):
 import posixpath
 return posixpath.relpath(target,posixpath.dirname(source))
def next_part(parts,source):
 p=Path(source); n=1
 while True:
  q=str(p.parent/(p.stem+f'_merge{n}'+p.suffix)).replace('\\','/')
  if q not in parts:return q
  n+=1
def tagged_slide_parts(z,part):
 root=xml(z.read(part)); rels=relationship_map(z,part); out=[]
 for node in root.findall('.//p:spTree/*',NS):
  if node.find('.//p:oleObj',NS) is not None: continue
  tagged=False
  for t in node.findall('.//p:tags',NS):
   rid=t.get('{%s}id'%NS['r']); target=rels.get(rid,{}).get('resolved')
   if target and any(x.get('name','').upper()=='THINKCELLSHAPEDONOTDELETE' for x in xml(z.read(target))): tagged=True
  if tagged: out.append(node)
 return root,out
def clone_part(source_parts, dest_parts, source_content, dest_content, source, memo, identity_maps):
 if source in memo:return memo[source]
 if source not in source_parts:fail('OOXML relationship target missing '+source)
 dest=next_part(dest_parts,source);memo[source]=dest;data=source_parts[source]
 if source.endswith(('.xml','.rels')):
  root=xml(data)
  substitute_identities(root,identity_maps)
  data=E.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
 dest_parts[dest]=data
 ct=source_content.xpath('//*[local-name()="Override" and @PartName=$n]',n='/'+source)
 if ct:
  q=copy.deepcopy(ct[0]);q.set('PartName','/'+dest);dest_content.append(q)
 sr=relname(source)
 if sr in source_parts:
  root=xml(source_parts[sr]); newrels=E.Element(root.tag,nsmap=root.nsmap)
  for r in root:
   q=copy.deepcopy(r); target=q.get('Target')
   if q.get('TargetMode')=='External': fail('external OOXML relationship in owned closure: '+source+' -> '+target)
   else:
    original=resolve(source,target)
    q.set('Target',relative(dest,clone_part(source_parts,dest_parts,source_content,dest_content,original,memo,identity_maps)))
   newrels.append(q)
  dest_parts[relname(dest)]=E.tostring(newrels,xml_declaration=True,encoding='UTF-8',standalone=True)
 return dest
def transplant_render_owner_tree(receiver_parts, donor_parts, receiver_slide, donor_slide, identity_maps):
 content=xml(receiver_parts['[Content_Types].xml']); source_content=xml(donor_parts['[Content_Types].xml']); rr=xml(receiver_parts[relname(receiver_slide)]); dr=xml(donor_parts[relname(donor_slide)])
 used={x.get('Id') for x in rr}; nextid=1
 def newrid():
  nonlocal nextid
  while 'rId'+str(nextid) in used:nextid+=1
  x='rId'+str(nextid);used.add(x);nextid+=1;return x
 donor_rels={x.get('Id'):x for x in dr}; memo={}; target_root, shapes=tagged_slide_parts(_MemoryZip(donor_parts),donor_slide)
 recvroot,_=tagged_slide_parts(_MemoryZip(receiver_parts),receiver_slide); tree=recvroot.find('.//p:spTree',NS)
 maxshape=max([int(x.get('id')) for x in recvroot.findall('.//p:cNvPr',NS) if x.get('id','').isdigit()] or [0])
 copied=0
 for node in shapes:
  q=copy.deepcopy(node)
  # Tagged owner-tree XML is copied directly rather than through clone_part.
  # Apply the same exact-value identity substitution before its relationships
  # are rewired so an identity reference in a shape property cannot survive.
  substitute_identities(q,identity_maps)
  for x in q.findall('.//*[@{%s}id]'%NS['r']):
   old=x.get('{%s}id'%NS['r']); rel=donor_rels.get(old)
   if rel is None:fail('shape relationship missing '+old)
   nr=copy.deepcopy(rel);nr.set('Id',newrid())
   if nr.get('TargetMode')=='External': fail('external slide relationship in owned closure: '+old)
   nr.set('Target',relative(receiver_slide,clone_part(donor_parts,receiver_parts,source_content,content,resolve(donor_slide,rel.get('Target')),memo,identity_maps)))
   rr.append(nr);x.set('{%s}id'%NS['r'],nr.get('Id'))
  for cn in q.findall('.//p:cNvPr',NS):
   maxshape+=1;cn.set('id',str(maxshape));cn.set('name','Merged '+cn.get('name','native'))
  tree.append(q);copied+=1
 receiver_parts[receiver_slide]=E.tostring(recvroot,xml_declaration=True,encoding='UTF-8',standalone=True);receiver_parts[relname(receiver_slide)]=E.tostring(rr,xml_declaration=True,encoding='UTF-8',standalone=True);receiver_parts['[Content_Types].xml']=E.tostring(content,xml_declaration=True,encoding='UTF-8',standalone=True)
 return copied,memo
class _MemoryZip:
 def __init__(self,parts):self.parts=parts
 def read(self,n):return self.parts[n]
 def namelist(self):return list(self.parts)
def native_tag_values(parts,slide):
 root,shapes=tagged_slide_parts(_MemoryZip(parts),slide); rels=relationship_map(_MemoryZip(parts),slide); tags=[]
 for node in shapes:
  for t in node.findall('.//p:tags',NS):
   target=rels[t.get('{%s}id'%NS['r'])]['resolved']
   for x in xml(parts[target]):
    if x.get('name','').upper()=='THINKCELLSHAPEDONOTDELETE':tags.append(x.get('val'))
 return tags
def all_native_tags(parts,slide): return set(native_tag_values(parts,slide))
def fresh_tag(old, occupied):
 import base64
 i=0
 while True:
  h=hashlib.sha256((old+'|'+str(i)).encode()).digest(); v='t'+base64.urlsafe_b64encode(h).decode().rstrip('=')[:22]
  if v not in occupied:return v
  i+=1
def render_tag_map(receiver_tags, donor_tags, shape_map):
 """Remap copied render-only carrier tags outside the think-cell model graph.

 A donor can have a tagged anchor whose value does not appear in
 ``m_bstrShapeName``. It is copied in OOXML with its tag part, so model-only
 identity remapping cannot prevent a duplicate live PowerPoint tag.
 """
 if len(donor_tags)!=len(set(donor_tags)): fail('donor render tags are not unique')
 occupied=set(receiver_tags)|set(donor_tags)|set(shape_map.values()); mapped={}
 for old in sorted(set(donor_tags)&set(receiver_tags)):
  mapped[old]=shape_map[old] if old in shape_map else fresh_tag(old,occupied|set(mapped.values()))
  occupied.add(mapped[old])
 if len(set(mapped.values()))!=len(mapped): fail('render-tag remap is not one-to-one')
 return mapped
def slide_theme(parts, slide):
 """Resolve the theme actually inherited by this slide through layout and master."""
 z=_MemoryZip(parts); current=slide
 for relationship_type in ('slideLayout','slideMaster','theme'):
  rels=relationship_map(z,current); hits=[r['resolved'] for r in rels.values() if r.get('type','').endswith('/'+relationship_type)]
  if len(hits)!=1: fail('cannot resolve one '+relationship_type+' for '+current)
  current=hits[0]
 return current,parts[current]
def theme_signature(data):
 root=xml(data); ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
 nodes=root.findall('.//a:clrScheme',ns)+root.findall('.//a:fontScheme',ns)+root.findall('.//a:fmtScheme',ns)
 if len(nodes)!=3: fail('theme lacks one color, font, and format scheme')
 return tuple(E.tostring(x,method='c14n') for x in nodes)
def merge(receiver_path,donor_path,out):
 rb=receiver_path.read_bytes();db=donor_path.read_bytes(); rd,rc,_=inventory(rb); dd,dc,_=inventory(db)
 rslides=logical_slides(zipfile.ZipFile(io.BytesIO(rb))); dslides=logical_slides(zipfile.ZipFile(io.BytesIO(db)))
 if len(rslides)!=1 or len(dslides)!=1: fail('requires exactly one logical slide in each source')
 receiver_slide=rslides[0]['part']; donor_slide=dslides[0]['part']
 receiver=rd[0] if len(rd)==1 else None; donor_doc=active(dc[0],True) if len(dc)==1 else None
 if receiver is None or donor_doc is None: fail('requires one receiver carrier and one donor carrier/chart')
 with zipfile.ZipFile(io.BytesIO(rb)) as rz, zipfile.ZipFile(io.BytesIO(db)) as dz:
  rs=xml(rz.read('ppt/presentation.xml')).find('p:sldSz',NS); ds=xml(dz.read('ppt/presentation.xml')).find('p:sldSz',NS)
  if rs is None or ds is None or (rs.get('cx'),rs.get('cy'))!=(ds.get('cx'),ds.get('cy')): fail('slide dimensions differ')
  receiver_parts0={n:rz.read(n) for n in rz.namelist()}; donor_parts0={n:dz.read(n) for n in dz.namelist()}
  if theme_signature(slide_theme(receiver_parts0,receiver_slide)[1])!=theme_signature(slide_theme(donor_parts0,donor_slide)[1]): fail('inherited slide color/font schemes differ')
 if len(rd)!=1 or len(rc)<1 or len(dd)!=1 or len(dc)!=1: fail('requires one receiver carrier with one or more charts and one donor carrier/chart')
 donor=dict(donor_doc); donor['name']=dc[0]['owner_name']
 if receiver['root'].get('reqver')!=donor['root'].get('reqver'): fail('think-cell model versions differ')
 for c in rc: active(c); link_contract(c)
 names=[c['owner_name'].casefold() for c in rc]; new_name=donor['name']
 if new_name.casefold() in names:
  base=f'TC_AUTO_{sha(db)[:12]}_{donor_doc["slide_id"]}_{dc[0]["owner"].get("id")}';new_name=base; serial=2
  while new_name.casefold() in names:new_name=f'{base}_{serial}';serial+=1
 stores={link_contract(c)['storage'] for c in rc}; i=0
 while f'think-cellChild{i}' in stores:i+=1
 child_name=f'think-cellChild{i}'
 receiver_tag_values=native_tag_values(receiver_parts0,receiver_slide)
 if len(receiver_tag_values)!=len(set(receiver_tag_values)): fail('receiver render tags are not unique')
 receiver_tags=set(receiver_tag_values)
 shape_map,guid_map=identity_maps(receiver['root'],donor['root'],receiver_tags)
 donor_tags=native_tag_values(donor_parts0,donor_slide)
 tag_map=render_tag_map(receiver_tags,donor_tags,shape_map)
 identity_map=(shape_map,guid_map,tag_map)
 require_identity_closure(donor,donor_parts0,identity_map)
 mapping,model=remap_foreign(receiver,donor,child_name,identity_map,new_name)
 # Copy every tagged donor render shape and its relationship closure, excluding the donor carrier.
 with tempfile.TemporaryDirectory(dir=str(out.parent),prefix='tc_merge_') as td:
  td=Path(td); carrier=td/'receiver.bin'; donorbin=td/'donor.bin'; modelp=td/'model.xml'; carrier.write_bytes(receiver['ole']);donorbin.write_bytes(donor['ole']);modelp.write_bytes(model)
  ps=Path(__file__).with_name('copy_cfb_child.ps1')
  p=subprocess.run([powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(ps),'-Receiver',str(carrier),'-Donor',str(donorbin),'-DonorChild','think-cellChild0','-ReceiverChild',child_name,'-ModelBytes',str(modelp)],capture_output=True,text=True,env=powershell_env(),creationflags=subprocess.CREATE_NO_WINDOW)
  if p.returncode: fail('CFB child copy rejected: '+p.stderr[-1000:])
  changed=carrier.read_bytes(); ss=streams(changed)
  expected=set(receiver['streams'])|{(child_name,'Package'),(child_name,'think-cellXML')}
  if set(ss)!=expected: fail('unexpected CFB stream inventory')
  with zipfile.ZipFile(io.BytesIO(rb)) as z: parts={n:z.read(n) for n in z.namelist()}
  with zipfile.ZipFile(io.BytesIO(db)) as z: donor_parts={n:z.read(n) for n in z.namelist()}
  parts[receiver['part']]=changed
  copied,memo=transplant_render_owner_tree(parts,donor_parts,receiver_slide,donor_slide,identity_map)
  if copied==0: fail('no donor tagged render shapes copied')
  merged_tags=native_tag_values(parts,receiver_slide)
  if len(merged_tags)!=len(set(merged_tags)): fail('merged live render tags are not globally unique')
  if out.exists():fail('output exists')
  with zipfile.ZipFile(out,'x',zipfile.ZIP_DEFLATED) as z:
   for n,v in parts.items():z.writestr(n,v)
  if zipfile.ZipFile(out).testzip() is not None:fail('output ZIP CRC failure')
  docs,candidates,names=inventory(out.read_bytes())
  if len(docs)!=1 or len(candidates)!=len(rc)+1:fail('merged package does not expose one carrier and every original chart plus donor')
  if len({c['frames'][0]['shape_tag'] for c in candidates})!=len(candidates):fail('merged chart tags are not globally unique')
  final_root=docs[0]['root']
  final_shapes=unique_model_keys(final_root,'m_bstrShapeName','merged shape-name identity')
  final_guids=unique_model_keys(final_root,'m_guid','merged GUID identity')
  report={'status':'RESEARCH_CANDIDATE_ONLY','receiver_sha256':sha(rb),'donor_sha256':sha(db),'output_sha256':sha(out.read_bytes()),'model_id_offset':min(int(v) for v in mapping.values()),'receiver_chart_count_before':len(rc),'chart_count_after':len(candidates),'new_datasource_storage':child_name,'automation_name_adapter':{'old':donor['name'],'new':new_name,'changed':new_name!=donor['name']},'tag_bijection':shape_map,'guid_bijection':guid_map,'render_tag_bijection':tag_map,'copied_tagged_render_shapes':copied,'cloned_dependency_parts':memo,'global_identity_invariants':{'shape_names_unique':True,'guids_unique':True,'live_render_tags_unique':True,'shape_name_count':len(final_shapes),'guid_count':len(final_guids),'live_render_tag_count':len(merged_tags)},'charts':[{'name':c['owner_name'],'tag':c['frames'][0]['shape_tag'],'storage':link_contract(c)['storage']} for c in candidates]}
  print(json.dumps(report))
def main():
 p=argparse.ArgumentParser();p.add_argument('--receiver',type=Path,required=True);p.add_argument('--donor',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
 try:
  if a.output.exists(): fail('output must not exist')
  merge(a.receiver.resolve(),a.donor.resolve(),a.output.resolve())
 except Exception as e: print(json.dumps({'status':'REJECTED_OR_STOPPED','error':str(e)}));return 1
 return 0
if __name__=='__main__': raise SystemExit(main())
