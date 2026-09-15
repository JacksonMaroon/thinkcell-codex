"""Guarded, experimental offline think-cell naming on distinct local copies.
Dry run is default. No Office processes, data, or geometry are changed.
"""
from __future__ import annotations
import argparse,copy,hashlib,io,json,re,subprocess,sys,tempfile,zipfile
from pathlib import Path
from lxml import etree as E
SKILL=Path(__file__).resolve().parents[3]
sys.path.insert(0,str(SKILL/'scripts'))
import olefile
from runtime import powershell, powershell_env
from audit_thinkcell_integrity import logical_slides,relationship_map,NS,pie_tables,sequence_tables,scatter_tables,inspect_presentation
SUPPORTED={'CPieChartSE':'CPieChartDataTable','CSequenceChartSE':'CSequenceChartDataTable','CScatterChartSE':'CScatterChartDataTable'}
class PreparationError(ValueError):pass
def need(ok,message):
 if not ok:raise PreparationError(message)
def sha(b):return hashlib.sha256(b).hexdigest().upper()
def xml(b):return E.fromstring(b,E.XMLParser(remove_blank_text=False,resolve_entities=False))
def value(n):return '' if n is None else (n.text or n.get('val') or '')
def streams(b):
 with olefile.OleFileIO(io.BytesIO(b)) as cf:return {tuple(p):cf.openstream(p).read() for p in cf.listdir(streams=True,storages=False)}
def only(parent,tag,optional=False):
 a=parent.findall(tag);need(len(a)==1 or (optional and not a),'Missing or repeated '+tag)
 return a[0] if a else None

def inventory(data):
 docs=[];names=[];candidates=[];part_owners={}
 with zipfile.ZipFile(io.BytesIO(data)) as z:
  need(len(set(z.namelist()))==len(z.namelist()),'Duplicate ZIP entries')
  for sn,slide in enumerate(logical_slides(z),1):
   sr=xml(z.read(slide['part']));rels=relationship_map(z,slide['part']);tagged={}
   for f in sr.findall('.//p:graphicFrame',NS)+sr.findall('.//p:sp',NS)+sr.findall('.//p:cxnSp',NS):
    chart=f.find('.//c:chart',NS)
    cn=f.find('.//p:cNvPr',NS);xf=f.find('p:xfrm',NS)
    if xf is None:xf=f.find('p:spPr/a:xfrm',NS)
    if cn is None or xf is None:continue
    for tags in f.findall('.//p:tags',NS):
     for tag in xml(z.read(rels[tags.get('{'+NS['r']+'}id')]['resolved'])):
      if tag.get('name','').upper()=='THINKCELLSHAPEDONOTDELETE':
       record={'shape_id':int(cn.get('id')),'shape_name':cn.get('name'),'shape_tag':tag.get('val'),'native_chart_part':rels[chart.get('{'+NS['r']+'}id')]['resolved'] if chart is not None else None,'bounds_emu':{'left':int(xf.find('a:off',NS).get('x')),'top':int(xf.find('a:off',NS).get('y')),'width':int(xf.find('a:ext',NS).get('cx')),'height':int(xf.find('a:ext',NS).get('cy'))}}
       tagged.setdefault(tag.get('val'),[]).append(record)
   parts={rels[o.get('{'+NS['r']+'}id')]['resolved'] for o in sr.findall('.//p:oleObj',NS) if o.get('progId','').startswith('TCLayout.ActiveDocument')}
   for part in sorted(parts):
    need(part not in part_owners,'Active document shared across slides')
    part_owners[part]=sn;ole=z.read(part);ss=streams(ole);need(('think-cellXML',) in ss,'Missing model stream')
    root=xml(ss[('think-cellXML',)]);ids={n.get('id'):n for n in root if n.get('id')}
    need(len(ids)==sum(bool(n.get('id')) for n in root),'Duplicate model IDs')
    d={'slide_number':sn,'slide_id':slide['id'],'part':part,'ole':ole,'streams':ss,'root':root,'ids':ids};docs.append(d)
    paired_tables=set()
    for owner in root:
     if owner.tag not in SUPPORTED:continue
     tr=only(owner,'m_dtable');table=ids.get(tr.get('idref'));need(table is not None and table.tag==SUPPORTED[owner.tag],'Wrong table owner')
     paired_tables.add(table.get('id'))
     if owner.tag=='CPieChartSE':
      refs=table.findall('ocol/elem');tags=[ids[r.get('idref')].findtext('m_pptpiechart/m_bstrShapeName') for r in refs if r.get('idref') in ids]
     elif owner.tag=='CScatterChartSE':
      tags=[owner.findtext('m_pptscatchart/m_bstrShapeName')]
     else:
      ptr=owner.find('m_pptseqchart');native=ids.get(ptr.get('idref')) if ptr is not None else None;tags=[native.findtext('m_bstrShapeName')] if native is not None else []
      if not tags and owner.find('m_ect') is not None and owner.find('m_ect').get('val') in {'11','12'}:
       # Mekko renders as native shapes. Resolve a rectangle owned by an
       # actual scalar in this table, never a nearby or guessed shape.
       for vr in table.findall('ocol/elem'):
        for sr0 in ids[vr.get('idref')].findall('ocol/elem'):
         sc=ids[sr0.get('idref')];rp=sc.find('m_pptrect');rect=ids.get(rp.get('idref')) if rp is not None else None
         tag=rect.findtext('m_bstrShapeName') if rect is not None else None
         if tag and len(tagged.get(tag,[]))==1:tags=[tag];break
        if tags:break
     records=[r for t in tags for r in tagged.get(t,[])]
     on=only(owner,'m_strName',True);tn=only(table,'m_strName',True)
     need(value(on)==value(tn),'Chart and data-table names disagree in generation scope')
     candidate={'doc':d,'owner':owner,'table':table,'owner_name':value(on),'table_name':value(tn),'tags':tags,'frames':records,'exact':len(tags)==1 and bool(tags[0]) and len(tagged.get(tags[0],[]))==1}
     candidates.append(candidate)
    # Full presentation generation scope: collapse only known chart/table pairs.
    for n in root:
     if n.get('id') in paired_tables:continue
     name=only(n,'m_strName',True)
     if value(name):names.append({'name':value(name),'slide_id':slide['id'],'part':part,'owner_id':n.get('id'),'owner_type':n.tag})
 return docs,candidates,names

def choose(candidates,slide_id=None,slide_number=None,shape_id=None,shape_tag=None):
 need(slide_id is not None or slide_number is not None,'Require slide ID or number')
 need(shape_id is not None or shape_tag is not None,'Require native chart shape ID or tag')
 matches=[]
 for c in candidates:
  d=c['doc']
  if slide_id is not None and d['slide_id']!=slide_id:continue
  if slide_number is not None and d['slide_number']!=slide_number:continue
  if not c['exact']:continue
  f=c['frames'][0]
  if shape_id is not None and f['shape_id']!=shape_id:continue
  if shape_tag is not None and f['shape_tag']!=shape_tag:continue
  matches.append(c)
 need(len(matches)==1,'Target is missing, ambiguous, unsupported, or selectors disagree')
 return matches[0]

def link_contract(c):
 t=c['table'];d=c['doc'];sink=only(t,'m_advisesink',True)
 need(sink is not None and sink.get('idref') is not None,'Unknown link state: missing advisesink')
 rid=sink.get('idref')
 if rid!='0':
  node=d['ids'].get(rid)
  kind=node.tag if node is not None else 'unresolved'
  raise PreparationError('Persistent or unknown external link: m_advisesink '+rid+' -> '+kind)
 for n in list(c['owner'].iter())+list(t.iter()):
  if n.tag in {'m_guidLink','m_lnkid','m_vecbMoniker'} and value(n):raise PreparationError('Unexpected external-link metadata')
 storage=value(only(t,'m_bstrRangeName',True));need(re.fullmatch(r'think-cellChild\d+',storage) is not None,'Unknown embedded datasource storage')
 children={p:v for p,v in d['streams'].items() if len(p)>1 and p[0]==storage}
 need(bool(children),'Embedded datasource missing')
 supported=any((p[-1] in {'Workbook','Book'} and v[:2]==b'\x09\x08') or v.startswith(b'PK\x03\x04') for p,v in children.items())
 need(supported,'Unknown embedded datasource format')
 return {'state':'INTERNAL_DATASHEET_OBSERVED','advisesink_idref':'0','storage':storage,'external_storage_field':value(only(t,'m_bExternalStorage',True)),'external_storage_field_is_not_a_link_indicator':True,'stream_sha256':{'/'.join(p[1:]):sha(v) for p,v in children.items()},'matrix_orientation':'Read actual embedded datasource; not inferred from chart direction','requires_exact_datasheet_matrix_before_json_update':True}

def name_plan(c,names,digest,requested=None):
 on=only(c['owner'],'m_strName');tn=only(c['table'],'m_strName',True)
 owner_name=value(on);table_name=value(tn)
 need(owner_name==table_name,'Chart and data-table names disagree')
 grouped={}
 for n in names:grouped.setdefault(n['name'].casefold(),[]).append(n)
 need(not any(len(v)>1 for v in grouped.values()),'Duplicate automation names within full presentation generation scope')
 chosen=owner_name or requested or f"TC_AUTO_{digest[:12]}_{c['doc']['slide_id']}_{c['owner'].get('id')}"
 need(bool(re.fullmatch(r'[\x20-\x7E]+',chosen)) and bool(chosen.strip()),'Automation name must be nonempty printable ASCII')
 need(not owner_name or not requested or requested==owner_name,'Existing unique name is reused; renaming is not allowed')
 if not owner_name:need(chosen.casefold() not in grouped,'Requested/generated automation name already exists')
 if tn is None:only(c['table'],'m_bExcelOnTop')
 return chosen,bool(owner_name)

def rewrite(c,name,reused):
 before=c['doc']['streams'][('think-cellXML',)]
 if reused:return before,[]
 root=xml(before);ids={n.get('id'):n for n in root if n.get('id')};owner=ids[c['owner'].get('id')];table=ids[c['table'].get('id')]
 on=only(owner,'m_strName');tn=only(table,'m_strName',True);on.text=name
 changes=[{'owner_type':owner.tag,'owner_id':owner.get('id'),'field':'m_strName','operation':'set-empty-text'}]
 if tn is None:
  tn=E.Element('m_strName');table.insert(table.index(only(table,'m_bExcelOnTop')),tn);op='insert-before-m_bExcelOnTop'
 else:op='set-empty-text'
 tn.text=name;changes.append({'owner_type':table.tag,'owner_id':table.get('id'),'field':'m_strName','operation':op})
 after=E.tostring(root,encoding='utf-8');verify_diff(c,after,name,reused)
 return after,changes

def verify_diff(c,after,name,reused):
 before=c['doc']['streams'][('think-cellXML',)]
 if reused:need(before==after,'Reuse changed model');return
 original=xml(before);root=xml(after);ids={n.get('id'):n for n in root if n.get('id')};owner=ids[c['owner'].get('id')];table=ids[c['table'].get('id')]
 on=only(owner,'m_strName');tn=only(table,'m_strName');need(value(on)==name and value(tn)==name,'Names not persisted')
 on.text=c['owner'].find('m_strName').text
 old=only(c['table'],'m_strName',True)
 if old is None:
  need(not tn.attrib and len(tn)==0 and tn.tail is None,'Unexpected inserted field')
  need(table.index(tn)+1==table.index(only(table,'m_bExcelOnTop')),'Wrong insertion position')
  table.remove(tn)
 else:tn.text=old.text
 need(E.tostring(original,method='c14n')==E.tostring(root,method='c14n'),'Non-name XML changed')

def logical_tables(c):
 r=c['doc']['root'];family=c['owner'].tag
 if family=='CPieChartSE':return pie_tables(r)
 if family=='CSequenceChartSE':return sequence_tables(r)
 if family=='CScatterChartSE':return scatter_tables(r)
 raise PreparationError('Unsupported logical table')

def prepare(input_path,expected_sha256,output_path,slide_id=None,slide_number=None,shape_id=None,shape_tag=None,name=None,execute=False):
 src=Path(input_path).resolve();out=Path(output_path).resolve();need(src!=out,'Input and output must differ')
 need(not str(out).startswith('\\\\'),'Output must be local')
 need(out.suffix.lower()=='.pptx','Output must be pptx')
 need(not out.is_relative_to(SKILL.resolve()),'Refusing output inside installed skill assets')
 data=src.read_bytes();digest=sha(data);need(digest==expected_sha256.upper(),'Input hash mismatch')
 docs,candidates,names=inventory(data);c=choose(candidates,slide_id,slide_number,shape_id,shape_tag)
 datasource=link_contract(c);chosen,reused=name_plan(c,names,digest,name)
 after,changes=rewrite(c,chosen,reused)
 tables=logical_tables(c)
 logical=[t for t in tables if str(t.get('chart_id'))==c['owner'].get('id') or (str(t.get('id'))==c['table'].get('id'))]
 need(len(logical)==1,'Logical datasource contract ambiguous')
 target={'slide_id':c['doc']['slide_id'],'slide_number':c['doc']['slide_number'],'model_type':c['owner'].tag,'logical_chart_id':c['owner'].get('id'),'logical_table_id':c['table'].get('id'),'active_document_part':c['doc']['part'],**c['frames'][0]}
 report={'schema_version':1,'status':'DRY_RUN_PASS','experimental_copy_only':True,'source':str(src),'source_sha256':digest,'output':str(out),'automation_name':chosen,'reused_existing_name':reused,'generation_scope':'entire presentation; unique names required','target':target,'datasource':{**datasource,'logical_model':logical[0]},'changes':changes,'only_name_fields_changed':True,'requires_json_update_and_native_verification':True}
 if execute:
  need(not out.exists(),'Output exists; refusing overwrite');need(out.parent.exists(),'Output parent does not exist')
  from chart_semantics import audit_scope
  audit=inspect_presentation(src,True);report['integrity_scope']=audit_scope(src,audit)
  with tempfile.TemporaryDirectory(prefix='tc_name_',dir=str(out.parent)) as td:
   temp=Path(td);changed_ole=c['doc']['ole']
   if not reused:
    carrier=temp/'carrier.bin';model_path=temp/'root.xml';carrier.write_bytes(changed_ole);model_path.write_bytes(after)
    proc=subprocess.run([powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(Path(__file__).with_name('replace_ole_stream.ps1')),'-StoragePath',str(carrier),'-StreamBytesPath',str(model_path)],capture_output=True,text=True,env=powershell_env())
    need(proc.returncode==0,'Structured-storage writer failed: '+proc.stderr[:2000]);changed_ole=carrier.read_bytes()
    ss=streams(changed_ole);need(set(ss)==set(c['doc']['streams']),'OLE stream inventory changed')
    need(all(ss[p]==v for p,v in c['doc']['streams'].items() if p!=('think-cellXML',)),'Other OLE stream changed')
    verify_diff(c,ss[('think-cellXML',)],chosen,reused)
   candidate=temp/'prepared.pptx'
   with zipfile.ZipFile(io.BytesIO(data)) as z,zipfile.ZipFile(candidate,'w') as dest:
    dest.comment=z.comment
    for entry in z.infolist():dest.writestr(copy.copy(entry),changed_ole if entry.filename==c['doc']['part'] else z.read(entry.filename))
   with zipfile.ZipFile(io.BytesIO(data)) as z,zipfile.ZipFile(candidate) as dz:
    need(z.namelist()==dz.namelist(),'ZIP inventory changed');need(dz.testzip() is None,'ZIP CRC failed')
    need(all(z.read(n)==dz.read(n) for n in z.namelist() if n!=c['doc']['part']),'Other ZIP entry changed')
   _,after_candidates,after_names=inventory(candidate.read_bytes());actual=choose(after_candidates,target['slide_id'],target['slide_number'],target['shape_id'],target['shape_tag'])
   need(actual['owner_name']==chosen and actual['table_name']==chosen,'Output name re-resolution failed')
   name_plan(actual,after_names,sha(candidate.read_bytes()),chosen);link_contract(actual)
   audit=inspect_presentation(candidate,True);audit_scope(candidate,audit)
   need(src.read_bytes()==data,'Source changed during preparation')
   with out.open('xb') as f:f.write(candidate.read_bytes())
  report.update(status='PREPARED_PENDING_NATIVE_VERIFICATION',output_sha256=sha(out.read_bytes()),all_integrity_checks=True,other_ole_streams_byte_identical=True,other_zip_entries_byte_identical=True,source_unchanged=True)
 return report

def main():
 p=argparse.ArgumentParser(description=__doc__)
 p.add_argument('--input',required=True,type=Path);p.add_argument('--expected-sha256',required=True);p.add_argument('--output',required=True,type=Path);p.add_argument('--manifest',required=True,type=Path)
 p.add_argument('--slide-id',type=int);p.add_argument('--slide-number',type=int);p.add_argument('--shape-id',type=int);p.add_argument('--shape-tag');p.add_argument('--name');p.add_argument('--execute',action='store_true')
 a=p.parse_args()
 try:
  need(a.manifest.resolve() not in {a.input.resolve(),a.output.resolve()},'Manifest overlaps presentation')
  need(not a.manifest.resolve().is_relative_to(SKILL.resolve()),'Refusing manifest inside installed skill assets')
  need(a.manifest.suffix.lower()=='.json','Manifest must be JSON')
  need(not a.manifest.exists(),'Manifest exists; use a new manifest path')
  report=prepare(a.input,a.expected_sha256,a.output,a.slide_id,a.slide_number,a.shape_id,a.shape_tag,a.name,a.execute)
  need(a.manifest.resolve() not in {a.input.resolve(),a.output.resolve()},'Manifest overlaps presentation')
  a.manifest.write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps({'status':report['status'],'automation_name':report['automation_name'],'manifest':str(a.manifest.resolve())}))
 except Exception as e:
  print(json.dumps({'status':'REJECTED','error':str(e)}),file=sys.stderr);sys.exit(1)
if __name__=='__main__':main()

