"""Portable paired native Gantt date + geometry adapter.

This lane-local reference implementation discovers owner/table/bar/line/shape IDs
from the package and uses only the selected record after semantic filtering.
It reads each weekly scale box's native date variable and rejects a scale with
missing, irregular, or non-contiguous calendar identity. It does no Office
automation; production callers must native reopen/read back.
"""
from __future__ import annotations
from dataclasses import dataclass
from datetime import date
from pathlib import Path
import bisect, json, posixpath, re, subprocess, tempfile, zipfile

# Bundle `replace_ole_stream.ps1` next to this module. This keeps the adapter
# relocatable when a skill is installed outside the research workspace.
REPLACE_STREAM=Path(__file__).with_name('replace_ole_stream.ps1')
@dataclass
class Bar:
    id:str; row:int; start:str; end:str; line_id:str; shape_name:str; line_rect:tuple
@dataclass
class Axis:
    boundaries:list[int]; dates:list[date]

class GanttPackage:
  def __init__(self,path):
    self.path=Path(path); self.z=zipfile.ZipFile(self.path); self.blobs={n:self.z.read(n) for n in self.z.namelist()}; self.entries=self.z.infolist(); self.ole_part,self.model=self._load_model(); self._bars=self._discover_bars(); self.axis=self._discover_axis()
  def _load_model(self):
    found=[]
    for n,b in self.blobs.items():
      if n.startswith('ppt/embeddings/') and n.endswith('.bin'):
       m=re.search(rb'<root[^>]*>.*?</root>',b,re.S)
       if m and b'<CGanttSE' in m.group(): found.append((n,m.group().decode('utf8')))
    if len(found)!=1: raise ValueError(f'expected one native Gantt owner, found {len(found)}')
    return found[0]
  def _discover_bars(self):
    bars=[]
    vectors=list(re.finditer(r'<CGanttVector id="([^"]+)".*?</CGanttVector>',self.model,re.S))
    for row,vm in enumerate(reversed(vectors)):
      v=vm.group()
      ids=re.findall(r'<m_cganttbar[^>]*>.*?</m_cganttbar>',v,re.S)
      refs=[]
      for x in ids: refs += re.findall(r'<elem idref="([^"]+)"',x)
      for bid in refs:
       q=re.search(r'<CGanttBar id="'+re.escape(bid)+r'".*?</CGanttBar>',self.model,re.S)
       if not q: continue
       b=q.group(); ds=re.findall(r'<m_datetime val="([^"]+)"',b)
       if len(ds)!=2: continue
       lr=re.search(r'<m_pptgenline idref="([^"]+)"',b)
       if not lr: continue
       gen=re.search(r'<CPPTGenericLine id="'+lr.group(1)+r'".*?</CPPTGenericLine>',self.model,re.S)
       shp=re.search(r'<m_pptautoshpline idref="([^"]+)"',gen.group()) if gen else None
       line_id=shp.group(1) if shp else None
       line=re.search(r'<CPPTAutoShapeLine id="'+re.escape(line_id or '')+r'".*?</CPPTAutoShapeLine>',self.model,re.S)
       name=re.search(r'<m_bstrShapeName>([^<]+)',line.group()) if line else None
       rect=re.search(r'<m_rectPPTShape left="(\d+)" top="(\d+)" right="(\d+)" bottom="(\d+)"',line.group()) if line else None
       if line_id and name and rect: bars.append(Bar(bid,row,ds[0],ds[1],line_id,name.group(1),tuple(map(int,rect.groups()))))
    return bars
  def _discover_axis(self):
    sm=re.search(r'<CGanttScaleWeeks id="([^"]+)".*?</CGanttScaleWeeks>',self.model,re.S)
    if not sm: raise ValueError('weekly native scale missing')
    refs=re.findall(r'<elem idref="([^"]+)"',sm.group())
    vals=[]
    for rid in refs:
      q=re.search(r'<CPPTGanttScaleBox id="'+re.escape(rid)+r'".*?</CPPTGanttScaleBox>',self.model,re.S)
      if not q: raise ValueError(f'weekly scale box {rid} missing')
      r=re.search(r'<m_rectPPTShape left="(\d+)".*?right="(\d+)"',q.group(),re.S)
      vr=re.search(r'<m_varsrc idref="([^"]+)"',q.group())
      src=re.search(r'<CVariableSource id="'+re.escape(vr.group(1) if vr else '')+r'".*?</CVariableSource>',self.model,re.S) if vr else None
      dt=re.search(r'<m_datetime val="(\d{4}-\d{2}-\d{2})T',src.group()) if src else None
      if not r or not dt: raise ValueError('weekly scale lacks native calendar date identity')
      vals.append((date.fromisoformat(dt.group(1)),int(r.group(1)),int(r.group(2))))
    if len(vals)<2: raise ValueError('weekly boundaries missing')
    vals.sort(key=lambda x:x[0])
    dates=[v[0] for v in vals]
    if len(set(dates))!=len(dates): raise ValueError('weekly scale dates are not unique')
    if any((b-a).days!=7 for a,b in zip(dates,dates[1:])): raise ValueError('weekly scale dates are not seven-day intervals')
    if any(left>=right for _,left,right in vals): raise ValueError('weekly scale bounds are invalid')
    if any(vals[i][2]!=vals[i+1][1] for i in range(len(vals)-1)): raise ValueError('weekly scale bounds are not contiguous')
    return Axis([vals[0][1]]+[v[2] for v in vals],dates)
  def list_bars(self): return [b.__dict__ for b in self._bars]
  def resolve_bar(self, *, row=None, start=None, end=None, bar_id=None):
    c=self._bars
    if row is not None: c=[b for b in c if b.row==row]
    if start is not None: c=[b for b in c if b.start.startswith(start)]
    if end is not None: c=[b for b in c if b.end.startswith(end)]
    if bar_id is not None: c=[b for b in c if b.id==str(bar_id)]
    if len(c)!=1: raise ValueError(f'semantic bar selector matched {len(c)} records')
    return c[0]
  def x_for(self,iso,edge='start'):
    d=date.fromisoformat(iso)
    if edge not in ('start','end'): raise ValueError(f'unknown axis edge: {edge}')
    if d<self.axis.dates[0] or d>self.axis.dates[-1]: raise ValueError(f'date {iso} is outside native weekly scale')
    idx=bisect.bisect_right(self.axis.dates,d)-(1 if edge=='start' else 0)
    return self.axis.boundaries[max(0,min(idx,len(self.axis.boundaries)-1))]
  def _shape(self,name):
    t=self.blobs['ppt/slides/slide1.xml'].decode('utf8')
    rels=self.blobs['ppt/slides/_rels/slide1.xml.rels'].decode('utf8')
    for b in re.findall(r'<p:sp>.*?</p:sp>',t,re.S):
      rid=re.search(r'<p:tags r:id="([^"]+)"',b); nm=re.search(r'<p:cNvPr[^>]* name="([^"]+)"',b)
      if not rid: continue
      rel=re.search(r'<Relationship[^>]* Id="'+re.escape(rid.group(1))+r'"[^>]* Target="([^"]+)"',rels)
      if not rel: continue
      target=posixpath.normpath('ppt/slides/'+rel.group(1)); tag=self.blobs.get(target,b'').decode('utf8','ignore')
      tv=re.search(r'name="THINKCELLSHAPEDONOTDELETE" val="([^"]+)"',tag)
      if tv and tv.group(1)==name:
       tr=re.search(r'<a:off x="(\d+)" y="(\d+)"/><a:ext cx="(\d+)" cy="(\d+)"',b)
       return b, list(map(int,tr.groups())) if tr else None
    raise ValueError(f'visible tagged shape not found for {name}')
  def edit_bar(self,out,*,selector,new_start,new_end):
    b=self.resolve_bar(**selector); old_line=b.line_rect; ml,mr=self.x_for(new_start,'start'),self.x_for(new_end,'end')
    q=re.search(r'<CGanttBar id="'+re.escape(b.id)+r'".*?</CGanttBar>',self.model,re.S); block=q.group(); old=re.findall(r'<m_datetime val="([^"]+)"',block)
    block2=block.replace(old[0],new_start+'T00:00:00',1).replace(old[1],new_end+'T00:00:00',1); model=self.model[:q.start()]+block2+self.model[q.end():]
    lq=re.search(r'<CPPTAutoShapeLine id="'+re.escape(b.line_id)+r'".*?</CPPTAutoShapeLine>',model,re.S); rect=re.search(r'<m_rectPPTShape left="(\d+)" top="(\d+)" right="(\d+)" bottom="(\d+)"',lq.group()); vals=list(map(int,rect.groups())); vals[0],vals[2]=ml,mr; lq2=lq.group().replace(rect.group(),f'<m_rectPPTShape left="{vals[0]}" top="{vals[1]}" right="{vals[2]}" bottom="{vals[3]}"',1); model=model[:lq.start()]+lq2+model[lq.end():]
    _,tr=self._shape(b.shape_name); sl,sy,scy,sbot=tr
    # shape transform is [left,top,cx,cy]; derive slide x from old model bounds.
    old_sl=tr[0]; old_sr=tr[0]+tr[2]; new_sl=round(old_sl+(ml-old_line[0])*(old_sr-old_sl)/(old_line[2]-old_line[0])); new_sr=round(old_sl+(mr-old_line[0])*(old_sr-old_sl)/(old_line[2]-old_line[0]))
    slide=self.blobs['ppt/slides/slide1.xml'].decode('utf8'); shape,_=self._shape(b.shape_name); tm=re.search(r'<a:off x="(\d+)" y="(\d+)"/><a:ext cx="(\d+)" cy="(\d+)"',shape); shape2=shape[:tm.start()]+f'<a:off x="{new_sl}" y="{tm.group(2)}"/><a:ext cx="{new_sr-new_sl}" cy="{tm.group(4)}"'+shape[tm.end():]; self.blobs['ppt/slides/slide1.xml']=slide.replace(shape,shape2,1).encode()
    with tempfile.TemporaryDirectory(dir=self.path.parent) as td:
      td=Path(td); carrier=td/'carrier.bin'; xml=td/'model.xml'; carrier.write_bytes(self.blobs[self.ole_part]); xml.write_text(model,encoding='utf8'); p=subprocess.run(['powershell','-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(REPLACE_STREAM),'-StoragePath',str(carrier),'-StreamBytesPath',str(xml)],capture_output=True,text=True)
      if p.returncode: raise RuntimeError(p.stderr or p.stdout)
      self.blobs[self.ole_part]=carrier.read_bytes()
    out=Path(out)
    with zipfile.ZipFile(out,'w') as zout:
      for e in self.entries:zout.writestr(e,self.blobs[e.filename])
    return {'bar_id':b.id,'row':b.row,'old_dates':[b.start,b.end],'new_dates':[new_start+'T00:00:00',new_end+'T00:00:00'],'old_model_line':list(old_line),'new_model_line':[ml,old_line[1],mr,old_line[3]],'shape_name':b.shape_name,'new_visible_bounds':[new_sl,new_sr]}

def paired_date_geometry_edit(source,out,selector,start,end): return GanttPackage(source).edit_bar(out,selector=selector,new_start=start,new_end=end)
