"""Offline readback verifier for native Gantt candidates.

Run this after root's serialized native reopen/save verifier, against the saved
output package. It never opens PowerPoint.
"""
from __future__ import annotations
import argparse, json, posixpath, re, sys, zipfile
from pathlib import Path

BASE_BARS={'199':['2012-01-30T00:00:00','2012-03-31T00:00:00'],'200':['2012-04-02T00:00:00','2012-04-21T00:00:00'],'187':['2012-04-23T00:00:00','2012-05-19T00:00:00']}
BASE_MILESTONES={'175':'2012-04-23T00:00:00'}
def model(path):
 with zipfile.ZipFile(path) as z:
  raw=z.read('ppt/embeddings/oleObject13.bin'); m=re.search(rb'<root[^>]*>.*?</root>',raw,re.S)
  if not m: raise RuntimeError('Gantt model root missing')
  return m.group().decode('utf8')
def rec(t,typ,rid):
 q=re.search(r'<'+typ+r' id="'+rid+r'".*?</'+typ+'>',t,re.S)
 if not q: raise RuntimeError(f'{typ} {rid} missing')
 return q.group()
def dates(q): return re.findall(r'<m_datetime val="([^"]+)"',q)
def rect(q):
 m=re.search(r'<m_rectPPTShape left="(\d+)" top="(\d+)" right="(\d+)" bottom="(\d+)"/>',q)
 return list(map(int,m.groups())) if m else None
def slide_shape(path,name):
 with zipfile.ZipFile(path) as z:
  t=z.read('ppt/slides/slide1.xml').decode()
  for b in re.findall(r'<p:sp>.*?</p:sp>',t,re.S):
   n=re.search(r'<p:cNvPr[^>]* name="([^"]*)"',b)
   if n and n.group(1)==name:
    m=re.search(r'<a:off x="(\d+)" y="(\d+)"/><a:ext cx="(\d+)" cy="(\d+)"',b)
    rid=re.search(r'<p:tags r:id="([^"]+)"',b)
    tagval=None
    if rid:
     rels=z.read('ppt/slides/_rels/slide1.xml.rels').decode()
     rel=re.search(r'<Relationship[^>]* Id="'+re.escape(rid.group(1))+r'"[^>]* Target="([^"]+)"',rels)
     if rel:
      target=posixpath.normpath('ppt/slides/'+rel.group(1))
      try: tag=z.read(target).decode(); tv=re.search(r'name="THINKCELLSHAPEDONOTDELETE" val="([^"]+)"',tag); tagval=tv.group(1) if tv else None
      except KeyError: pass
    return {'transform':list(map(int,m.groups())) if m else None,'tag':tagval}
 raise RuntimeError(f'slide shape {name} missing')
def owner_binding(path):
 with zipfile.ZipFile(path) as z:
  rels=z.read('ppt/slides/_rels/slide1.xml.rels').decode()
  return bool(re.search(r'Type="[^"]*oleObject"[^>]*Target="\.\./embeddings/oleObject13\.bin"',rels))
def main():
 ap=argparse.ArgumentParser(); ap.add_argument('path'); ap.add_argument('--bar-dates',nargs=2,required=True); ap.add_argument('--line-bounds',nargs=2,type=int,required=True); ap.add_argument('--visible-x',nargs=2,type=int,required=True); ap.add_argument('--milestone-date'); ap.add_argument('--milestone-bounds',nargs=4,type=int); ap.add_argument('--visible-milestone-x',type=int); a=ap.parse_args()
 t=model(a.path); out={'path':str(Path(a.path).resolve()),'checks':{},'errors':[]}
 try:
  se=rec(t,'CGanttSE','8'); out['checks']['model_consistent']='m_bConsistent val="1"' in se
  out['checks']['slide_owner_ole_binding']=owner_binding(a.path)
  b=rec(t,'CGanttBar','214'); out['bar_214_dates']=dates(b); out['checks']['bar_214_dates']=out['bar_214_dates']==[x+'T00:00:00' for x in a.bar_dates]
  ln=rec(t,'CPPTAutoShapeLine','222'); out['line_222_bounds']=rect(ln); out['checks']['line_222_bounds']=out['line_222_bounds'][:1]+out['line_222_bounds'][2:3]==a.line_bounds if out['line_222_bounds'] else False
  sh=slide_shape(a.path,'Rectangle 191'); out['visible_bar_transform']=sh; out['checks']['visible_bar_bounds']=sh['transform'] and [sh['transform'][0],sh['transform'][0]+sh['transform'][2]]==a.visible_x
  gen=rec(t,'CPPTGenericLine','221'); line_name=re.search(r'<CPPTAutoShapeLine id="222".*?<m_bstrShapeName>([^<]+)',t,re.S).group(1); out['checks']['bar_model_graph']='<m_pptautoshpline idref="222"' in gen and sh['tag']==line_name; out['bar_214_shape_tag']=line_name; out['visible_bar_tag']=sh['tag']
  for rid,ds in BASE_BARS.items(): out['checks']['bar_'+rid+'_unchanged']=dates(rec(t,'CGanttBar',rid))==ds
  for rid,ds in BASE_MILESTONES.items(): out['checks']['milestone_'+rid+'_unchanged']=dates(rec(t,'CGanttMilestone',rid))==[ds]
  if a.milestone_date:
   mm=rec(t,'CGanttMilestone','174'); out['milestone_174_date']=dates(mm); out['checks']['milestone_174_date']=out['milestone_174_date']==[a.milestone_date+'T00:00:00']; out['milestone_174_marker_bounds']=rect(mm); out['checks']['milestone_174_marker_bounds']=out['milestone_174_marker_bounds']==a.milestone_bounds
   sh=slide_shape(a.path,'Isosceles Triangle 207'); out['visible_milestone_transform']=sh; out['checks']['visible_milestone_center']=sh['transform'] and sh['transform'][0]+sh['transform'][2]//2==a.visible_milestone_x
   marker_name=re.search(r'<m_pptautoshpmarker>.*?<m_bstrShapeName>([^<]+)',mm,re.S).group(1); out['checks']['milestone_model_graph']=sh['tag']==marker_name; out['milestone_174_shape_tag']=marker_name; out['visible_milestone_tag']=sh['tag']
  else: out['checks']['milestone_174_untouched']=dates(rec(t,'CGanttMilestone','174'))==['2012-01-30T00:00:00']
 except Exception as e: out['errors'].append(str(e))
 out['pass']=not out['errors'] and all(out['checks'].values()); print(json.dumps(out,indent=2)); return 0 if out['pass'] else 2
if __name__=='__main__': sys.exit(main())
