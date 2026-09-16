"""Prepare a semantic native Gantt milestone edit offline.

The simple donor has real CGanttMilestone records and tagged visible marker
shapes.  This adapter updates the typed date, native marker rectangle, and the
matching visible shape as one package edit.  PowerPoint reopen/readback is
owned by the root runner.
"""
from __future__ import annotations
import argparse, bisect, hashlib, json, posixpath, re, subprocess, sys, tempfile, zipfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path

REPLACE_STREAM = Path(__file__).with_name("replace_ole_stream.ps1")
_SCRIPTS_DIR = Path(__file__).resolve().parents[1]
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))
from runtime import powershell_env

@dataclass
class Milestone:
    id: str
    when: str
    style: str
    shape_name: str
    marker_rect: tuple[int, int, int, int]

class Package:
    def __init__(self, path: Path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as z:
            self.entries = z.infolist()
            self.blobs = {e.filename: z.read(e.filename) for e in self.entries}
        found = []
        for name, raw in self.blobs.items():
            if name.startswith('ppt/embeddings/') and name.endswith('.bin'):
                m = re.search(rb'<root[^>]*>.*?</root>', raw, re.S)
                if m and b'<CGanttSE' in m.group(): found.append((name, m.group().decode('utf8')))
        if len(found) != 1: raise ValueError(f'expected one Gantt owner, found {len(found)}')
        self.ole_part, self.model = found[0]
        self.milestones = self._discover_milestones()
        self.axis_dates, self.axis_bounds = self._discover_axis()

    def _discover_milestones(self):
        out = []
        for m in re.finditer(r'<CGanttMilestone id="([^"]+)".*?</CGanttMilestone>', self.model, re.S):
            q = m.group()
            dt = re.search(r'<m_datetime val="([^"]+)"', q)
            style = re.search(r'<m_emarkerstyle val="([^"]+)"', q)
            name = re.search(r'<m_pptautoshpmarker>.*?<m_bstrShapeName>([^<]+)', q, re.S)
            rect = re.search(r'<m_rectPPTShape left="(\d+)" top="(\d+)" right="(\d+)" bottom="(\d+)"', q)
            if dt and style and name and rect:
                out.append(Milestone(m.group(1), dt.group(1), style.group(1), name.group(1), tuple(map(int, rect.groups()))))
        if not out: raise ValueError('no typed Gantt milestones found')
        return out

    def _discover_axis(self):
        sm = re.search(r'<CGanttScaleWeeks id="[^"]+".*?</CGanttScaleWeeks>', self.model, re.S)
        if not sm: raise ValueError('weekly native scale missing')
        vals = []
        for rid in re.findall(r'<elem idref="([^"]+)"', sm.group()):
            q = re.search(r'<CPPTGanttScaleBox id="'+re.escape(rid)+r'".*?</CPPTGanttScaleBox>', self.model, re.S)
            if not q: raise ValueError(f'scale box {rid} missing')
            rr = re.search(r'<m_rectPPTShape left="(\d+)".*?right="(\d+)"', q.group(), re.S)
            vr = re.search(r'<m_varsrc idref="([^"]+)"', q.group())
            src = re.search(r'<CVariableSource id="'+re.escape(vr.group(1) if vr else '')+r'".*?</CVariableSource>', self.model, re.S) if vr else None
            dd = re.search(r'<m_datetime val="(\d{4}-\d{2}-\d{2})T', src.group()) if src else None
            if not rr or not dd: raise ValueError('scale box lacks native date identity')
            vals.append((date.fromisoformat(dd.group(1)), int(rr.group(1)), int(rr.group(2))))
        vals.sort()
        ds = [x[0] for x in vals]
        if len(ds) < 2 or len(set(ds)) != len(ds) or any((b-a).days != 7 for a,b in zip(ds, ds[1:])):
            raise ValueError('weekly scale dates are not unique seven-day intervals')
        if any(x[2] != vals[i+1][1] for i,x in enumerate(vals[:-1])): raise ValueError('weekly scale bounds are not contiguous')
        return ds, [vals[0][1]] + [x[2] for x in vals]

    def x_for(self, iso):
        d = date.fromisoformat(iso)
        if d < self.axis_dates[0] or d > self.axis_dates[-1]: raise ValueError(f'date {iso} outside native scale')
        return self.axis_bounds[max(0, min(bisect.bisect_right(self.axis_dates, d)-1, len(self.axis_bounds)-1))]

    def resolve(self, *, milestone_id=None, date_value=None, style=None):
        c = self.milestones
        if milestone_id is not None: c = [x for x in c if x.id == str(milestone_id)]
        if date_value is not None: c = [x for x in c if x.when.startswith(date_value)]
        if style is not None: c = [x for x in c if x.style == style]
        if len(c) != 1: raise ValueError(f'semantic milestone selector matched {len(c)} records')
        return c[0]

    def visible(self, shape_name):
        slide = self.blobs['ppt/slides/slide1.xml'].decode('utf8')
        rels = self.blobs['ppt/slides/_rels/slide1.xml.rels'].decode('utf8')
        for block in re.findall(r'<p:sp>.*?</p:sp>', slide, re.S):
            tag = re.search(r'<p:tags r:id="([^"]+)"', block)
            if not tag: continue
            rel = re.search(r'<Relationship[^>]* Id="'+re.escape(tag.group(1))+r'"[^>]* Target="([^"]+)"', rels)
            if not rel: continue
            part = posixpath.normpath('ppt/slides/'+rel.group(1))
            tv = re.search(r'name="THINKCELLSHAPEDONOTDELETE" val="([^"]+)"', self.blobs.get(part,b'').decode('utf8','ignore'))
            if tv and tv.group(1) == shape_name:
                tr = re.search(r'<a:off x="(\d+)" y="(\d+)"/><a:ext cx="(\d+)" cy="(\d+)"', block)
                if not tr: raise ValueError('visible marker transform missing')
                return block, tuple(map(int, tr.groups()))
        raise ValueError(f'visible tagged marker not found: {shape_name}')

    def edit(self, output, *, selector, new_date):
        old = self.resolve(**selector)
        new_center = self.x_for(new_date)
        nq = re.search(r'<CGanttMilestone id="'+re.escape(old.id)+r'".*?</CGanttMilestone>', self.model, re.S)
        block = nq.group(); old_dt = old.when
        block = block.replace(old_dt, new_date+'T00:00:00', 1)
        mr = re.search(r'<m_rectPPTShape left="(\d+)" top="(\d+)" right="(\d+)" bottom="(\d+)"', block)
        l, t, r, b = map(int, mr.groups()); half = (r-l)//2
        block = block.replace(mr.group(), f'<m_rectPPTShape left="{new_center-half}" top="{t}" right="{new_center+half}" bottom="{b}"/>', 1)
        model = self.model[:nq.start()] + block + self.model[nq.end():]
        shape, tr = self.visible(old.shape_name); old_left, y, cx, cy = tr
        old_center = old_left + cx//2
        old_model_center = (old.marker_rect[0] + old.marker_rect[2])//2
        # The donor's native model-to-slide scale is 1587.5 slide units/model unit,
        # established by the bound taskbar and retained for marker placement.
        slide_x = round(old_center + (new_center - old_model_center) * 1587.5)
        new_left = slide_x - cx//2
        tm = re.search(r'<a:off x="(\d+)" y="(\d+)"/><a:ext cx="(\d+)" cy="(\d+)"', shape)
        shape2 = shape[:tm.start()] + f'<a:off x="{new_left}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"' + shape[tm.end():]
        self.blobs['ppt/slides/slide1.xml'] = self.blobs['ppt/slides/slide1.xml'].decode('utf8').replace(shape, shape2, 1).encode()
        with tempfile.TemporaryDirectory(dir=self.path.parent) as td:
            carrier = Path(td)/'carrier.bin'; xml = Path(td)/'model.xml'; carrier.write_bytes(self.blobs[self.ole_part]); xml.write_text(model, encoding='utf8')
            p = subprocess.run(['powershell','-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(REPLACE_STREAM),'-StoragePath',str(carrier),'-StreamBytesPath',str(xml)], capture_output=True, text=True, env=powershell_env())
            if p.returncode: raise RuntimeError(p.stderr or p.stdout)
            self.blobs[self.ole_part] = carrier.read_bytes()
        with zipfile.ZipFile(output,'w') as z:
            for e in self.entries: z.writestr(e, self.blobs[e.filename])
        return {'milestone_id':old.id,'old_date':old_dt,'new_date':new_date+'T00:00:00','old_model_bounds':list(old.marker_rect),'new_model_bounds':[new_center-half, old.marker_rect[1], new_center+half, old.marker_rect[3]],'shape_name':old.shape_name,'old_visible_transform':list(tr),'new_visible_center':slide_x}

def sha(path):
    h=hashlib.sha256()
    with open(path,'rb') as f:
        for chunk in iter(lambda:f.read(1<<20), b''): h.update(chunk)
    return h.hexdigest()

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--source',required=True); ap.add_argument('--expected-sha256',required=True); ap.add_argument('--output',required=True); ap.add_argument('--report',required=True); ap.add_argument('--new-date',required=True); ap.add_argument('--milestone-date'); ap.add_argument('--milestone-id'); ap.add_argument('--style'); a=ap.parse_args()
    src,out,report=map(Path,(a.source,a.output,a.report)); before=sha(src)
    if before.lower()!=a.expected_sha256.lower(): raise SystemExit('source SHA256 mismatch')
    if out.exists() or report.exists(): raise SystemExit('output/report already exists')
    edit=Package(src).edit(out,selector={'milestone_id':a.milestone_id,'date_value':a.milestone_date,'style':a.style},new_date=a.new_date)
    result={'status':'PAIRED_GANTT_MILESTONE_EDIT_PREPARED','source':str(src.resolve()),'output':str(out.resolve()),'source_sha256_before':before,'source_sha256_after':sha(src),'source_unchanged':before==sha(src),'edit':edit}
    report.write_text(json.dumps(result,indent=2)+'\n',encoding='utf8'); print(json.dumps(result,indent=2))
if __name__=='__main__': main()
