"""Read-only saved-snapshot guard. Does not control Office or alter presentations.

record snapshot.pptx manifest.json --slide-id 549
compare fresh-snapshot.pptx manifest.json

The caller must create snapshots from the EXACT held presentation object with
SaveCopyAs, and recheck its identity immediately before insertion. This guard
cannot observe coauthoring state that has not arrived in the local Office model.
"""
import argparse
import hashlib
import json
import posixpath
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def digest(data):
    return hashlib.sha256(data).hexdigest()

def canonical_part(name, data):
    """Normalize observed native serialization only; retain actual user content."""
    if not (name.endswith('.xml') or name.endswith('.rels')):
        return data
    root = ET.fromstring(data)
    # Installed think-cell rebrands this OLE type display label on native save.
    # This is not a cNvPr shape name or an automation name.
    for node in root.iter('{'+P+'}oleObj'):
        if node.get('progId') == 'TCLayout.ActiveDocument.1' and node.get('name') == 'think-cell Slide':
            node.set('name', 'thinkcell Slide')
    if name == 'docProps/app.xml':
        e = '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}'
        vt = '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}'
        timer = root.find(e+'TotalTime')
        if timer is not None:
            timer.text = '0'
        headings = root.find(e+'HeadingPairs/'+vt+'vector')
        titles = root.find(e+'TitlesOfParts/'+vt+'vector')
        if headings is not None and titles is not None:
            groups = list(headings); values = list(titles); offset = 0; normalized = []
            for i in range(0, len(groups), 2):
                label = groups[i].find(vt+'lpstr')
                count = groups[i+1].find(vt+'i4')
                if label is None or count is None:
                    raise ValueError('Unsupported extended-properties grouping')
                n = int(count.text); segment = values[offset:offset+n]; offset += n
                if label.text == 'Embedded OLE Servers':
                    seen_alias = False; kept = []
                    for item in segment:
                        if item.text in ('think-cell Slide', 'thinkcell Slide'):
                            if seen_alias:
                                continue
                            item.text = 'thinkcell Slide'; seen_alias = True
                        kept.append(item)
                    segment = kept; count.text = str(len(segment))
                normalized.extend(segment)
            if offset != len(values):
                raise ValueError('Extended-properties grouping does not cover titles')
            titles[:] = normalized; titles.set('size', str(len(normalized)))
    return ET.canonicalize(ET.tostring(root, encoding='unicode')).encode('utf-8')

def inspect(path, slide_id):
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist() if not n.endswith('/')}
    # Conservative: compare all Office content, including comments and links.
    # Ignore core metadata timestamps only, not custom properties or content.
    hashes = {n: digest(canonical_part(n,b)) for n, b in sorted(parts.items()) if n != 'docProps/core.xml'}
    rels = {}
    for r in ET.fromstring(parts['ppt/_rels/presentation.xml.rels']):
        if r.get('TargetMode') != 'External':
            rels[r.get('Id')] = posixpath.normpath(posixpath.join('ppt', r.get('Target'))).lstrip('/')
    presentation = ET.fromstring(parts['ppt/presentation.xml'])
    slides = [{'id': int(s.get('id')), 'part': rels[s.get('{'+R+'}id')]} for s in presentation.findall('./{'+P+'}sldIdLst/{'+P+'}sldId')]
    matches = [i for i, s in enumerate(slides) if s['id'] == slide_id]
    if len(matches) != 1:
        raise ValueError('Expected exactly one stable slide ID')
    i = matches[0]
    comments = [n for n in parts if 'comment' in n.lower()]
    customshows = presentation.find('{'+P+'}custShowLst') is not None
    internal_links = []
    external = []
    for n, b in parts.items():
        if n.endswith('.rels'):
            for r in ET.fromstring(b):
                if r.get('TargetMode') == 'External':
                    external.append({'part': n, 'type': r.get('Type'), 'target': r.get('Target')})
                elif r.get('Type', '').endswith('/slide') and '/slides/' in n:
                    internal_links.append({'part': n, 'target': r.get('Target')})
    return {'schema': 2, 'canonicalization': 'tc-ole-display-label-and-derived-ole-list-total-time-v1', 'source': str(Path(path).resolve()), 'file_sha256': digest(Path(path).read_bytes()),
            'target': slides[i], 'index': i + 1, 'slide_count': len(slides),
            'prior': slides[i-1] if i else None, 'following': slides[i+1] if i+1 < len(slides) else None,
            'slide_ids': [s['id'] for s in slides], 'part_hashes': hashes,
            'cautions': {'comment_parts': comments, 'custom_shows': customshows,
                         'internal_slide_relationships': internal_links, 'external_relationships': external},
            'limitations': ['No COM/live identity check', 'No server coauthoring lock',
                           'External relationship list does not identify every think-cell link stored inside OLE',
                           'Conservative byte comparison may reject harmless native serialization changes']}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('mode', choices=['record', 'compare'])
    ap.add_argument('snapshot', type=Path)
    ap.add_argument('manifest', type=Path)
    ap.add_argument('--slide-id', type=int)
    a = ap.parse_args()
    if a.mode == 'record':
        if not a.slide_id or a.manifest.exists():
            raise ValueError('Record requires slide ID and a new manifest path')
        result = inspect(a.snapshot, a.slide_id)
        a.manifest.parent.mkdir(parents=True, exist_ok=True)
        a.manifest.write_text(json.dumps(result, indent=2), encoding='utf-8')
        print(json.dumps({'status': 'SNAPSHOT_RECORDED_NO_MUTATION', 'target': result['target'], 'cautions': result['cautions']}))
    else:
        before = json.loads(a.manifest.read_text(encoding='utf-8'))
        if before.get('schema') != 2:
            raise ValueError('Manifest predates canonicalization; re-record it from the exact original sealed snapshot, never from a changed live source')
        after = inspect(a.snapshot, before['target']['id'])
        keys = set(before['part_hashes']) | set(after['part_hashes'])
        changes = [k for k in sorted(keys) if before['part_hashes'].get(k) != after['part_hashes'].get(k)]
        passed = not changes and before['slide_ids'] == after['slide_ids']
        print(json.dumps({'status': 'UNCHANGED_SAVED_SNAPSHOT' if passed else 'STOP_CHANGED_SNAPSHOT', 'changed_parts': changes, 'pass': passed}))
        return 0 if passed else 2
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
