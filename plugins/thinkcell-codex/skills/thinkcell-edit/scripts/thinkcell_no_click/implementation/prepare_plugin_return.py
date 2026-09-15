"""Prepare plugin transport arguments for a verified, single-slide PPTX. Does not call Office.
Finalize against a fresh list_slides result after staging; separately check source changes.
"""
import argparse, base64, hashlib, io, json, uuid, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

def prepare(data, expected_sha256, slide_id, slide_index):
    digest = hashlib.sha256(data).hexdigest()
    if digest != expected_sha256.lower():
        raise ValueError('Candidate hash mismatch')
    if not slide_id or slide_index < 0:
        raise ValueError('Bound slide identity and current zero-based index required')
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        if len(z.namelist()) != len(set(z.namelist())) or z.testzip():
            raise ValueError('Invalid or duplicate ZIP entries')
        root = ET.fromstring(z.read('ppt/presentation.xml'))
        ns = {'p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
        if len(root.findall('p:sldIdLst/p:sldId', ns)) != 1:
            raise ValueError('Candidate must contain exactly one slide')
        if 'ppt/slides/slide1.xml' not in z.namelist():
            raise ValueError('Expected single-slide export layout')
    payload = base64.b64encode(data).decode('ascii')
    prefix = 'tc-return-' + uuid.uuid4().hex
    chunks = [payload[i:i+40000] for i in range(0,len(payload),40000)]
    staging = [{'slide_index':slide_index, 'summary':'Stage verified native slide data',
                'code':f'await writeFile({json.dumps(prefix+"-"+str(i)+".b64")}, {json.dumps(chunk)});'}
               for i,chunk in enumerate(chunks)]
    code = (f"let data=''; for(let i=0;i<{len(chunks)};i++) data+=await readFile({json.dumps(prefix)}+'-'+i+'.b64'); "
            f"if(data.length!=={len(payload)} || !data.startsWith('UEsDB')) throw new Error('Staged payload mismatch'); "
            "for(const name of Object.keys(zip.files)) zip.remove(name); "
            "await zip.loadAsync(data,{base64:true,checkCRC32:true}); "
            "if(!zip.file('ppt/slides/slide1.xml') || !zip.file('ppt/presentation.xml')) throw new Error('Missing native parts'); markDirty();")
    return {'tool_name':'edit_slide_ooxml', 'candidate_sha256':digest, 'bound_slide_id':slide_id,
            'requires_fresh_target_and_conflict_check':True, 'staging_args':staging,
            'replacement_template':{'summary':'Return verified native slide','code':code},
            'scope':'Transport only; not a conflict guard or native certification'}

def finalize(plan, listing):
    """Bind only the final write to its current position. Listing freshness is caller-owned."""
    slides = listing.get('slides')
    if not isinstance(slides, list):
        raise ValueError('Expected the fresh list_slides result object with slides')
    matches = [s for s in slides if s.get('id') == plan['bound_slide_id']]
    if len(matches) != 1:
        raise ValueError('Bound slide is missing or ambiguous; refresh/reconcile before writing')
    index = matches[0].get('slideIndex')
    if type(index) is not int or index < 0:
        raise ValueError('Invalid current slide index')
    if sum(s.get('slideIndex') == index for s in slides) != 1:
        raise ValueError('Ambiguous current slide index')
    args = dict(plan['replacement_template'])
    args['slide_index'] = index
    return {'tool_name':plan['tool_name'], 'bound_slide_id':plan['bound_slide_id'],
            'candidate_sha256':plan['candidate_sha256'], 'replacement_args':args,
            'source_content_recheck_required':True,
            'scope':'Execute promptly in the same verified session after source reconciliation; not atomic locking'}

if __name__ == '__main__':
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--candidate',type=Path);p.add_argument('--expected-sha256')
    p.add_argument('--slide-id');p.add_argument('--slide-index',type=int)
    p.add_argument('--plan',type=Path);p.add_argument('--slides-json',type=Path)
    p.add_argument('--output',type=Path,required=True);a=p.parse_args()
    if a.plan or a.slides_json:
        if not (a.plan and a.slides_json) or any(v is not None for v in [a.candidate,a.expected_sha256,a.slide_id,a.slide_index]):
            p.error('Finalize requires only --plan, --slides-json and --output')
        plan=finalize(json.loads(a.plan.read_text(encoding='utf-8-sig')),json.loads(a.slides_json.read_text(encoding='utf-8-sig')))
    else:
        if any(v is None for v in [a.candidate,a.expected_sha256,a.slide_id,a.slide_index]):
            p.error('Prepare requires --candidate, --expected-sha256, --slide-id and --slide-index')
        plan=prepare(a.candidate.read_bytes(),a.expected_sha256,a.slide_id,a.slide_index)
    with a.output.open('x',encoding='utf-8') as f:json.dump(plan,f,indent=2)
    print(json.dumps({'plan':str(a.output.resolve()),'stage':'finalized' if a.plan else 'prepared'}))
