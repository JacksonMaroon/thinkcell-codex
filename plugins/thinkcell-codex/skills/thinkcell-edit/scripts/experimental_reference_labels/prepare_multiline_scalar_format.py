"""Prepare a hash-bound multiline format for one existing native scalar label.

Experimental and profile-specific: this changes only the existing label's
native prefix/suffix formatting. It never inserts labels, annotations, shapes,
or data. Official regeneration and native visual proof remain required.
"""
from __future__ import annotations
import argparse, copy, io, json, subprocess, sys, tempfile, zipfile
from pathlib import Path
from lxml import etree as E

HERE = Path(__file__).resolve().parent; SCRIPTS = HERE.parent
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / 'thinkcell_no_click' / 'implementation')]
from chart_geometry import inventory, need, sha, streams, xml
from runtime import powershell, powershell_env

PROFILE = {'model_version':'38764','label_id':'458','scalar_id':'431','source_id':'454',
           'source_value':'2.96999999999999992895E+01','font_size':'11','prefix':'$', 'suffix':'B'}
NEW_SUFFIX='B\nExpansion\nopportunity'

def target_document(raw):
    _, charts, _ = inventory(raw)
    docs = {}
    for chart in charts:
        doc = chart['doc']; root = doc['root']; label = root.find("./CSequenceChartDataScalarLabel[@id='458']")
        if label is not None: docs[doc['part']] = doc
    need(len(docs)==1, 'label profile must occur in exactly one carrier')
    return next(iter(docs.values()))

def profile(root):
    version=root.find('version'); need(version is not None and version.get('val') == PROFILE['model_version'], 'model version differs')
    ids = {n.get('id'):n for n in root if n.get('id')}
    label = ids.get(PROFILE['label_id']); scalar = ids.get(PROFILE['scalar_id']); source = ids.get(PROFILE['source_id'])
    need(label is not None and label.tag=='CSequenceChartDataScalarLabel', 'label profile differs')
    need(scalar is not None and scalar.tag=='CSequenceChartDataScalar', 'scalar profile differs')
    need(source is not None and source.tag=='CVariableSource', 'source profile differs')
    ref = scalar.find('m_scdlabel'); src = scalar.find('m_varsrcAbsolute')
    need(ref is not None and ref.get('idref')==PROFILE['label_id'], 'scalar-label binding differs')
    need(src is not None and src.get('idref')==PROFILE['source_id'], 'scalar-source binding differs')
    value = source.find('m_varval'); need(value is not None and value.get('val')==PROFILE['source_value'], 'source value differs')
    fonts=[label.find('m_ppttb/m_font/m_nSize'),label.find('m_font/m_nSize')]
    prefix=label.find('m_prec/m_strPrefix'); suffix=label.find('m_prec/m_strSuffix17909')
    need(all(x is not None and float(x.get('val'))==float(PROFILE['font_size']) for x in fonts),'font profile differs')
    need(prefix is not None and prefix.text==PROFILE['prefix'] and suffix is not None and suffix.text==PROFILE['suffix'],'format profile differs')
    return label, prefix, suffix

def prepare(source_path, expected_source_sha256, output_path, report_path):
    source, output, report_file = map(Path, (source_path, output_path, report_path))
    need(len({source.resolve(), output.resolve(), report_file.resolve()})==3, 'source, output, and report paths must differ')
    need(not output.exists() and not report_file.exists(), 'output and report must be new')
    raw=source.read_bytes(); digest=sha(raw); need(digest==expected_source_sha256.upper(), 'source SHA-256 mismatch before mutation')
    doc=target_document(raw); before=doc['streams'][('think-cellXML',)]; root=xml(before); label,prefix,suffix=profile(root)
    original=copy.deepcopy(label); suffix.text=NEW_SUFFIX
    after=E.tostring(root,encoding='utf-8')
    reversed_root=xml(after); changed=reversed_root.find("./CSequenceChartDataScalarLabel[@id='458']"); changed.getparent().replace(changed,original)
    need(E.tostring(reversed_root,method='c14n')==E.tostring(xml(before),method='c14n'),'whole-model reversal failed')
    with tempfile.TemporaryDirectory(dir=output.parent,prefix='multiline_scalar_') as td:
        td=Path(td); carrier=td/'carrier.bin'; model=td/'model.xml'; carrier.write_bytes(doc['ole']); model.write_bytes(after)
        q=subprocess.run([powershell(),'-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(SCRIPTS/'thinkcell_no_click'/'implementation'/'replace_ole_stream.ps1'),'-StoragePath',str(carrier),'-StreamBytesPath',str(model)],capture_output=True,text=True,env=powershell_env(),timeout=60)
        need(q.returncode==0,'OLE patch failed: '+q.stderr[-400:]); changed_carrier=carrier.read_bytes(); updated=streams(changed_carrier)
        need(set(updated)==set(doc['streams']) and all(updated[k]==v for k,v in doc['streams'].items() if k!=('think-cellXML',)),'non-model CFB stream drift')
        with zipfile.ZipFile(io.BytesIO(raw)) as zin, zipfile.ZipFile(output,'x') as zout:
            zout.comment=zin.comment
            for item in zin.infolist(): zout.writestr(copy.copy(item),changed_carrier if item.filename==doc['part'] else zin.read(item.filename))
    with zipfile.ZipFile(io.BytesIO(raw)) as zin, zipfile.ZipFile(output) as zout:
        changed=[n for n in zout.namelist() if zin.read(n)!=zout.read(n)]
        need(set(changed)=={doc['part']} and set(zin.namelist())==set(zout.namelist()),'package scope drift')
    need(sha(source.read_bytes())==digest,'source changed')
    report={'status':'MULTILINE_SCALAR_FORMAT_PREPARED_REGEN_REQUIRED','source_path':str(source.resolve()),'output_path':str(output.resolve()),'report_path':str(report_file.resolve()),'expected_source_sha256':expected_source_sha256.upper(),'source_sha256':digest,'output_sha256':sha(output.read_bytes()),'carrier':doc['part'],'profile':PROFILE,'preserved_prefix':PROFILE['prefix'],'new_suffix':NEW_SUFFIX,'whole_model_reversal_pass':True,'non_model_cfb_streams_unchanged':True,'only_selected_carrier_changed':True,'source_unchanged':True,'native_gates_remaining':['official regeneration','native reopen/render','changed-data verification']}
    report_file.write_text(json.dumps(report,indent=2),encoding='utf-8'); return report

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__); p.add_argument('--input',required=True); p.add_argument('--expected-source-sha256',required=True); p.add_argument('--output',required=True); p.add_argument('--report',required=True)
    a=p.parse_args(); print(json.dumps(prepare(a.input,a.expected_source_sha256,a.output,a.report),indent=2))
