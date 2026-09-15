"""Portable regression tests. No Office, proprietary decks, or network required."""
from pathlib import Path
import argparse
import copy
import hashlib
import io
import json
import os
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile

SCRIPTS = Path(__file__).resolve().parents[1]/'skills/thinkcell-edit/scripts'
sys.path[:0] = [str(SCRIPTS),str(SCRIPTS/'thinkcell_no_click/implementation')]
from lxml import etree as E
import prepare_thinkcell_name as namer
import prepare_plugin_return as handoff
import update_thinkcell_json as updater
import runtime
import thinkcell

def fixture():
    raw = b'<root><CPieChartSE id="1"><m_strName></m_strName><m_dtable idref="2"/></CPieChartSE><CPieChartDataTable id="2"><m_bExcelOnTop val="0"/><m_advisesink idref="0"/></CPieChartDataTable><unrelated val="keep"/></root>'
    root = E.fromstring(raw)
    ids = {n.get('id'):n for n in root if n.get('id')}
    return {'owner':ids['1'],'table':ids['2'],'owner_name':'','doc':{'root':root,'streams':{('think-cellXML',):raw},'slide_id':256,'slide_number':1,'ids':ids},
            'exact':True,'frames':[{'shape_id':21,'shape_tag':'stable-tag'}]}

class NamingTests(unittest.TestCase):
    def test_default_naming_changes_only_two_name_fields(self):
        c = fixture()
        original = E.tostring(c['doc']['root'])
        name,reused = namer.name_plan(c,[],'A'*64)
        data,changes = namer.rewrite(c,name,reused)
        self.assertFalse(reused)
        self.assertEqual(len(changes),2)
        self.assertEqual(E.tostring(c['doc']['root']),original)
        namer.verify_diff(c,data,name,False)
        self.assertEqual(E.fromstring(data).find('unrelated').get('val'),'keep')

    def test_hidden_data_change_is_rejected(self):
        c=fixture();name,reused=namer.name_plan(c,[],'A'*64)
        data,_=namer.rewrite(c,name,reused)
        changed=data.replace(b'val="keep"',b'val="changed"')
        with self.assertRaisesRegex(ValueError,'Non-name XML'):
            namer.verify_diff(c,changed,name,False)

    def test_case_insensitive_duplicate_names_rejected(self):
        with self.assertRaisesRegex(ValueError,'Duplicate automation names'):
            namer.name_plan(fixture(),[{'name':'Sales'},{'name':'sales'}],'A'*64)

    def test_existing_name_reused_without_serialization(self):
        c=fixture();c['owner'].find('m_strName').text='Existing'
        E.SubElement(c['table'],'m_strName').text='Existing'
        c['doc']['streams'][('think-cellXML',)]=E.tostring(c['doc']['root'])
        name,reused=namer.name_plan(c,[{'name':'Existing'}],'A'*64)
        data,changes=namer.rewrite(c,name,reused)
        self.assertTrue(reused);self.assertEqual(changes,[])
        self.assertEqual(data,c['doc']['streams'][('think-cellXML',)])

    def test_wrong_shape_selector_rejected(self):
        with self.assertRaisesRegex(ValueError,'selectors disagree'):
            namer.choose([fixture()],slide_number=1,shape_id=22,shape_tag='stable-tag')

    def test_unknown_external_link_rejected(self):
        c=fixture();c['table'].find('m_advisesink').set('idref','99')
        with self.assertRaisesRegex(ValueError,'external link'):
            namer.link_contract(c)

class DataTests(unittest.TestCase):
    def test_booleans_and_nonfinite_cells_rejected(self):
        for value in (True,float('nan'),float('inf')):
            with self.subTest(value=value),self.assertRaises(ValueError):
                updater.validate_request({'matrix':[[None,None],['A',value]],'expected_model':{'categories':['A'],'values':[value]}},'CPieChartSE')

    def test_pie_category_rows_accept_horizontal_source(self):
        c=fixture();c['doc']['streams'][('store','Package')]=b'fake'
        request={'matrix':[[None,None],['A',65],['B',35]],'expected_model':{'categories':['A','B'],'values':[65,35]}}
        cells=[{'row':1,'column':2,'value':'A'},{'row':1,'column':3,'value':'B'},{'row':2,'column':2,'value':60},{'row':2,'column':3,'value':40}]
        with patch.object(updater,'model_of',return_value={'categories':['A','B'],'values':[60,40]}),patch.object(updater,'link_contract',return_value={'storage':'store'}),patch.object(updater,'read_blob',return_value=[{'nonempty_cells':cells}]):
            result=updater.canonical_pie_contract(c,request)
            self.assertEqual(result['json_orientation'],'category_rows_label_value_columns')
            wrong=copy.deepcopy(request);wrong['matrix']=[[None,'A','B'],[None,65,35]]
            with self.assertRaisesRegex(ValueError,'category/value rows'):
                updater.canonical_pie_contract(c,wrong)
            changed=copy.deepcopy(request);changed['matrix'][0]=['New header',None]
            with self.assertRaisesRegex(ValueError,'header row changed'):
                updater.canonical_pie_contract(c,changed)

class HandoffTests(unittest.TestCase):
    def test_finalization_follows_identity_after_slide_move(self):
        plan={'bound_slide_id':'stable','tool_name':'edit_slide_ooxml','candidate_sha256':'digest','replacement_template':{'summary':'Return slide','code':'markDirty();'}}
        result=handoff.finalize(plan,{'slides':[{'id':'other','slideIndex':0},{'id':'stable','slideIndex':4}]})
        self.assertEqual(result['replacement_args']['slide_index'],4)
        with self.assertRaises(ValueError):
            handoff.finalize(plan,{'slides':[{'id':'missing','slideIndex':4}]})

    def test_multi_slide_candidate_rejected(self):
        stream=io.BytesIO()
        with zipfile.ZipFile(stream,'w') as z:
            z.writestr('ppt/presentation.xml','<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId id="1"/><p:sldId id="2"/></p:sldIdLst></p:presentation>')
            z.writestr('ppt/slides/slide1.xml','<slide/>')
        data=stream.getvalue()
        with self.assertRaisesRegex(ValueError,'exactly one slide'):
            handoff.prepare(data,hashlib.sha256(data).hexdigest(),'stable',0)

class RuntimeTests(unittest.TestCase):
    def test_invalid_explicit_executable_does_not_fall_back(self):
        with patch.dict(os.environ,{'THINKCELL_PPTTC':'/nonexistent-explicit/ppttc.exe'}):
            self.assertIsNone(runtime.find_ppttc())

    def test_report_cannot_alias_input(self):
        with tempfile.TemporaryDirectory() as folder:
            base=Path(folder);src=base/'input.pptx';src.write_bytes(b'sealed')
            a=argparse.Namespace(input=src,output=base/'output.pptx',data_json=base/'request.json',report=src,expected_sha256='0'*64)
            with self.assertRaisesRegex(ValueError,'Report must be a separate'):
                thinkcell.update(a)
            self.assertEqual(src.read_bytes(),b'sealed')

if __name__ == '__main__':
    unittest.main()
