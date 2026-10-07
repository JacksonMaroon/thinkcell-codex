from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'skills/thinkcell-edit/scripts/gantt'))
import unittest
from endpoint_propagation import plan_endpoint_dates

D=lambda d:'2025-05-'+d+'T00:00:00'
class EndpointTests(unittest.TestCase):
    def setUp(self):
        self.bars=[{'bar_id':'1','dates':[D('01'),D('07')]},{'bar_id':'2','dates':[D('07'),D('21')]},{'bar_id':'3','dates':[D('07'),D('28')]}]
        self.edges=[('1','High','2','Low','4'),('1','High','3','Low','4')]
    def test_fanout_exact_equality_without_duration_assumption(self):
        p=plan_endpoint_dates(self.bars,self.edges,[('1','High',D('14'))])
        self.assertEqual(p,{'1':[D('01'),D('14')],'2':[D('14'),D('21')],'3':[D('14'),D('28')]})
    def test_second_change_from_previous_state(self):
        p=plan_endpoint_dates(self.bars,self.edges,[('1','High',D('14'))])
        bars=[{'bar_id':b['bar_id'],'dates':p[b['bar_id']]} for b in self.bars]
        q=plan_endpoint_dates(bars,self.edges,[('1','High',D('10'))])
        self.assertEqual(q['2'],[D('10'),D('21')])
    def test_no_change(self):
        self.assertEqual(plan_endpoint_dates(self.bars,self.edges,[('1','High',D('07'))]),{})
    def test_unrelated_endpoint_unchanged(self):
        self.assertEqual(plan_endpoint_dates(self.bars,self.edges,[('3','High',D('29'))]),{'3':[D('07'),D('29')]})
    def test_reject_conflicting_explicit_request(self):
        with self.assertRaisesRegex(ValueError,'conflicting'):
            plan_endpoint_dates(self.bars,self.edges,[('1','High',D('14')),('2','Low',D('15'))])
    def test_reject_bar_inversion(self):
        with self.assertRaisesRegex(ValueError,'invert'):
            plan_endpoint_dates(self.bars,self.edges,[('1','High',D('22'))])
    def test_reject_unsupported_affected_shape(self):
        with self.assertRaisesRegex(ValueError,'supported bar'):
            plan_endpoint_dates(self.bars,[('1','High','99','Low','4')],[('1','High',D('14'))])
    def test_reject_other_relationship(self):
        with self.assertRaisesRegex(ValueError,'FS only'):
            plan_endpoint_dates(self.bars,[('1','High','2','High','4')],[('1','High',D('14'))])
    def test_reject_detaching_bound_child_endpoint(self):
        with self.assertRaisesRegex(ValueError,'break an existing anchor'):
            plan_endpoint_dates(self.bars,self.edges,[('2','Low',D('10'))])

import copy,json,tempfile,hashlib,zipfile
from unittest.mock import patch,MagicMock
import endpoint_propagation as adapter

class EndpointGradeTests(unittest.TestCase):
    def setUp(self):
        self.folder=tempfile.TemporaryDirectory();self.addCleanup(self.folder.cleanup)
        self.prepared=Path(self.folder.name)/'prepared.pptx';self.prepared.write_bytes(b'synthetic prepared')
        self.native=Path(self.folder.name)/'native.pptx';self.native.write_bytes(b'synthetic native')
        def bar(bid,lo,hi,x,width):
            return {'bar_id':str(bid),'dates':[D(lo),D(hi)],'date_guids':[str(bid)+'-start',str(bid)+'-end'],'visible':True,'model_rect':[x,0,x+width,10],'physical_transform':[x,0,width,10]}
        self.state={'consistent':True,'bars':[bar(1,'01','07',0,10),bar(2,'07','21',10,20),bar(3,'10','28',40,25),bar(4,'01','02',70,10)],'edges':[('1','High','2','Low','9')]}
        self.semantic=({'barDateSources':('synthetic',)}, {'worksheet':'samehash','package:sibling.bin':'samehash'}, [('parent','High','child','Low')])
    def evaluate(self,actual,semantic=None):
        with patch.object(adapter,'inspect_native_anchors',side_effect=[self.state,actual]),patch.object(adapter,'native_semantics',side_effect=[self.semantic,semantic or self.semantic]):
            return adapter.grade(self.prepared,self.native,'1-end')
    def test_matching_saved_endpoint_passes(self):
        self.assertEqual(self.evaluate(copy.deepcopy(self.state))['status'],'ADAPTER_ENDPOINT_NATIVE_PASS')
    def test_equal_pair_moved_away_from_request_rejected(self):
        actual=copy.deepcopy(self.state);actual['bars'][0]['dates'][1]=D('14');actual['bars'][1]['dates'][0]=D('14')
        result=self.evaluate(actual)
        self.assertTrue(result['checks']['exact_endpoint_date_equality'])
        self.assertFalse(result['checks']['all_exact_native_dates_match_request'])
    def test_unrelated_bar_model_changed_rejected(self):
        actual=copy.deepcopy(self.state);actual['bars'][2]['model_rect'][0]+=3
        self.assertFalse(self.evaluate(actual)['checks']['all_bar_native_rects_preserved'])
    def test_unrelated_bar_physical_changed_rejected(self):
        actual=copy.deepcopy(self.state);actual['bars'][2]['physical_transform'][0]+=3
        self.assertFalse(self.evaluate(actual)['checks']['all_bar_physical_transforms_preserved'])
    def test_visibility_change_rejected(self):
        actual=copy.deepcopy(self.state);actual['bars'][2]['visible']=False
        self.assertFalse(self.evaluate(actual)['checks']['all_bar_visibility_preserved'])
    def test_duplicate_source_guid_rejected(self):
        actual=copy.deepcopy(self.state);actual['bars'][2]['date_guids'][0]='1-start'
        with self.assertRaisesRegex(ValueError,'duplicate'):self.evaluate(actual)
    def test_nonbar_date_source_changed_rejected(self):
        dates,streams,graph=self.semantic
        result=self.evaluate(copy.deepcopy(self.state),({'milestone':'changed'},streams,graph))
        self.assertFalse(result['checks']['all_native_date_sources_preserved'])
    def test_equal_count_unrelated_graph_retarget_rejected(self):
        dates,streams,graph=self.semantic
        result=self.evaluate(copy.deepcopy(self.state),(dates,streams,[('parent','High','other','Low')]))
        self.assertFalse(result['checks']['all_anchor_semantics_preserved'])
    def test_sibling_embedding_change_rejected(self):
        dates,streams,graph=self.semantic
        result=self.evaluate(copy.deepcopy(self.state),(dates,{**streams,'package:sibling.bin':'changed'},graph))
        self.assertFalse(result['checks']['unrelated_datasheet_stream_hashes_equal'])
    def test_worksheet_stream_change_rejected(self):
        dates,streams,graph=self.semantic
        result=self.evaluate(copy.deepcopy(self.state),(dates,{**streams,'worksheet':'changed'},graph))
        self.assertEqual(result['status'],'ADAPTER_ENDPOINT_NATIVE_FAIL')
    def test_multiple_selected_children_outside_verified_profile(self):
        actual=copy.deepcopy(self.state);actual['edges'].append(('1','High','3','Low','9'))
        with self.assertRaisesRegex(ValueError,'finish/start'):self.evaluate(actual)

class EndpointArtifactGuards(unittest.TestCase):
    def setUp(self):
        self.folder=tempfile.TemporaryDirectory();self.addCleanup(self.folder.cleanup)
        self.root=Path(self.folder.name);self.source=self.root/'source.pptx';self.source.write_bytes(b'original')
    def test_path_alias_rejected_before_mutation(self):
        with self.assertRaisesRegex(ValueError,'distinct'):
            adapter.guard_paths([self.source,self.root/'child'/ '..'/'source.pptx'],fresh=[])
    def test_existing_sidecar_rejected(self):
        report=self.root/'report.json';report.write_text('keep')
        with self.assertRaisesRegex(ValueError,'fresh'):
            adapter.guard_paths([self.source,report],fresh=[report])
        self.assertEqual(report.read_text(),'keep')
    def test_hash_mismatch_before_parser_or_write(self):
        with patch.object(adapter,'inspect_native_anchors') as inspect:
            with self.assertRaisesRegex(ValueError,'hash differs'):
                adapter.prepare(self.source,self.root/'out.pptx','end','2025-05-21',expected_sha256='0'*64,report=self.root/'report.json',job=self.root/'job.ppttc')
            inspect.assert_not_called()
        self.assertEqual(list(self.root.iterdir()),[self.source])
    def test_unsupported_schema_rejected_before_sheet_or_mutation(self):
        sha=hashlib.sha256(self.source.read_bytes()).hexdigest()
        with patch.object(adapter,'inspect_native_anchors',return_value={'consistent':True,'schema_version':'99999'}),patch.object(adapter,'authenticated_empty_sheet') as sheet:
            with self.assertRaisesRegex(ValueError,'schema'):
                adapter.prepare(self.source,self.root/'out.pptx','end','2025-05-21',expected_sha256=sha,report=self.root/'report.json',job=self.root/'job.ppttc')
            sheet.assert_not_called()
        self.assertEqual(list(self.root.iterdir()),[self.source])
    def test_unsupported_bar_count_rejected_before_sheet_or_mutation(self):
        sha=hashlib.sha256(self.source.read_bytes()).hexdigest()
        state={'consistent':True,'schema_version':'38790','bars':[{'visible':True} for _ in range(3)]}
        with patch.object(adapter,'inspect_native_anchors',return_value=state),patch.object(adapter,'authenticated_empty_sheet') as sheet:
            with self.assertRaisesRegex(ValueError,'four-visible'):
                adapter.prepare(self.source,self.root/'out.pptx','end','2025-05-21',expected_sha256=sha,report=self.root/'report.json',job=self.root/'job.ppttc')
            sheet.assert_not_called()
        self.assertEqual(list(self.root.iterdir()),[self.source])
    def test_reserved_sidecar_alias_rejected(self):
        with self.assertRaisesRegex(ValueError,'distinct'):
            adapter.prepare(self.source,self.root/'out.pptx','end','2025-05-21',expected_sha256='0'*64,report=self.root/'out.pptx',job=self.root/'job.ppttc')
    def test_competing_writer_not_overwritten_and_own_partial_removed(self):
        output=self.root/'out.pptx';job=self.root/'job.ppttc';job.write_bytes(b'competing writer')
        with self.assertRaises(FileExistsError):adapter.publish_files([(output,b'candidate'),(job,b'job')])
        self.assertFalse(output.exists());self.assertEqual(job.read_bytes(),b'competing writer')
    def test_source_changed_before_publication_leaves_nothing(self):
        output=self.root/'out.pptx'
        with self.assertRaisesRegex(ValueError,'source changed'):
            adapter.publish_files([(output,b'candidate')],source=self.source,expected_sha256='0'*64)
        self.assertFalse(output.exists())
    def test_forged_candidate_source_model_rejected(self):
        with patch.object(adapter,'_model_and_streams',side_effect=[('expected',{}),('forged',{})]):
            with self.assertRaisesRegex(ValueError,'reconstruction'):adapter.assert_candidate_matches(self.source,self.root/'forged.pptx')
    def test_forged_sibling_carrier_rejected(self):
        with patch.object(adapter,'_model_and_streams',side_effect=[('model',{'package:sibling.bin':b'expected'}),('model',{'package:sibling.bin':b'forged'})]):
            with self.assertRaisesRegex(ValueError,'reconstruction'):adapter.assert_candidate_matches(self.source,self.root/'forged.pptx')
    def test_forged_unrelated_slide_part_rejected(self):
        other=self.root/'candidate.pptx'
        for path,payload in [(self.source,b'expected'),(other,b'forged')]:
            with zipfile.ZipFile(path,'w') as z:z.writestr('ppt/slides/slide1.xml',payload)
        with patch.object(adapter,'_model_and_streams',return_value=('model',{})):
            with self.assertRaisesRegex(ValueError,'unrelated'):adapter.assert_candidate_matches(self.source,other)
    def test_preflight_rejects_nonempty_sheet(self):
        model='<root><CGanttTable><m_advisesink idref="0"/></CGanttTable></root>'
        workbook=MagicMock();workbook.sheets=['Sheet1'];sheet=MagicMock();sheet.rows.return_value=[[MagicMock(v=123)]]
        workbook.get_sheet.return_value.__enter__.return_value=sheet
        with patch.object(adapter,'_model_and_streams',return_value=(model,{'think-cellChild0/Package':b'synthetic'})),patch.object(adapter.pyxlsb,'open_workbook') as opened:
            opened.return_value.__enter__.return_value=workbook
            with self.assertRaisesRegex(ValueError,'empty'):adapter.authenticated_empty_sheet(self.source)
    def test_preflight_rejects_linked_empty_sheet(self):
        model='<root><CGanttTable><m_advisesink idref="77"/></CGanttTable></root>'
        with patch.object(adapter,'_model_and_streams',return_value=(model,{})):
            with self.assertRaisesRegex(ValueError,'linked'):adapter.authenticated_empty_sheet(self.source)
    def test_preflight_rejects_scalar_bound_sheet(self):
        model='<root><CGanttTable><m_advisesink idref="0"/></CGanttTable><CGanttScalarAnchor/></root>'
        with patch.object(adapter,'_model_and_streams',return_value=(model,{})):
            with self.assertRaisesRegex(ValueError,'scalar'):adapter.authenticated_empty_sheet(self.source)
    def test_midday_or_timezone_dates_rejected(self):
        bars=[{'bar_id':'1','dates':[D('01'),D('07')]}]
        for date in ['2025-05-14T12:00:00','2025-05-14T00:00:00+00:00','2025-05-14']:
            with self.assertRaisesRegex(ValueError,'midnight'):adapter.plan_endpoint_dates(bars,[],[('1','High',date)])

if __name__=='__main__':unittest.main()
