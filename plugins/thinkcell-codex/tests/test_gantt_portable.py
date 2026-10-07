"""Synthetic Gantt contracts, independent of Office and proprietary donors."""
from pathlib import Path
import hashlib
import json
import re
import sys
import tempfile
import unittest
from unittest.mock import patch
from xml.etree import ElementTree as ET
import zipfile

SCRIPTS = Path(__file__).resolve().parents[1] / 'skills/thinkcell-edit/scripts'
sys.path[:0] = [str(SCRIPTS/'gantt'), str(SCRIPTS)]
from portable_gantt_adapter import GanttPackage, require_distinct_paths
from milestone_adapter import Package
from verify_prepared_candidate import verify
import milestone_adapter
import run_milestone_edit
import run_paired_gantt_edit


def fixture(path, width=2000, marker_width=5):
    # IDs and carrier names deliberately differ from the historical donor.
    model = '<root><CGanttSE id="owner"><m_bConsistent val="1"/></CGanttSE>'
    model += '<CGanttVector id="vector"><m_cganttbar><elem idref="bar"/></m_cganttbar></CGanttVector>'
    model += '<CGanttBar id="bar"><m_datetime val="2026-01-05T00:00:00"/><m_datetime val="2026-01-11T00:00:00"/><m_pptgenline idref="gen"/></CGanttBar>'
    model += '<CPPTGenericLine id="gen"><m_pptautoshpline idref="line"/></CPPTGenericLine>'
    model += '<CPPTAutoShapeLine id="line"><m_bstrShapeName>bar-tag</m_bstrShapeName><m_rectPPTShape left="100" top="10" right="200" bottom="20"/></CPPTAutoShapeLine>'
    model += f'<CGanttMilestone id="milestone"><m_datetime val="2026-01-05T00:00:00"/><m_emarkerstyle val="7"/><m_pptautoshpmarker><m_bstrShapeName>marker-tag</m_bstrShapeName><m_rectPPTShape left="98" top="30" right="{98+marker_width}" bottom="35"/></m_pptautoshpmarker></CGanttMilestone>'
    model += '<CGanttScaleWeeks id="axis"><elem idref="box1"/><elem idref="box2"/></CGanttScaleWeeks>'
    for i, dt in enumerate(['2026-01-05', '2026-01-12'], 1):
        model += f'<CPPTGanttScaleBox id="box{i}"><m_rectPPTShape left="{i*100}" top="0" right="{(i+1)*100}" bottom="10"/><m_varsrc idref="date{i}"/></CPPTGanttScaleBox><CVariableSource id="date{i}"><m_datetime val="{dt}T00:00:00"/></CVariableSource>'
    model += '<Unrelated id="keep" value="preserve"/></root>'
    shapes=[]
    for i, (name, left, size) in enumerate([('bar-tag', 5000, width), ('marker-tag', 4960, 100)], 1):
        shapes.append(f'<p:sp><p:cNvPr id="{i}" name="arbitrary{i}"/><p:tags r:id="tag{i}"/><a:off x="{left}" y="10"/><a:ext cx="{size}" cy="100"/><a:prstGeom prst="triangle"/></p:sp>')
    slide='<p:sld xmlns:p="urn:p" xmlns:a="urn:a" xmlns:r="urn:r">'+''.join(shapes)+'</p:sld>'
    rels='<Relationships>'+''.join(f'<Relationship Id="tag{i}" Target="../tags/tag{i}.xml"/>' for i in [1,2])+'</Relationships>'
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('ppt/embeddings/random-carrier.bin', model)
        z.writestr('ppt/slides/slide1.xml', slide)
        z.writestr('ppt/slides/_rels/slide1.xml.rels', rels)
        for i, name in enumerate(['bar-tag','marker-tag'], 1):
            z.writestr(f'ppt/tags/tag{i}.xml', f'<tag name="THINKCELLSHAPEDONOTDELETE" val="{name}"/>')


def fake_replace(args, **kwargs):
    # Exercise all package transformations while substituting the platform OLE
    # stream writer with a raw XML carrier for this synthetic fixture only.
    Path(args[args.index('-StoragePath')+1]).write_bytes(Path(args[args.index('-StreamBytesPath')+1]).read_bytes())
    return type('Result', (), {'returncode':0})()


class GanttContracts(unittest.TestCase):
    def setUp(self):
        # Native executable discovery stays enforced in production; synthetic
        # stream-write tests must not require Windows or installed PowerShell.
        for target in ['portable_gantt_adapter.powershell', 'milestone_adapter.powershell']:
            discovery=patch(target,return_value='synthetic-powershell')
            discovery.start()
            self.addCleanup(discovery.stop)
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.source=Path(self.tmp.name)/'source.pptx'; fixture(self.source)
        self.out=Path(self.tmp.name)/'output.pptx'

    def test_reversed_dates_refused(self):
        with self.assertRaisesRegex(ValueError, 'start date'):
            GanttPackage(self.source).edit_bar(self.out, selector={'bar_id':'bar'}, new_start='2026-01-12', new_end='2026-01-05')
        self.assertFalse(self.out.exists())

    def test_source_overwrite_refused(self):
        with self.assertRaisesRegex(ValueError, 'fresh path'):
            Package(self.source).edit(self.source, selector={'milestone_id':'milestone'}, new_date='2026-01-12')

    def test_existing_output_refused(self):
        self.out.write_bytes(b'preserve')
        with self.assertRaisesRegex(ValueError, 'fresh path'):
            GanttPackage(self.source).edit_bar(self.out, selector={'bar_id':'bar'}, new_start='2026-01-05', new_end='2026-01-12')
        self.assertEqual(self.out.read_bytes(), b'preserve')

    def test_calendar_outside_scale_refused(self):
        with self.assertRaisesRegex(ValueError, 'outside'):
            GanttPackage(self.source).x_for('2027-01-01')

    def test_ambiguous_selector_refused(self):
        with self.assertRaisesRegex(ValueError, 'matched 0'):
            Package(self.source).resolve(milestone_id='unknown')

    def test_multislide_refused_before_mutation(self):
        with zipfile.ZipFile(self.source, 'a') as z:
            z.writestr('ppt/slides/slide2.xml', '<slide/>')
        with self.assertRaisesRegex(ValueError, 'one-slide'):
            Package(self.source).edit(self.out, selector={'milestone_id':'milestone'}, new_date='2026-01-12')
        self.assertFalse(self.out.exists())

    def test_unproven_marker_shape_refused(self):
        package=Package(self.source)
        package.blobs['ppt/slides/slide1.xml']=package.blobs['ppt/slides/slide1.xml'].replace(b'prst="triangle"', b'prst="diamond"')
        with self.assertRaisesRegex(ValueError, 'triangle'):
            package.edit(self.out, selector={'milestone_id':'milestone'}, new_date='2026-01-12')
        self.assertFalse(self.out.exists())

    def test_duplicate_physical_bar_tag_refused_before_output(self):
        package=GanttPackage(self.source)
        slide=package.blobs['ppt/slides/slide1.xml'].decode()
        shape=re.findall(r'<p:sp>.*?</p:sp>',slide)[0]
        package.blobs['ppt/slides/slide1.xml']=slide.replace('</p:sld>',shape+'</p:sld>').encode()
        with self.assertRaisesRegex(ValueError,'matched 2 physical shapes'):
            package.edit_bar(self.out,selector={'bar_id':'bar'},new_start='2026-01-12',new_end='2026-01-12')
        self.assertFalse(self.out.exists())

    def test_duplicate_physical_milestone_tag_refused_before_output(self):
        package=Package(self.source)
        slide=package.blobs['ppt/slides/slide1.xml'].decode()
        shape=re.findall(r'<p:sp>.*?</p:sp>',slide)[1]
        package.blobs['ppt/slides/slide1.xml']=slide.replace('</p:sld>',shape+'</p:sld>').encode()
        with self.assertRaisesRegex(ValueError,'matched 2 physical shapes'):
            package.edit(self.out,selector={'milestone_id':'milestone'},new_date='2026-01-12')
        self.assertFalse(self.out.exists())

    def test_irregular_weekly_dates_refused(self):
        package=GanttPackage(self.source)
        package.model=package.model.replace('2026-01-12T', '2026-01-13T')
        with self.assertRaisesRegex(ValueError, 'seven-day'):
            package._discover_axis()

    def test_all_cli_report_output_aliases_refused_before_edit(self):
        digest=hashlib.sha256(self.source.read_bytes()).hexdigest()
        for module in [run_paired_gantt_edit,run_milestone_edit,milestone_adapter]:
            args=['gantt','--source',str(self.source),'--output',str(self.out),'--report',str(self.out.parent/'unused'/ '..'/self.out.name),'--expected-sha256',digest]
            if module is run_paired_gantt_edit:
                args += ['--select-start','2026-01-05','--select-end','2026-01-11','--start-date','2026-01-12','--end-date','2026-01-12']
            else:
                args += ['--milestone-id','milestone','--new-date','2026-01-12']
            with self.subTest(cli=module.__name__), patch.object(sys,'argv',args), patch('subprocess.run') as run:
                with self.assertRaisesRegex(ValueError,'paths must be distinct'):
                    module.main()
                run.assert_not_called()
            self.assertFalse(self.out.exists())
            self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(),digest)

    def test_three_way_path_alias_guard(self):
        for source,output,report in [(self.source,self.source,self.out),(self.source,self.out,self.source),(self.source,self.out,self.out)]:
            with self.subTest(paths=(source,output,report)), self.assertRaisesRegex(ValueError,'distinct'):
                require_distinct_paths(source,output,report)

    def test_report_created_during_prepare_not_overwritten(self):
        report=Path(self.tmp.name)/'report.json'
        def prepare(*args, **kwargs):
            report.write_bytes(b'external report')
            return {}
        args=['gantt','--source',str(self.source),'--output',str(self.out),'--report',str(report),'--expected-sha256',hashlib.sha256(self.source.read_bytes()).hexdigest(),'--select-start','2026-01-05','--select-end','2026-01-11','--start-date','2026-01-12','--end-date','2026-01-12']
        with patch.object(sys,'argv',args),patch.object(run_paired_gantt_edit,'paired_date_geometry_edit',side_effect=prepare):
            with self.assertRaises(FileExistsError): run_paired_gantt_edit.main()
        self.assertEqual(report.read_bytes(), b'external report')

    @patch('portable_gantt_adapter.subprocess.run', side_effect=fake_replace)
    def test_output_created_during_prepare_not_overwritten(self, run):
        def race(args, **kwargs):
            result=fake_replace(args,**kwargs)
            self.out.write_bytes(b'external file created during preparation')
            return result
        run.side_effect=race
        with self.assertRaises(FileExistsError):
            GanttPackage(self.source).edit_bar(self.out, selector={'bar_id':'bar'}, new_start='2026-01-12',new_end='2026-01-12')
        self.assertEqual(self.out.read_bytes(), b'external file created during preparation')

    @patch('milestone_adapter.subprocess.run', side_effect=fake_replace)
    def test_milestone_derived_scale_preserves_odd_width_and_valid_xml(self, run):
        edit = Package(self.source).edit(self.out, selector={'milestone_id':'milestone'}, new_date='2026-01-12')
        self.assertEqual(edit['new_visible_center'], 7010)
        self.assertEqual(edit['new_model_bounds'], [198,30,203,35])
        self.assertEqual(run.call_args.kwargs['timeout'],60)
        ET.fromstring(Package(self.out).model)
        self._readback(edit)

    @patch('portable_gantt_adapter.subprocess.run', side_effect=fake_replace)
    def test_bar_geometry_readback_and_non_target_preservation(self, run):
        edit=GanttPackage(self.source).edit_bar(self.out, selector={'bar_id':'bar'}, new_start='2026-01-12', new_end='2026-01-12')
        self.assertEqual(edit['new_visible_bounds'], [7000,9000])
        self._readback(edit)

    @patch('portable_gantt_adapter.subprocess.run', side_effect=fake_replace)
    def test_new_start_equal_old_end_preserves_both_date_fields(self, run):
        edit=GanttPackage(self.source).edit_bar(self.out,selector={'bar_id':'bar'},new_start='2026-01-11',new_end='2026-01-12')
        bar=GanttPackage(self.out).resolve_bar(bar_id='bar')
        self.assertEqual([bar.start,bar.end],['2026-01-11T00:00:00','2026-01-12T00:00:00'])
        self._readback(edit)

    @patch('portable_gantt_adapter.subprocess.run', side_effect=fake_replace)
    def test_new_end_equal_old_start_preserves_both_date_fields(self, run):
        edit=GanttPackage(self.source).edit_bar(self.out,selector={'bar_id':'bar'},new_start='2026-01-05',new_end='2026-01-05')
        bar=GanttPackage(self.out).resolve_bar(bar_id='bar')
        self.assertEqual([bar.start,bar.end],['2026-01-05T00:00:00','2026-01-05T00:00:00'])
        self._readback(edit)

    @patch('portable_gantt_adapter.subprocess.run', side_effect=fake_replace)
    def test_identical_original_dates_replace_distinct_spans(self, run):
        with zipfile.ZipFile(self.source) as z: blobs={n:z.read(n) for n in z.namelist()}
        blobs['ppt/embeddings/random-carrier.bin']=blobs['ppt/embeddings/random-carrier.bin'].replace(b'2026-01-11T00:00:00',b'2026-01-05T00:00:00')
        with zipfile.ZipFile(self.source,'w') as z:
            for name,raw in blobs.items():z.writestr(name,raw)
        edit=GanttPackage(self.source).edit_bar(self.out,selector={'bar_id':'bar'},new_start='2026-01-05',new_end='2026-01-12')
        bar=GanttPackage(self.out).resolve_bar(bar_id='bar')
        self.assertEqual([bar.start,bar.end],['2026-01-05T00:00:00','2026-01-12T00:00:00'])
        self._readback(edit)

    def _readback(self, edit):
        report=Path(self.tmp.name)/'report.json'
        report.write_text(json.dumps({'source':str(self.source),'output':str(self.out),'source_sha256_before':hashlib.sha256(self.source.read_bytes()).hexdigest(),'edit':edit}))
        result=verify(self.source,self.out,report)
        self.assertTrue(result['pass'], result)
        self.assertFalse(result['native_verified'])
        with zipfile.ZipFile(self.out) as z: blobs={n:z.read(n) for n in z.namelist()}
        blobs['ppt/embeddings/random-carrier.bin']=blobs['ppt/embeddings/random-carrier.bin'].replace(b'value="preserve"', b'value="corrupt"')
        with zipfile.ZipFile(self.out,'w') as z:
            for name,raw in blobs.items(): z.writestr(name,raw)
        self.assertFalse(verify(self.source,self.out,report)['pass'])


if __name__=='__main__':
    unittest.main()
