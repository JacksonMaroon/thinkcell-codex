"""Object assembly must retain native carriers, bindings and untouched siblings."""
from pathlib import Path
import io
import json
import sys
import tempfile
import unittest
import zipfile
from unittest.mock import patch

SCRIPTS = Path(__file__).resolve().parents[1] / "skills/thinkcell-edit/scripts"
sys.path.insert(0, str(SCRIPTS))
import ordinary_slide_edits as edits
import assemble_native_slides as assembly


def shape(identity, text="Original", extra=""):
    return f'<p:sp><p:nvSpPr><p:cNvPr id="{identity}" name="Object {identity}"/><p:cNvSpPr/><p:nvPr>{extra}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400" b="1"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>'


def slide(*shapes):
    return (f'<p:sld xmlns:p="{edits.P}" xmlns:a="{edits.A}" xmlns:r="{edits.R}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + ''.join(shapes) + '</p:spTree></p:cSld></p:sld>').encode()


def selector(raw, identity, op="set", **fields):
    objects = {item["shape_id"]: item for item in edits.inspect_shapes(raw)}
    return dict(op=op, shape_id=identity, sha256=objects[identity]["sha256"], **fields)


def addition(identity=10):
    return {"op": "add", "shape_id": identity, "name": "New editable title", "kind": "textbox",
            "bounds_emu": {"left": 1000, "top": 1000, "width": 1000000, "height": 500000},
            "text": "New title", "font": {"name": "Arial", "size": 1800, "color": "123456", "bold": True},
            "fill": None, "line": None}


class OrdinaryAssemblyTests(unittest.TestCase):
    def test_text_geometry_remove_add_and_native_sibling_exact(self):
        raw = slide(shape(2), shape(3, "Remove me"), shape(4, "Native label", '<p:tags r:id="native-tag"/>'))
        protected = edits._tree(raw)[2][4]
        result = edits.edit_slide(raw, [selector(raw, 2, text_runs=["Updated"], bounds_emu={"left": 50, "top": 60, "width": 70, "height": 80}), selector(raw, 3, "remove"), addition()])
        _, _, shapes, _ = edits._tree(result)
        self.assertEqual(edits.shape_sha(protected), edits.shape_sha(shapes[4]))
        self.assertNotIn(3, shapes)
        self.assertEqual(shapes[2].find('.//a:t', edits.NS).text, "Updated")
        self.assertEqual(shapes[2].find('.//a:rPr', edits.NS).get('b'), '1')
        self.assertEqual(shapes[2].find('.//a:off', edits.NS).get('x'), '50')
        self.assertEqual(shapes[10].find('.//a:t', edits.NS).text, 'New title')

    def test_all_non_slide_parts_byte_identical(self):
        raw = slide(shape(2), shape(3, 'Native', '<p:tags r:id="tag"/>'))
        buffer = io.BytesIO()
        parts = {'ppt/slides/slide1.xml': raw, 'ppt/embeddings/native.bin': b'native authoring model',
                 'ppt/tags/tag1.xml': b'<tags/>', 'ppt/slides/_rels/slide1.xml.rels': b'native chart relationships',
                 'ppt/notesSlides/notesSlide1.xml': b'unchanged notes', 'ppt/media/image1.png': b'unchanged image'}
        with zipfile.ZipFile(buffer, 'w') as package:
            for name, data in parts.items(): package.writestr(name, data)
        changed = edits.prepare_source(buffer.getvalue(), [selector(raw, 2, text_runs=['Changed'])], 'ppt/slides/slide1.xml')
        with zipfile.ZipFile(io.BytesIO(changed)) as package:
            for name, data in parts.items():
                if name != 'ppt/slides/slide1.xml': self.assertEqual(package.read(name), data)

    def test_native_field_relationship_and_group_rejected(self):
        variants = [shape(2, extra='<p:tags r:id="tag"/>'),
                    shape(2).replace('<a:r>', '<a:fld id="native-field">').replace('</a:r>', '</a:fld>'),
                    shape(2).replace('<a:rPr sz="1400" b="1"/>', '<a:rPr><a:hlinkClick r:id="link"/></a:rPr>'),
                    shape(2).replace('<p:sp>', '<p:grpSp>').replace('</p:sp>', '</p:grpSp>')]
        for variant in variants:
            raw = slide(variant)
            with self.subTest(variant=variant), self.assertRaises(ValueError):
                edits.edit_slide(raw, [selector(raw, 2, text_runs=['Unsafe'])])

    def test_ordinary_placeholder_text_keeps_inheritance_and_cannot_be_removed(self):
        raw = slide(shape(2, extra='<p:ph type="title"/>'))
        output = edits.edit_slide(raw, [selector(raw, 2, text_runs=['New title'])])
        self.assertEqual(edits._tree(output)[2][2].find('.//p:ph', edits.NS).get('type'), 'title')
        with self.assertRaisesRegex(ValueError, 'cannot be removed'):
            edits.edit_slide(raw, [selector(raw, 2, 'remove')])

    def test_hash_wrong_input_duplicate_id_duplicate_selector_and_collision_rejected(self):
        raw = slide(shape(2), shape(3))
        wrong = selector(raw, 2, text_runs=['Changed']); wrong['sha256'] = '0' * 64
        repeated = selector(raw, 2, text_runs=['Changed'])
        for operations in ([wrong], [repeated, repeated], [addition(2)], [addition(), addition()],
                           [selector(raw, 2, text_runs=[])], [selector(raw, 2, bounds_emu={'left': True, 'top': 0, 'width': 1, 'height': 1})]):
            with self.subTest(operations=operations), self.assertRaises(ValueError): edits.edit_slide(raw, operations)
        with self.assertRaisesRegex(ValueError, 'Duplicate shape IDs'):
            edits.edit_slide(slide(shape(2), shape(2)), [])
        stale = selector(raw, 2, text_runs=['Changed'])
        with self.assertRaisesRegex(ValueError, 'precondition'):
            edits.edit_slide(slide(shape(2, 'Raced'), shape(3)), [stale])

    def test_connector_target_cannot_be_removed(self):
        raw = slide(shape(2), '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="connector"/><p:cNvCxnSpPr><a:stCxn id="2" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr></p:cxnSp>')
        with self.assertRaisesRegex(ValueError, 'connector'):
            edits.edit_slide(raw, [selector(raw, 2, 'remove')])

    def test_ole_fallback_identity_is_not_a_second_slide_object(self):
        carrier = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="native carrier"/></p:nvGraphicFramePr><a:graphic><a:graphicData><p:oleObj><p:pic><p:nvPicPr><p:cNvPr id="7" name="fallback"/></p:nvPicPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>'
        raw = slide(shape(2), carrier)
        output = edits.edit_slide(raw, [selector(raw, 2, text_runs=['Updated'])])
        self.assertEqual(edits.shape_sha(edits._tree(raw)[2][7]), edits.shape_sha(edits._tree(output)[2][7]))
        with self.assertRaisesRegex(ValueError, 'Duplicate shape IDs'):
            edits.edit_slide(slide(shape(7), carrier), [])

    def test_office_creation_records_allowed_and_unknown_extensions_protected(self):
        known = shape(2).replace('<p:cNvSpPr/>', '<p:cNvSpPr/><p:extLst><p:ext uri="known"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1"/></p:ext></p:extLst>')
        raw = slide(known)
        self.assertIsNone(edits.inspect_shapes(raw)[1]['protected_reason'])
        edits.edit_slide(raw, [selector(raw, 2, text_runs=['Updated'])])
        unknown = slide(known.replace('modId', 'unknownBinding'))
        with self.assertRaisesRegex(ValueError, 'Unknown extension'):
            edits.edit_slide(unknown, [selector(unknown, 2, text_runs=['Updated'])])

    def test_timing_target_cannot_be_removed(self):
        raw = slide(shape(2)).replace(b'</p:sld>', b'<p:timing><p:spTgt spid="2"/></p:timing></p:sld>')
        with self.assertRaisesRegex(ValueError, 'timing'):
            edits.edit_slide(raw, [selector(raw, 2, 'remove')])

    def test_output_snapshot_rejects_duplicate_physical_ids(self):
        def package(raw):
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, 'w') as archive:
                archive.writestr('ppt/presentation.xml', f'<p:presentation xmlns:p="{edits.P}" xmlns:r="{edits.R}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>')
                archive.writestr('ppt/_rels/presentation.xml.rels', f'<Relationships xmlns="{assembly.REL}"><Relationship Id="rId1" Type="{edits.R}/slide" Target="slides/slide1.xml"/></Relationships>')
                archive.writestr('ppt/slides/slide1.xml', raw)
            return buffer.getvalue()
        valid = package(slide(shape(2), shape(3)))
        self.assertEqual(assembly.snapshot(valid)['slide_count'], 1)
        with self.assertRaisesRegex(ValueError, 'Duplicate shape IDs'):
            assembly.snapshot(package(slide(shape(2), shape(2))))

    def test_ordinary_chart_cache_mutation_cannot_pass_physical_or_closure_checks(self):
        def package(value):
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, 'w') as archive:
                archive.writestr('ppt/slides/slide1.xml', slide('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Ordinary chart"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="chartRel"/></a:graphicData></a:graphic></p:graphicFrame>'))
                archive.writestr('ppt/slides/_rels/slide1.xml.rels', f'<Relationships xmlns="{assembly.REL}"><Relationship Id="chartRel" Type="{edits.R}/chart" Target="../charts/chart1.xml"/></Relationships>')
                archive.writestr('ppt/charts/chart1.xml', f'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:numCache><c:pt idx="0"><c:v>{value}</c:v></c:pt></c:numCache></c:chartSpace>')
            return buffer.getvalue()
        with zipfile.ZipFile(io.BytesIO(package(1))) as before, zipfile.ZipFile(io.BytesIO(package(999))) as after:
            self.assertNotEqual(assembly.physical_snapshot(before, 'ppt/slides/slide1.xml'), assembly.physical_snapshot(after, 'ppt/slides/slide1.xml'))
            self.assertNotEqual(assembly.dependencies(before, 'ppt/slides/slide1.xml'), assembly.dependencies(after, 'ppt/slides/slide1.xml'))

    def test_source_preflight_reads_bound_bytes_and_remains_readonly(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory); source = root/'source.pptx'; source.write_bytes(b'bound source')
            manifest = root/'request.json'
            manifest.write_text(json.dumps({'schema': 'tc.slide-sequence.v1', 'slides': [
                {'id': 'a', 'path': source.name, 'sha256': assembly.sha(source), 'edits': [addition()]},
                {'id': 'b', 'path': source.name, 'sha256': assembly.sha(source)}]}))
            state = {'slide_count': 1, 'dimensions': (), 'physical': 'source', 'notes': None, 'carriers': ['native'], 'dependencies': {'native': 1}}
            with patch.object(assembly, 'snapshot', return_value=state) as snapshots, patch.object(assembly, '_prepared', return_value=b'edited source'):
                result = assembly.run(manifest, assembly.sha(manifest), root/'out.pptx', root/'report.json')
                self.assertFalse(result['writes'])
                self.assertIn('ordinary object edits', result['scope'])
                self.assertTrue(all(isinstance(call.args[0], bytes) for call in snapshots.call_args_list))
                self.assertEqual(len(list(root.iterdir())), 2)

    def test_preflight_rejects_native_dependency_change(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory); source = root/'source.pptx'; source.write_bytes(b'source')
            manifest = root/'request.json'
            manifest.write_text(json.dumps({'schema': 'tc.slide-sequence.v1', 'slides': [
                {'id': 'a', 'path': source.name, 'sha256': assembly.sha(source), 'edits': [addition()]},
                {'id': 'b', 'path': source.name, 'sha256': assembly.sha(source)}]}))
            state = {'slide_count': 1, 'dimensions': (), 'physical': 'source', 'notes': None, 'carriers': ['native'], 'dependencies': {'native': 1}}
            with patch.object(assembly, 'snapshot', side_effect=[state, dict(state, dependencies={'native': 2})]), patch.object(assembly, '_prepared', return_value=b'changed'):
                with self.assertRaisesRegex(ValueError, 'native/dependency'): assembly.preflight(manifest, assembly.sha(manifest))

    def test_source_race_between_preflight_and_clone_withholds_output_before_office(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory); source = root/'source.pptx'; source.write_bytes(b'source')
            manifest = root/'request.json'
            manifest.write_text(json.dumps({'schema': 'tc.slide-sequence.v1', 'slides': [
                {'id': 'a', 'path': source.name, 'sha256': assembly.sha(source)},
                {'id': 'b', 'path': source.name, 'sha256': assembly.sha(source)}]}))
            state = {'slide_count': 1, 'dimensions': (), 'physical': 'source', 'notes': None, 'carriers': [], 'dependencies': {}}
            real_preflight = assembly.preflight
            def race(*args):
                result = real_preflight(*args)
                source.write_bytes(b'replaced after preflight')
                return result
            with patch.object(assembly, 'snapshot', return_value=state), patch.object(assembly, 'preflight', side_effect=race), patch.object(assembly, 'run_locked_subprocess') as native:
                with self.assertRaisesRegex(ValueError, 'changed while cloning'):
                    assembly.run(manifest, assembly.sha(manifest), root/'out.pptx', root/'report.json', execute=True, ppttc=sys.executable)
                native.assert_not_called()
                self.assertFalse((root/'out.pptx').exists())
                self.assertFalse((root/'report.json').exists())


if __name__ == '__main__': unittest.main()
