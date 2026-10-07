"""Exercise sequence source preservation and dependency binding without Office."""
from pathlib import Path
import json
import sys
import tempfile
import unittest
import zipfile
from unittest.mock import patch

SCRIPTS = Path(__file__).resolve().parents[1] / "skills/thinkcell-edit/scripts"
sys.path.insert(0, str(SCRIPTS))
import assemble_native_slides as assembly


class SequenceTests(unittest.TestCase):
    def test_preflight_validates_source_hashes_and_writes_nothing(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.pptx"; source.write_bytes(b"source")
            manifest = root / "manifest.json"
            manifest.write_text(json.dumps({"schema": "tc.slide-sequence.v1", "slides": [
                {"id": "a", "path": source.name, "sha256": assembly.sha(source)},
                {"id": "b", "path": source.name, "sha256": assembly.sha(source)}]}))
            state = {"dimensions": (("cx", "1"),), "slide_count": 1}
            with patch.object(assembly, "snapshot", return_value=state):
                result = assembly.run(manifest, assembly.sha(manifest), root / "out.pptx", root / "report.json")
                self.assertFalse(result["writes"])
                self.assertEqual(sorted(path.name for path in root.iterdir()), ["manifest.json", "source.pptx"])
                with self.assertRaisesRegex(ValueError, "alias"):
                    assembly.run(manifest, assembly.sha(manifest), source, root / "report.json")
                source.write_bytes(b"changed")
                with self.assertRaisesRegex(ValueError, "Source SHA"):
                    assembly.preflight(manifest, assembly.sha(manifest))

    def test_output_must_preserve_every_ordered_slide(self):
        wanted = {"dimensions": (), "physical": "before", "notes": None, "carriers": [], "dependencies": {}, "slide_count": 1}
        actual = dict(wanted, slide_count=2)
        with patch.object(assembly, "snapshot", return_value=actual):
            self.assertEqual(assembly.verify_output("output", [wanted, wanted])["slide_count"], 2)
        with patch.object(assembly, "snapshot", return_value=dict(actual, physical="changed")):
            with self.assertRaisesRegex(ValueError, "physical preservation"):
                assembly.verify_output("output", [wanted, wanted])

    def test_dependency_multiplicity_and_dangling_relationships(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "source.pptx"
            with zipfile.ZipFile(path, "w") as package:
                package.writestr("ppt/slides/slide1.xml", b"slide")
                package.writestr("ppt/slides/_rels/slide1.xml.rels", '<Relationships xmlns="'+assembly.REL+'"><Relationship Id="a" Target="../media/a.png"/><Relationship Id="b" Target="../media/b.png"/></Relationships>')
                package.writestr("ppt/media/a.png", b"same")
                package.writestr("ppt/media/b.png", b"same")
            with zipfile.ZipFile(path) as package:
                self.assertEqual(sum(assembly.dependencies(package, "ppt/slides/slide1.xml").values()), 2)
            with zipfile.ZipFile(path, "a") as package:
                package.writestr("ppt/media/_rels/a.png.rels", '<Relationships xmlns="'+assembly.REL+'"><Relationship Id="missing" Target="missing.png"/></Relationships>')
            with zipfile.ZipFile(path) as package:
                with self.assertRaisesRegex(ValueError,"Dangling"):
                    assembly.dependencies(package,"ppt/slides/slide1.xml")

    def test_physical_relationship_binding_and_text_spaces_preserved(self):
        def build(path, swap=False):
            with zipfile.ZipFile(path, "w") as package:
                package.writestr("ppt/slides/slide1.xml", '<p:sld xmlns:p="'+assembly.P+'" xmlns:r="'+assembly.R+'" xmlns:a="'+assembly.NS['a']+'"><p:cSld><p:spTree><p:sp><a:t> spaced </a:t><a:blip r:embed="a"/></p:sp></p:spTree></p:cSld></p:sld>')
                package.writestr("ppt/slides/_rels/slide1.xml.rels", '<Relationships xmlns="'+assembly.REL+'"><Relationship Id="a" Type="image" Target="../media/'+('b' if swap else 'a')+'.png"/></Relationships>')
                package.writestr("ppt/media/a.png", b"first")
                package.writestr("ppt/media/b.png", b"second")
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory); a=root/'a.pptx'; b=root/'b.pptx'; build(a); build(b,True)
            with zipfile.ZipFile(a) as package:
                before=assembly.physical_snapshot(package,'ppt/slides/slide1.xml')
            with zipfile.ZipFile(b) as package:
                after=assembly.physical_snapshot(package,'ppt/slides/slide1.xml')
            self.assertNotEqual(before,after)
            self.assertIn(' spaced ',repr(before))

    def test_creation_id_and_notes_master_datetime_cache_only(self):
        def write(path, creation, display, field_type="datetimeFigureOut", placeholder="dt", root_name="notesMaster"):
            with zipfile.ZipFile(path,"w") as package:
                package.writestr("ppt/notesMasters/notesMaster1.xml", f'<p:{root_name} xmlns:p="{assembly.P}" xmlns:a="{assembly.NS["a"]}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="{placeholder}"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:fld id="field-id" type="{field_type}"><a:rPr sz="1200"/><a:t>{display}</a:t></a:fld><a:r><a:t>Literal 9/15/2026</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:extLst><p:ext uri="fixed"><p14:creationId val="{creation}"/></p:ext><p:ext uri="mod"><p14:modId val="{creation}"/></p:ext><p:ext uri="another"><p14:creationId val="{creation}"/></p:ext></p:extLst></p:{root_name}>')
        def read(path):
            with zipfile.ZipFile(path) as package:
                return assembly.physical_snapshot(package,"ppt/notesMasters/notesMaster1.xml")
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory); before=root/'a.pptx'; after=root/'b.pptx'
            write(before,"1","9/15/2026"); write(after,"2","10/6/2026")
            self.assertEqual(read(before),read(after))
            self.assertIn("Literal 9/15/2026",repr(read(after)))
            write(after,"2","10/6/2026",field_type="custom-date")
            self.assertNotEqual(read(before),read(after))
            write(before,"1","9/15/2026",placeholder="body"); write(after,"2","10/6/2026",placeholder="body")
            self.assertNotEqual(read(before),read(after))
            write(before,"1","9/15/2026",root_name="sld"); write(after,"2","10/6/2026",root_name="sld")
            self.assertNotEqual(read(before),read(after))

    def test_cfb_binding_keeps_complete_model_while_ignoring_storage_header(self):
        magic=bytes.fromhex("D0CF11E0A1B11AE1")
        first={('think-cellXML',):b'<root><value val="1"/></root>',('Workbook',):b'exact data'}
        with patch.object(assembly,"streams",return_value=first):
            before=assembly._carrier_signature(magic+b'header-one',include_model=True)
            after=assembly._carrier_signature(magic+b'header-two',include_model=True)
        self.assertEqual(before,after)
        changed=dict(first);changed[('think-cellXML',)]=b'<root><value val="2"/></root>'
        with patch.object(assembly,"streams",return_value=changed):
            self.assertNotEqual(before,assembly._carrier_signature(magic,include_model=True))

    def test_hidden_dependency_native_feature_graph_mutation_rejected(self):
        magic=bytes.fromhex("D0CF11E0A1B11AE1")
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'hidden.pptx'
            with zipfile.ZipFile(path,'w') as package:
                package.writestr('ppt/slides/slide1.xml',b'slide')
                package.writestr('ppt/slides/_rels/slide1.xml.rels','<Relationships xmlns="'+assembly.REL+'"><Relationship Id="hidden" Target="../embeddings/hidden.bin"/></Relationships>')
                package.writestr('ppt/embeddings/hidden.bin',magic+b'carrier storage')
            source={('think-cellXML',):b'<root><CLabel id="2"><m_format val="0.0%"/></CLabel><CCAGR id="3"><m_source idref="1"/></CCAGR></root>',('Workbook',):b'unchanged data'}
            with zipfile.ZipFile(path) as package,patch.object(assembly,'streams',return_value=source):
                before=assembly.dependencies(package,'ppt/slides/slide1.xml')
            changed=dict(source);changed[('think-cellXML',)]=source[('think-cellXML',)].replace(b'0.0%',b'0%')
            with zipfile.ZipFile(path) as package,patch.object(assembly,'streams',return_value=changed):
                after=assembly.dependencies(package,'ppt/slides/slide1.xml')
            self.assertNotEqual(before,after)
            changed[('think-cellXML',)]=source[('think-cellXML',)].replace(b'idref="1"',b'idref="99"')
            with zipfile.ZipFile(path) as package,patch.object(assembly,'streams',return_value=changed):
                self.assertNotEqual(before,assembly.dependencies(package,'ppt/slides/slide1.xml'))

    def test_native_label_string_whitespace_is_semantic(self):
        first=assembly.E.fromstring(b'<root>\n <m_ostring> category </m_ostring>\n</root>')
        serialized=assembly.E.fromstring(b'<root><m_ostring> category </m_ostring></root>')
        trimmed=assembly.E.fromstring(b'<root><m_ostring>category</m_ostring></root>')
        self.assertEqual(assembly._normal_xml(first),assembly._normal_xml(serialized))
        self.assertNotEqual(assembly._normal_xml(first),assembly._normal_xml(trimmed))


if __name__ == "__main__":
    unittest.main()
