"""Sequence complete one-slide PPTX sources using official think-cell generation.

Manifest: {"schema":"tc.slide-sequence.v1", "slides":[{"id":"slide-a",
"path":"source-a.pptx", "sha256":"64 hex characters"}, ...]}.
Optional per-slide `edits` apply guarded ordinary PowerPoint object changes.
Native-owned, tagged and linked objects remain protected. Placeholder text
edits retain inheritance; inherited placeholders cannot be removed.
Default is read-only preflight. --execute uses task-owned clones, then requires
per-slide shape/text/font/geometry, chart data/grammar and dependency checks
before and after native save/reopen. Inspect every slide visually before use.
"""
from pathlib import Path
from collections import Counter
import argparse
import hashlib
import io
import json
import posixpath
import re
import subprocess
import sys
import zipfile
from lxml import etree as E

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "thinkcell_no_click/implementation"))
from prepare_thinkcell_name import inventory, logical_slides, link_contract, need
from chart_geometry import streams
from chart_semantics import identity_invariants
from multi_chart_update import model_of, owner_semantics
from office_operation_lock import OfficeOperationLock, run_locked_subprocess
from runtime import find_ppttc, powershell, powershell_env
from ordinary_slide_edits import prepare_source, inspect_shapes

REL = "http://schemas.openxmlformats.org/package/2006/relationships"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"p": P, "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest().upper()


def _normal_xml(node):
    # Namespace-prefix and XML attribute order are serialization details.
    text = node.text or ""
    # Ignore indentation between child elements, preserving every leaf/string
    # value exactly, including meaningful leading/trailing label whitespace.
    if len(node) and text.isspace():
        text = ""
    return (node.tag, tuple(sorted(node.attrib.items())), text,
            tuple(_normal_xml(child) for child in node))


def _carrier_signature(raw, include_model=True):
    if not raw.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        return hashlib.sha256(raw).hexdigest()
    result = []
    for key, value in streams(raw).items():
        if key == ("think-cellXML",) and not include_model:
            continue  # Native model semantics and physical shapes are checked independently.
        try:
            root = E.fromstring(value)
        except E.XMLSyntaxError:
            digest = hashlib.sha256(value).hexdigest()
        else:
            digest = hashlib.sha256(repr(_normal_xml(root)).encode()).hexdigest()
        result.append((key, digest))
    return tuple(sorted(result))


def dependencies(package, slide):
    """Keep multiplicity of native carriers and all linked binary resources."""
    queue, visited, signatures = [slide], set(), []
    while queue:
        part = queue.pop()
        if part in visited:
            continue
        visited.add(part)
        need(part in package.namelist(), "Dangling internal relationship: " + part)
        if "/embeddings/" in part or part.startswith(("customXml/", "ppt/media/")):
            # Hidden or indirect carriers may contain labels/annotation/style
            # nodes beyond the chart data model. Preserve their entire native
            # authoring XML as well as every other stream in the closure.
            signatures.append(repr(_carrier_signature(package.read(part), include_model=True)))
        elif part.startswith("ppt/theme/"):
            # Theme values supply inherited colors/fonts absent from shape XML.
            signatures.append("theme:" + repr(_normal_xml(E.fromstring(package.read(part)))))
        elif part.startswith("ppt/charts/"):
            # Native owning models do not cover ordinary PowerPoint charts.
            # Preserve actual chart caches/style parts too, retaining closure
            # multiplicity instead of reducing every chart to a generic token.
            signatures.append("chart:" + repr(_normal_xml(E.fromstring(package.read(part)))))
        elif part.startswith(("ppt/slideMasters/", "ppt/slideLayouts/", "ppt/notesMasters/")):
            # Placeholders may inherit geometry or visible objects from these.
            signatures.append("layout:" + repr(physical_snapshot(package, part)))
        folder, name = posixpath.split(part)
        relpart = posixpath.join(folder, "_rels", name + ".rels")
        if relpart not in package.namelist():
            continue
        for relationship in E.fromstring(package.read(relpart)):
            need(relationship.get("TargetMode") != "External", "External slide dependencies are unsupported")
            if part.startswith("ppt/slideMasters/") and relationship.get("Type", "").endswith("/slideLayout"):
                continue
            target = relationship.get("Target", "")
            resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
            need(not resolved.startswith("../"), "Relationship escapes package")
            queue.append(resolved)
    return Counter(signatures)


def physical_snapshot(package, slide):
    """Read ordered complete shape content, ignoring package-local reference IDs."""
    root = E.fromstring(package.read(slide))
    tree = root.find("p:cSld/p:spTree", NS)
    need(tree is not None, "Slide shape tree missing")
    folder, name = posixpath.split(slide)
    relpart = posixpath.join(folder, "_rels", name + ".rels")
    relationships = {}
    if relpart in package.namelist():
        for relationship in E.fromstring(package.read(relpart)):
            target = relationship.get("Target", "")
            resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
            raw = package.read(resolved)
            kind = relationship.get("Type", "").rsplit("/", 1)[-1]
            if raw.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
                # CFB storage timestamps may change while every stream remains
                # byte-identical. Bind the complete stream payload including
                # think-cellXML, rather than unrelated storage header bytes.
                token = repr(_carrier_signature(raw, include_model=True))
            else:
                try:
                    token = repr(_normal_xml(E.fromstring(raw)))
                except E.XMLSyntaxError:
                    token = hashlib.sha256(raw).hexdigest()
            relationships[relationship.get("Id")] = kind + ":" + token
    # Snapshot traversal first: removing one extension otherwise makes lxml's
    # live iterator skip later sibling creation records.
    for element in list(root.iter()):
        for attribute in list(element.attrib):
            if attribute.startswith("{" + R + "}"):
                need(element.attrib[attribute] in relationships, "Physical reference has no relationship")
                element.attrib[attribute] = relationships[element.attrib[attribute]]
        if element.tag == "{" + P + "}cNvPr":
            element.attrib.pop("id", None)
        # PowerPoint generates document-local creation identifiers, including
        # slide-level extension records outside the shape tree.
        for child in list(element):
            name = E.QName(child)
            if ((name.namespace == "http://schemas.microsoft.com/office/powerpoint/2010/main" and name.localname in {"creationId", "modId"})
                or (name.namespace == "http://schemas.microsoft.com/office/drawing/2014/main" and name.localname == "creationId")):
                element.remove(child)
    if E.QName(root).localname == "notesMaster":
        # Official generation refreshes the date placeholder's cached display.
        # Keep its genuine field instruction, ID, styles and placeholder owner;
        # literal text/date values anywhere else remain exact.
        for shape in root.findall(".//p:sp", NS):
            if not shape.xpath("./p:nvSpPr/p:nvPr/p:ph[@type='dt']", namespaces=NS):
                continue
            for field in shape.findall(".//a:fld", NS):
                if re.fullmatch(r"datetime(?:FigureOut|[1-9]|1[0-3])", field.get("type", "")):
                    for cached in field.findall("a:t", NS):
                        cached.text = "<dynamic datetime display cache>"
    # Preserve backgrounds, transition/build settings and slide-level extension
    # data as well as shapes. p:cSld/spTree has already been normalized above.
    return _normal_xml(root)


def snapshot(path, slide_number=1):
    data = path if isinstance(path, bytes) else Path(path).read_bytes()
    docs, charts, _ = inventory(data)
    identity_invariants(docs)
    selected = [chart for chart in charts if chart["doc"]["slide_number"] == slide_number]
    carriers = []
    for chart in selected:
        link_contract(chart)
        need(chart["exact"] and len(chart["frames"]) == 1, "Chart identity is ambiguous")
        carriers.append((chart["frames"][0]["shape_tag"], model_of(chart), owner_semantics(chart)))
    need(len({item[0] for item in carriers}) == len(carriers), "Duplicate chart carrier tags")
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        slides = logical_slides(package)
        need(1 <= slide_number <= len(slides), "Slide number out of range")
        slide = slides[slide_number - 1]["part"]
        # Physical equality normalizes package-local IDs, but duplicate slide
        # object IDs are never valid. OLE fallback pictures are one carrier's
        # alternate representation and are excluded by the shared inspector.
        inspect_shapes(package.read(slide))
        dimensions = E.fromstring(package.read("ppt/presentation.xml")).find("p:sldSz", NS)
        need(dimensions is not None, "Slide dimensions missing")
        folder, name = posixpath.split(slide)
        relpart = posixpath.join(folder, "_rels", name + ".rels")
        notes = None
        if relpart in package.namelist():
            for relationship in E.fromstring(package.read(relpart)):
                if relationship.get("Type", "").endswith("/notesSlide"):
                    target = relationship.get("Target", "")
                    notespart = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
                    need(notes is None, "Multiple notes slide dependencies")
                    notes = physical_snapshot(package, notespart)
        return {"dimensions": tuple(sorted(dimensions.attrib.items())),
                "physical": physical_snapshot(package, slide),
                "notes": notes,
                "carriers": carriers, "dependencies": dependencies(package, slide),
                "slide_count": len(slides)}


def _prepared(data, edits):
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        slides = logical_slides(package)
        need(len(slides) == 1, "Every source must contain exactly one slide")
        part = slides[0]["part"]
    return prepare_source(data, edits, part)


def preflight(manifest_path, expected_sha256):
    manifest_path = Path(manifest_path).resolve()
    manifest_bytes = manifest_path.read_bytes()
    need(hashlib.sha256(manifest_bytes).hexdigest().upper() == expected_sha256.upper(), "Manifest SHA-256 mismatch")
    manifest = json.loads(manifest_bytes.decode("utf-8-sig"))
    need(set(manifest) == {"schema", "slides"} and manifest["schema"] == "tc.slide-sequence.v1", "Invalid sequence manifest")
    need(isinstance(manifest["slides"], list) and len(manifest["slides"]) >= 2, "Sequence requires at least two slides")
    sources, snapshots, identities = [], [], set()
    for entry in manifest["slides"]:
        need(isinstance(entry, dict) and {"id", "path", "sha256"}.issubset(entry) and set(entry).issubset({"id", "path", "sha256", "edits"}), "Invalid source entry")
        need("edits" not in entry or isinstance(entry["edits"], list), "Edits must be a list")
        need(isinstance(entry["id"], str) and entry["id"].strip() and entry["id"] not in identities, "Slide IDs must be unique")
        need(isinstance(entry["sha256"], str) and re.fullmatch("[0-9A-Fa-f]{64}", entry["sha256"]), "Malformed source hash")
        need(isinstance(entry["path"], str) and entry["path"].strip(), "Missing source path")
        source = (manifest_path.parent / entry["path"]).resolve()
        need(source.is_file() and source.suffix.lower() == ".pptx", "Source must be an existing PPTX")
        data = source.read_bytes()
        need(hashlib.sha256(data).hexdigest().upper() == entry["sha256"].upper(), "Source SHA-256 mismatch: " + entry["id"])
        state = snapshot(data)
        need(state["slide_count"] == 1, "Every source must contain exactly one slide")
        if entry.get("edits"):
            edited = snapshot(_prepared(data, entry["edits"]))
            for key in ("dimensions", "notes", "carriers", "dependencies"):
                need(edited[key] == state[key], "Ordinary edits changed native/dependency state: " + key)
            state = edited
        if snapshots:
            need(state["dimensions"] == snapshots[0]["dimensions"], "Source slide dimensions differ")
        identities.add(entry["id"]); sources.append(source); snapshots.append(state)
    return manifest, sources, snapshots


def verify_output(output, expected):
    for index, wanted in enumerate(expected, 1):
        actual = snapshot(output, index)
        need(actual["slide_count"] == len(expected), "Assembled slide count differs")
        for key in ("dimensions", "physical", "notes", "carriers", "dependencies"):
            need(actual[key] == wanted[key], f"Slide {index} {key} preservation failed")
    return {"slide_count": len(expected), "physical_content_preserved": True,
            "chart_data_and_grammar_preserved": True, "dependency_closures_preserved": True}


def run(manifest_path, expected_sha256, output, report, execute=False, ppttc=None):
    manifest_path, output, report = map(lambda path: Path(path).resolve(), (manifest_path, output, report))
    manifest, sources, expected = preflight(manifest_path, expected_sha256)
    sealed = {manifest_path: expected_sha256.upper(), **{source: entry["sha256"].upper() for source, entry in zip(sources, manifest["slides"])}}
    stage = output.parent / (output.stem + ".assembly-work")
    need(len({output, report, *sealed}) == len(sealed) + 2, "Output/report alias source or manifest")
    need(output.suffix.lower() == ".pptx" and report.suffix.lower() == ".json", "Output must be PPTX; report must be JSON")
    need(not output.exists() and not report.exists() and not stage.exists(), "Use fresh output/report/staging paths")
    need(output.parent.is_dir() and report.parent.is_dir(), "Destination folders must exist")
    need(not output.is_relative_to(HERE.parent) and not report.is_relative_to(HERE.parent), "Outputs must be outside package assets")
    need(not report.is_relative_to(stage), "Report must be outside staging")
    result = {"status": "SEQUENCE_PREFLIGHT_PASS", "writes": False, "slide_count": len(sources),
              "scope": "source-slide sequencing with guarded ordinary object edits" if any(entry.get("edits") for entry in manifest["slides"]) else "complete source-slide sequencing; no object edits", "native_certification": False}
    if not execute:
        return result
    executable = Path(ppttc).resolve() if ppttc else find_ppttc()
    need(executable and executable.is_file(), "Official ppttc.exe unavailable")
    stage.mkdir()
    clones, clone_hashes = [], []
    for index, (source, entry) in enumerate(zip(sources, manifest["slides"]), 1):
        clone = stage / f"source-{index}.pptx"
        data = source.read_bytes()
        need(hashlib.sha256(data).hexdigest().upper() == sealed[source], "Source changed while cloning")
        prepared = _prepared(data, entry["edits"]) if entry.get("edits") else data
        prepared_state = snapshot(prepared)
        for key in ("dimensions", "physical", "notes", "carriers", "dependencies"):
            need(prepared_state[key] == expected[index - 1][key], "Prepared source differs from preflight")
        with clone.open("xb") as stream:
            stream.write(prepared)
        clone_hashes.append(hashlib.sha256(prepared).hexdigest().upper())
        clones.append(clone)
    job = stage / "sequence.ppttc"
    job.write_text(json.dumps([{"template": str(path), "data": []} for path in clones]), encoding="utf-8")
    need(all(sha(path) == digest for path, digest in sealed.items()), "An original input changed before generation")
    generated, reopened = stage / "generated.pptx", stage / "reopened.pptx"
    native_report, render = stage / "native.json", stage / "slide-1.png"
    with OfficeOperationLock("native-slide-sequence"):
        with (stage / "generation.log").open("w") as log:
            process = run_locked_subprocess([str(executable), str(job), "-o", str(generated)],
                operation="assembly-ppttc", timeout_seconds=240, stdout=log, stderr=log,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        need(process.returncode == 0 and generated.is_file(), "Official sequencing failed; inspect staging")
        verify_output(generated, expected)
        with (stage / "native.log").open("w") as log:
            process = run_locked_subprocess([powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File",
                str(HERE / "thinkcell_no_click/implementation/native_verify_scoped.ps1"), "-InputFile", str(generated),
                "-OutputFile", str(reopened), "-ReportFile", str(native_report), "-RenderFile", str(render)],
                operation="assembly-native-verify", timeout_seconds=240, stdout=log, stderr=log,
                env=powershell_env(), creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        need(process.returncode == 0 and native_report.is_file() and reopened.is_file() and render.is_file(), "Native reopen/render failed; inspect staging")
    native = json.loads(native_report.read_text(encoding="utf-8-sig"))
    need(all(native.get(key) for key in ("native_reopen_pass", "source_unchanged", "other_presentations_unchanged")), "Native scope gates failed")
    gates = verify_output(reopened, expected)
    need(all(sha(path) == digest for path, digest in sealed.items()), "An original input changed; output withheld")
    need(all(sha(path) == digest for path, digest in zip(clones, clone_hashes)), "Task-owned source clone changed")
    with output.open("xb") as stream:
        stream.write(reopened.read_bytes())
    result.update(status="NATIVE_SEQUENCE_VERIFIED_VISUAL_REVIEW_REQUIRED", writes=True,
                  source_unchanged=True, native_reopen_pass=True, output=str(output), output_sha256=sha(output),
                  preservation_gates=gates, staging_directory=str(stage), first_slide_preview=str(render),
                  visual_review="Review every slide; the automated preview covers slide 1 only")
    with report.open("x", encoding="utf-8") as stream:
        json.dump(result, stream, indent=2)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--expected-sha256", required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--ppttc", type=Path)
    parser.add_argument("--execute", action="store_true")
    args = parser.parse_args()
    print(json.dumps(run(args.manifest, args.expected_sha256, args.output, args.report, args.execute, args.ppttc), indent=2))


if __name__ == "__main__":
    main()
