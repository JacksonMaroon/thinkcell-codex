"""Portable builder for the authenticated official scatter trendline seed.

The builder performs only offline package work. It resolves the installed
think-cell naming helpers by package layout (or an explicit directory),
selects the target by slide and physical chart tag, and reports semantic
identity rather than relying on regenerated numeric IDs. Office regeneration
and native save/reopen remain separate verification steps.
"""
from __future__ import annotations

import argparse
import importlib.util
import io
import json
import subprocess
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
AUTOMATION_NAME = "TC_AUTO_OFFICIAL_TRENDLINE_01"


def resolve_implementation(explicit: Path | None = None) -> Path:
    """Find the installed helper pair without depending on the lane CWD."""
    candidates: list[Path] = []
    here = Path(__file__).resolve()
    # A copied installed script resolves helpers from its own package tree.
    # Keep these relative probes ahead of the explicit test override.
    for ancestor in (here.parent, *here.parents):
        candidates.extend([
            ancestor,
            ancestor / "implementation",
            ancestor / "scripts/thinkcell_no_click/implementation",
            ancestor / "skills/thinkcell-edit/scripts/thinkcell_no_click/implementation",
        ])
    if explicit is not None:
        candidates.append(explicit.expanduser().resolve())
    for candidate in candidates:
        if (candidate / "prepare_thinkcell_name.py").is_file() and (candidate / "replace_ole_stream.ps1").is_file():
            return candidate
    searched = "; ".join(str(p) for p in candidates)
    raise RuntimeError("Installed think-cell implementation helpers not found; searched: " + searched)


def load_naming(implementation: Path):
    module_path = implementation / "prepare_thinkcell_name.py"
    spec = importlib.util.spec_from_file_location("installed_thinkcell_naming", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load naming helper: {module_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def model_value(ids: dict[str, object], node: object | None) -> str:
    if node is None:
        return ""
    ref = node.get("idref")
    source = ids.get(ref) if ref else None
    if source is None:
        return ""
    value = source.find("m_varval")
    return "" if value is None else (value.text or value.get("val") or "")


def semantic_identity(data: bytes, naming, slide_number: int, shape_tag: str) -> dict[str, object]:
    """Return the target's stable semantic contract and physical closures."""
    _, candidates, _ = naming.inventory(data)
    chosen = naming.choose(candidates, slide_number=slide_number, shape_tag=shape_tag)
    ids = chosen["doc"]["ids"]
    table = chosen["table"]
    series_ids = [e.get("idref") for box in table.findall("m_cscatdseries") for e in box]
    series = {sid: model_value(ids, ids[sid].find("m_varsrc")) for sid in series_ids if sid in ids}
    partitions = []
    partition_refs = [e.get("idref") for box in table.findall("m_cscatpartition") for e in box]
    for partition_ref in partition_refs:
        partition = ids[partition_ref]
        sid = partition.find("m_scatdseries").get("idref")
        style = partition.find("m_linestyle")
        line = partition.find("m_linefPartition")
        physical = partition.find("m_pptfreeform")
        partitions.append({
            "series": series.get(sid, ""),
            "partition_type": partition.find("m_epartitiontype").get("val"),
            "trendline_type": partition.find("m_etrendlinetype").get("val"),
            "visible": style.find("m_bVisible").get("val") if style is not None else "",
            "polyline_present": partition.find("m_pptpolyline") is not None,
            "physical_freeform": physical.find("m_bstrShapeName").text if physical is not None else "",
            "line_present": line is not None,
        })
    return {
        "slide_number": slide_number,
        "shape_tag": shape_tag,
        "model_type": chosen["owner"].tag,
        "logical_chart_id": chosen["owner"].get("id"),
        "logical_table_id": table.get("id"),
        "active_document_part": chosen["doc"]["part"],
        "series_labels": sorted(v for v in series.values() if v),
        "partitions": sorted(partitions, key=lambda p: p["series"]),
    }


def compare_semantics(output: dict[str, object], proof: dict[str, object]) -> bool:
    for key in ("slide_number", "shape_tag", "model_type", "series_labels"):
        if output[key] != proof[key]:
            raise RuntimeError(f"semantic identity mismatch for {key}: {output[key]!r} vs {proof[key]!r}")
    out_parts = {p["series"]: p for p in output["partitions"]}
    proof_parts = {p["series"]: p for p in proof["partitions"]}
    if set(out_parts) != set(proof_parts):
        raise RuntimeError("semantic partition series differ from native-proven candidate")
    for series in out_parts:
        for key in ("partition_type", "trendline_type", "visible", "polyline_present", "line_present"):
            if out_parts[series][key] != proof_parts[series][key]:
                raise RuntimeError(f"semantic partition mismatch for {series}/{key}")
    return True


def content_type_normalization(payload: bytes, suffix: str) -> tuple[bytes, str]:
    template = b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
    presentation = b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
    if presentation in payload:
        return payload, "already presentation.main+xml"
    if template in payload:
        return payload.replace(template, presentation), f"{suffix} template.main+xml -> presentation.main+xml"
    raise RuntimeError("[Content_Types].xml has neither the PPTX nor POTX main content type")


def guard_paths(source: Path, output: Path, report: Path, proof_native: Path | None) -> None:
    inputs = [source] + ([proof_native] if proof_native is not None else [])
    all_paths = inputs + [output, report]
    resolved = [p.resolve() for p in all_paths]
    if len(set(resolved)) != len(resolved):
        raise RuntimeError("source, proof, output, and report paths must be pairwise distinct")
    if not source.is_file():
        raise RuntimeError(f"source does not exist: {source}")
    if proof_native is not None and not proof_native.is_file():
        raise RuntimeError(f"proof native output does not exist: {proof_native}")
    for destination in (output, report):
        if destination.exists():
            raise RuntimeError(f"refusing to overwrite existing output/report: {destination}")


def build(source: Path, output: Path, implementation: Path, expected_source_sha256: str,
          slide_number: int, shape_tag: str, automation_name: str,
          expected_series_names: list[str] | None = None,
          proof_native: Path | None = None) -> dict[str, object]:
    report_path = output.with_suffix(".portable.json")
    guard_paths(source, output, report_path, proof_native)
    naming = load_naming(implementation)
    raw = source.read_bytes()
    source_sha = naming.sha(raw)
    if source_sha != expected_source_sha256.upper():
        raise RuntimeError(f"source SHA mismatch: {source_sha} vs {expected_source_sha256.upper()}")
    _, candidates, _ = naming.inventory(raw)
    selected = naming.choose(candidates, slide_number=slide_number, shape_tag=shape_tag)
    source_identity = semantic_identity(raw, naming, slide_number, shape_tag)
    target_series_names = sorted(expected_series_names or source_identity["series_labels"])
    if source_identity["series_labels"] != target_series_names:
        raise RuntimeError(f"expected series names do not match source donor: {target_series_names!r} vs {source_identity['series_labels']!r}")
    model, changes = naming.rewrite(selected, automation_name, False)
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="trendline_portable_", dir=str(output.parent)) as td:
        temp = Path(td)
        carrier, model_file = temp / "carrier.bin", temp / "model.xml"
        carrier.write_bytes(selected["doc"]["ole"])
        model_file.write_bytes(model)
        helper = implementation / "replace_ole_stream.ps1"
        result = subprocess.run(
            [naming.powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(helper),
             "-StoragePath", str(carrier), "-StreamBytesPath", str(model_file)],
            capture_output=True, text=True, check=False, env=naming.powershell_env(), timeout=60,
        )
        if result.returncode:
            raise RuntimeError(result.stderr or "CFB model replacement failed")
        selected_ole = carrier.read_bytes()
        package_repair = None
        with zipfile.ZipFile(io.BytesIO(raw)) as src, zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as dst:
            dst.comment = src.comment
            for entry in src.infolist():
                payload = src.read(entry.filename)
                if entry.filename == selected["doc"]["part"]:
                    payload = selected_ole
                elif entry.filename == "[Content_Types].xml":
                    payload, package_repair = content_type_normalization(payload, source.suffix.lower())
                dst.writestr(entry, payload)
        if package_repair is None:
            raise RuntimeError("package is missing [Content_Types].xml")
    output_identity = semantic_identity(output.read_bytes(), naming, slide_number, shape_tag)
    if output_identity["series_labels"] != target_series_names:
        raise RuntimeError(f"output semantic series identity is wrong: {output_identity['series_labels']!r} vs {target_series_names!r}")
    if source.read_bytes() != raw:
        raise RuntimeError("source changed during portable preparation")
    proof_identity = None
    semantic_match = None
    if proof_native is not None:
        proof_identity = semantic_identity(proof_native.read_bytes(), naming, slide_number, shape_tag)
        semantic_match = compare_semantics(output_identity, proof_identity)
    report = {
        "status": "PORTABLE_AUTHENTIC_TRENDLINE_IMPLEMENTATION_PREPARED",
        "source": str(source.resolve()), "source_sha256": source_sha,
        "output": str(output.resolve()), "output_sha256": naming.sha(output.read_bytes()),
        "implementation_dir": str(implementation.resolve()),
        "selector": output_identity,
        "source_selector": source_identity,
        "expected_series_names": target_series_names,
        "proof_native": str(proof_native.resolve()) if proof_native else None,
        "proof_native_selector": proof_identity,
        "semantic_identity_matches_native_proven_candidate": semantic_match,
        "automation_name": automation_name, "name_changes": changes,
        "native_feature_preserved": ["CScatterChartPartition", "m_linefPartition", "m_pptpolyline", "m_pptfreeform", "m_scatdseries"],
        "package_repair": package_repair,
        "native_required": ["official ppttc regeneration", "native save/reopen/render", "changed-data regression readback"],
    }
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    return report


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True,
                        help="Explicit authentic donor path")
    parser.add_argument("--output", type=Path, required=True,
                        help="Fresh output path; never write into the installed package")
    parser.add_argument("--expected-source-sha256", required=True)
    parser.add_argument("--implementation-dir", type=Path)
    parser.add_argument("--slide-number", type=int, default=1)
    parser.add_argument("--shape-tag", required=True,
                        help="Explicit physical chart tag")
    parser.add_argument("--automation-name", default=AUTOMATION_NAME)
    parser.add_argument("--expected-series-name", action="append", dest="expected_series_names",
                        help="Optional expected donor series name; repeat for each series")
    parser.add_argument("--proof-native", type=Path,
                        help="Optional native-proven output used for semantic identity comparison")
    args = parser.parse_args()
    if args.slide_number < 1:
        parser.error("--slide-number must be positive")
    implementation = resolve_implementation(args.implementation_dir)
    report = build(args.source.resolve(), args.output.resolve(), implementation,
                   args.expected_source_sha256, args.slide_number, args.shape_tag, args.automation_name,
                   args.expected_series_names,
                   args.proof_native.resolve() if args.proof_native else None)
    print(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()
