"""Stage distinct existing-link rebinds, publishing only after all checks pass.

Each request uses portable_rebind's guarded native file-moniker or opaque
equal-UTF16-byte-length contract. No Office call or native certification claim.
"""
from __future__ import annotations
import argparse
import contextlib
import io
import json
from pathlib import Path
import tempfile
import zipfile

from portable_rebind import main as rebind_main, sha256


REQUIRED = {'source_workbook', 'source_workbook_sha256', 'target_workbook',
            'target_workbook_sha256'}
OPTIONAL = {'guid', 'link_id', 'range_name', 'expected_range'}


def load_requests(path: Path) -> list[dict[str, str]]:
    manifest = json.loads(path.read_text(encoding='utf-8'))
    if not isinstance(manifest, dict) or set(manifest) != {'schema', 'links'} or manifest['schema'] != 'thinkcell-batch-rebind-v1':
        raise ValueError('manifest must contain schema thinkcell-batch-rebind-v1 and links')
    links = manifest['links']
    if not isinstance(links, list) or not links:
        raise ValueError('links must be a nonempty list')
    for link in links:
        if not isinstance(link, dict) or not REQUIRED <= set(link) or set(link) - REQUIRED - OPTIONAL:
            raise ValueError('each link must contain the required workbook paths/hashes and only known guards')
        if not all(isinstance(value, str) and value.strip() for value in link.values()):
            raise ValueError('link request values must be nonempty strings')
        if not any(key in link for key in ('guid', 'link_id', 'range_name')):
            raise ValueError('each batch link requires an exact identity selector')
    return links


def run(input_path: Path, output_path: Path, manifest_path: Path,
        report_path: Path, expected_sha256: str) -> dict:
    input_path, output_path, manifest_path, report_path = [
        path.resolve() for path in (input_path, output_path, manifest_path, report_path)]
    links = load_requests(manifest_path)
    sources = {input_path, manifest_path}
    for link in links:
        sources.update(Path(link[key]).resolve() for key in ('source_workbook', 'target_workbook'))
    if output_path == report_path or output_path in sources or report_path in sources:
        raise ValueError('output/report must be distinct from all sealed inputs and each other')
    if output_path.exists() or report_path.exists():
        raise ValueError('fresh output and report paths required')
    sealed = {path: sha256(path) for path in sources}
    if sealed[input_path] != expected_sha256.upper():
        raise ValueError('input SHA-256 mismatch')
    reports = []
    selected_parts = set()
    with tempfile.TemporaryDirectory(prefix='thinkcell-rebind-') as folder:
        stage = Path(folder)
        current = input_path
        for index, link in enumerate(links):
            candidate = stage / f'{index}.pptx'
            evidence = stage / f'{index}.json'
            argv = ['--input-presentation', str(current), '--output-presentation', str(candidate),
                    '--input-sha256', sha256(current), '--report', str(evidence)]
            for key, value in link.items():
                argv.extend(['--' + key.replace('_', '-'), value])
            with contextlib.redirect_stdout(io.StringIO()):
                rebind_main(argv)
            report = json.loads(evidence.read_text(encoding='utf-8'))
            if report['selected_part'] in selected_parts:
                raise ValueError('batch requests must select distinct linked carriers')
            selected_parts.add(report['selected_part'])
            # Intermediate temp paths are not durable artifacts.
            reports.append({key: report[key] for key in (
                'selected_part', 'source_workbook', 'source_workbook_sha256',
                'target_workbook', 'target_workbook_sha256', 'old_moniker_workbook_path',
                'new_moniker_workbook_path', 'identity_before', 'identity_after', 'guards')})
            current = candidate
        with zipfile.ZipFile(input_path) as before, zipfile.ZipFile(current) as after:
            if sorted(before.namelist()) != sorted(after.namelist()):
                raise RuntimeError('batch changed package membership')
            changed = sorted(name for name in before.namelist() if before.read(name) != after.read(name))
        if changed != sorted(selected_parts):
            raise RuntimeError('batch changed unexpected package parts')
        if any(sha256(path) != digest for path, digest in sealed.items()):
            raise RuntimeError('sealed inputs changed during preparation')
        result = {'schema': 'thinkcell-batch-rebind-v1', 'status': 'PREPARED_OFFLINE_BATCH_REBIND',
                  'input_presentation': str(input_path), 'output_presentation': str(output_path),
                  'input_sha256': sealed[input_path], 'output_sha256': sha256(current),
                  'manifest_sha256': sealed[manifest_path], 'changed_parts': changed,
                  'links': reports, 'sealed_inputs_unchanged': True,
                  'native_certification': 'NOT_PERFORMED_OFFLINE_ADAPTER_ONLY'}
        output_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        published = []
        try:
            with output_path.open('xb') as handle:
                published.append(output_path)
                handle.write(current.read_bytes())
            with report_path.open('x', encoding='utf-8') as handle:
                published.append(report_path)
                handle.write(json.dumps(result, indent=2))
        except BaseException:
            for path in published:
                path.unlink()
            raise
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description='Prepare a guarded batch of distinct existing Excel link rebinds.')
    parser.add_argument('--input-presentation', type=Path, required=True)
    parser.add_argument('--output-presentation', type=Path, required=True)
    parser.add_argument('--manifest', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    parser.add_argument('--input-sha256', required=True)
    args = parser.parse_args()
    result = run(args.input_presentation, args.output_presentation, args.manifest,
                 args.report, args.input_sha256)
    print(json.dumps(result, indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
