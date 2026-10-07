"""Validate release identity, integrity hashes and local documentation links."""
from pathlib import Path
import argparse
import hashlib
import json
import re
import sys
from urllib.parse import unquote

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[4]


def validate(root=ROOT):
    root = Path(root).resolve()
    manifest = json.loads((root / 'CONTENTS.json').read_text(encoding='utf-8-sig'))
    version = manifest['version']
    if not isinstance(version, str) or not version.strip() or not isinstance(manifest['files'], dict):
        raise ValueError('Release version and file mapping are required')
    errors = []
    for relative, expected in manifest['files'].items():
        path = (root / relative).resolve()
        if Path(relative).is_absolute() or ':' in relative or not path.is_relative_to(root):
            errors.append('Unsafe integrity path: ' + relative)
            continue
        if not path.is_file():
            errors.append('Missing integrity file: ' + relative)
            continue
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        if digest != expected:
            errors.append('Integrity mismatch: ' + relative)
    plugin = root / 'plugins/thinkcell-codex'
    for path in (plugin / 'plugin.json', plugin / '.codex-plugin/plugin.json'):
        metadata = json.loads(path.read_text(encoding='utf-8-sig'))
        if metadata.get('name') != 'thinkcell-codex' or metadata.get('version') != version:
            errors.append('Plugin identity/version mismatch: ' + str(path.relative_to(root)))
    citation = (plugin / 'CITATION.cff').read_text(encoding='utf-8-sig')
    if not re.search(r'^version: ' + re.escape(version) + r'\s*$', citation, re.MULTILINE):
        errors.append('Citation version differs from release')
    checked = 0
    for relative in manifest['files']:
        if not relative.endswith('.md'):
            continue
        document = root / relative
        if not document.is_file():
            continue
        for match in re.finditer(r'\[[^\]]*\]\(([^)]+)\)', document.read_text(encoding='utf-8-sig')):
            target = match.group(1).strip().strip('<>')
            if re.match(r'^[A-Za-z][A-Za-z0-9+.-]*:', target) or target.startswith('#'):
                continue
            target = unquote(target.split('#', 1)[0])
            if not target:
                continue
            checked += 1
            if not (document.parent / target).is_file():
                errors.append('Missing documentation target in ' + relative + ': ' + target)
    if errors:
        raise ValueError('\n'.join(errors))
    return {'status': 'RELEASE_INTEGRITY_PASS', 'version': version,
            'integrity_hashes': len(manifest['files']), 'local_documentation_links': checked}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, default=ROOT)
    args = parser.parse_args()
    try:
        print(json.dumps(validate(args.root), indent=2))
        return 0
    except (OSError, ValueError, KeyError, TypeError) as exc:
        print(json.dumps({'status': 'RELEASE_INTEGRITY_FAIL', 'error': str(exc)}), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
