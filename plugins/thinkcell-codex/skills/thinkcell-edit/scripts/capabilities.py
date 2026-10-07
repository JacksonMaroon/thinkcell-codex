"""Read-only capability discovery and explicit dispatch to packaged adapters.

Availability proves a packaged route exists, never native runtime certification.
The registry retains operation-specific scope and required release gates.
"""
from pathlib import Path
import argparse
import json
import re
import subprocess
import sys

HERE = Path(__file__).resolve().parent
REGISTRY = HERE.parent / 'references' / 'capabilities.json'
STATES = {'bounded_native', 'prepare_only', 'preservation_only', 'coordination',
          'historical_only', 'research'}


def packaged_path(relative, root=HERE):
    if not isinstance(relative, str) or not relative or '\\' in relative:
        raise ValueError('Route must be a portable relative script path')
    path = Path(relative)
    if path.is_absolute() or ':' in relative or '..' in path.parts:
        raise ValueError('Route must stay inside packaged scripts')
    resolved = (root / path).resolve()
    if not resolved.is_relative_to(root.resolve()):
        raise ValueError('Route escapes packaged scripts')
    if not resolved.is_file():
        raise ValueError('Packaged route is missing: ' + relative)
    return resolved


def load_registry(path=REGISTRY):
    data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    if set(data) != {'schema', 'native_evidence_policy', 'capabilities'} or data['schema'] != 'tc.capabilities.v1':
        raise ValueError('Unknown capability registry schema')
    entries = data['capabilities']
    if not isinstance(entries, list) or not entries:
        raise ValueError('Registry must contain capabilities')
    seen = set()
    for entry in entries:
        if set(entry) != {'id', 'title', 'status', 'scope', 'route', 'validation', 'limitations'}:
            raise ValueError('Invalid capability fields')
        identity = entry['id']
        if not isinstance(identity, str) or not re.fullmatch('[a-z][a-z0-9-]*', identity) or identity in seen:
            raise ValueError('Capability IDs must be unique portable slugs')
        seen.add(identity)
        if entry['status'] not in STATES:
            raise ValueError('Unknown capability status')
        for field in ('title', 'scope', 'limitations'):
            if not isinstance(entry[field], str) or not entry[field].strip():
                raise ValueError('Capability ' + field + ' is required')
        validation = entry['validation']
        if not isinstance(validation, list) or not validation or not all(isinstance(g, str) and g.strip() for g in validation):
            raise ValueError('Every capability requires explicit validation')
        route = entry['route']
        if route is not None:
            if not isinstance(route, dict) or set(route) != {'script', 'arguments', 'stage'}:
                raise ValueError('Invalid packaged route')
            packaged_path(route['script'])
            if not isinstance(route['arguments'], list) or not all(isinstance(a, str) for a in route['arguments']):
                raise ValueError('Route arguments must be a string array')
            if route['stage'] not in {'native_pipeline', 'preparation', 'verification', 'inspection'}:
                raise ValueError('Unknown route stage')
        elif entry['status'] not in {'historical_only', 'research', 'coordination'}:
            raise ValueError('Implemented capability requires a packaged route')
    return data


def discover(identity=None):
    data = load_registry()
    entries = data['capabilities']
    if identity is not None:
        entries = [c for c in entries if c['id'] == identity]
        if not entries:
            raise ValueError('Unknown capability: ' + identity)
    return {'schema': data['schema'], 'native_evidence_policy': data['native_evidence_policy'],
            'capabilities': entries, 'packaged_routes_checked': True,
            'native_certification_renewed': False}


def dispatch(identity, arguments):
    entry = discover(identity)['capabilities'][0]
    route = entry['route']
    if route is None:
        raise ValueError(entry['title'] + ': no packaged execution route. ' + entry['limitations'])
    argv = list(arguments)
    if argv and argv[0] == '--':
        argv.pop(0)
    # Child scripts retain their own selectors, SHA guards, output freshness,
    # execute flags, and shared Office locks. Never inject --execute here.
    return subprocess.run([sys.executable, str(packaged_path(route['script'])),
                           *route['arguments'], *argv], check=False).returncode


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    listing = sub.add_parser('list', help='List or inspect scoped capabilities without opening Office')
    listing.add_argument('--id')
    sub.add_parser('check', help='Validate registry fields and packaged route paths')
    route = sub.add_parser('run', help='Explicitly run one packaged adapter with its normal guards')
    route.add_argument('id')
    route.add_argument('arguments', nargs=argparse.REMAINDER)
    args = parser.parse_args()
    try:
        if args.command == 'run':
            return dispatch(args.id, args.arguments)
        result = discover(getattr(args, 'id', None))
        if args.command == 'check':
            result = {'status': 'CAPABILITY_REGISTRY_PASS', 'capabilities': len(result['capabilities']),
                      'runnable_routes': sum(c['route'] is not None for c in result['capabilities']),
                      'native_certification_renewed': False}
        print(json.dumps(result, indent=2, ensure_ascii=False, allow_nan=False))
        return 0
    except (ValueError, OSError, json.JSONDecodeError) as exc:
        print(json.dumps({'status': 'REJECTED', 'error': str(exc)}), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
