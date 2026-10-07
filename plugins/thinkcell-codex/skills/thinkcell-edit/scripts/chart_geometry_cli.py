"""Plan, prepare, or verify a bounded selected-chart geometry change.

Preparation is intermediate. Official regeneration, native reopen and this
readback must succeed before delivery; this command never certifies a cache-only edit.
"""
from pathlib import Path
import argparse
import json
import sys
import chart_geometry as geometry


def fresh_output(output, inputs):
    path = Path(output).expanduser().resolve()
    if path in {Path(p).expanduser().resolve() for p in inputs} or path.exists():
        raise ValueError('Output must be fresh and distinct from every input')
    return path


def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    plan = sub.add_parser('make-plan')
    plan.add_argument('--input', required=True, type=Path)
    plan.add_argument('--expected-sha256', required=True)
    plan.add_argument('--selection-json', required=True, type=Path)
    plan.add_argument('--frame-json', required=True, type=Path)
    mode = plan.add_mutually_exclusive_group()
    mode.add_argument('--translation-json', type=Path)
    mode.add_argument('--plot-bounds-json', type=Path)
    plan.add_argument('--plan-out', required=True, type=Path)
    prepare = sub.add_parser('prepare')
    prepare.add_argument('--input', required=True, type=Path)
    prepare.add_argument('--plan', required=True, type=Path)
    prepare.add_argument('--output', required=True, type=Path)
    verify = sub.add_parser('verify')
    verify.add_argument('--input', required=True, type=Path)
    verify.add_argument('--plan', required=True, type=Path)
    args = parser.parse_args()
    try:
        if args.command == 'make-plan':
            inputs = [args.input, args.selection_json, args.frame_json]
            inputs += [p for p in (args.translation_json, args.plot_bounds_json) if p]
            out = fresh_output(args.plan_out, inputs)
            if geometry.sha(args.input.read_bytes()) != args.expected_sha256.upper():
                raise ValueError('Source SHA-256 mismatch')
            result = geometry.make_plan(args.input, read_json(args.frame_json),
                selection=read_json(args.selection_json),
                translation=read_json(args.translation_json) if args.translation_json else None,
                plot_bounds=read_json(args.plot_bounds_json) if args.plot_bounds_json else None)
            with out.open('x', encoding='utf-8') as handle:
                json.dump(result, handle, indent=2, allow_nan=False)
            result = {'status': 'GEOMETRY_PLAN_READY', 'plan': str(out), 'native_validation_required': True}
        elif args.command == 'prepare':
            output = fresh_output(args.output, [args.input, args.plan])
            result = geometry.prepare(args.input, output, read_json(args.plan))
        else:
            result = geometry.verify(args.input, read_json(args.plan))
            result['native_reopen_proven_by_this_command'] = False
        print(json.dumps(result, indent=2, allow_nan=False))
        return 0
    except (ValueError, OSError, json.JSONDecodeError) as exc:
        print(json.dumps({'status': 'REJECTED', 'error': str(exc)}), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
