"""Plan a chart frame without changing slides or inventing chart/data semantics."""
import argparse
import hashlib
import json
import math
import re
from pathlib import Path


def number(value, name, positive=False):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{name} must be a finite number")
    if not math.isfinite(value) or value < 0 or (positive and value == 0):
        raise ValueError(f"{name} must be {'positive' if positive else 'nonnegative'} and finite")
    return float(value)


def rect(value):
    # Native horizontal/vertical rules can legitimately have a zero extent.
    return tuple(number(value[k], k)
                 for k in ('left', 'top', 'width', 'height'))


def plan(context, brief):
    """Return a review-required plan. All coordinates are PowerPoint points.

    Context: source_sha256, slide_id, width, height, shapes, optional reserved.
    Rectangles use left/top/width/height. Shapes additionally require id.
    Brief: requested_chart_type OR inferred_chart_type, rationale; optional
    frame (partial rectangle), side, margin, gap, min_width, min_height.
    Chart selection and rationale come from the agent inspecting real evidence.
    """
    allowed = {'requested_chart_type', 'inferred_chart_type', 'rationale',
               'frame', 'side', 'margin', 'gap', 'min_width', 'min_height'}
    if set(brief) - allowed:
        raise ValueError('Unknown brief fields: ' + ', '.join(sorted(set(brief) - allowed)))
    source = context['source_sha256']
    if not isinstance(source, str) or not re.fullmatch(r'[0-9a-fA-F]{64}', source):
        raise ValueError('source_sha256 must bind the plan to the source deck')
    if (isinstance(context['slide_id'], bool)
            or not isinstance(context['slide_id'], (int, str))
            or not str(context['slide_id']).strip()):
        raise ValueError('slide_id is required')
    width, height = (number(context[k], k, True) for k in ('width', 'height'))
    explicit_type = brief.get('requested_chart_type')
    if 'requested_chart_type' in brief and (not isinstance(explicit_type, str) or not explicit_type.strip()):
        raise ValueError('requested_chart_type cannot be empty')
    chart_type = explicit_type or brief.get('inferred_chart_type')
    if not isinstance(chart_type, str) or not chart_type.strip():
        raise ValueError('Agent must supply a chart type from the request or slide/data context')
    if not isinstance(brief.get('rationale'), str) or not brief['rationale'].strip():
        raise ValueError('Agent must explain chart choice using available evidence')
    frame = brief.get('frame', {})
    if not isinstance(frame, dict) or set(frame) - {'left', 'top', 'width', 'height'}:
        raise ValueError('frame accepts only left, top, width, height')
    frame = {k: number(v, k, k in ('width', 'height')) for k, v in frame.items()}
    margin = number(brief.get('margin', 12), 'margin')
    gap = number(brief.get('gap', 6), 'gap')
    min_w = number(brief.get('min_width', frame.get('width', 72)), 'min_width', True)
    min_h = number(brief.get('min_height', frame.get('height', 54)), 'min_height', True)
    left, top, right, bottom = margin, margin, width - margin, height - margin
    if 'margin' not in brief:
        # Defaults guide missing choices; explicit edges/sizes may use the canvas.
        def axis_bounds(extent, position_key, size_key):
            lo, hi = margin, extent - margin
            if position_key in frame:
                lo = min(lo, frame[position_key])
                if size_key in frame:
                    hi = min(extent, max(hi, frame[position_key] + frame[size_key]))
            elif size_key in frame:
                padding = max(0, min(margin, (extent - frame[size_key]) / 2))
                lo, hi = padding, extent - padding
            return lo, hi
        left, right = axis_bounds(width, 'left', 'width')
        top, bottom = axis_bounds(height, 'top', 'height')
    side = brief.get('side')
    if side not in (None, 'left', 'right', 'top', 'bottom'):
        raise ValueError('side must be left, right, top, or bottom')
    if side == 'left': right = min(right, width / 2)
    if side == 'right': left = max(left, width / 2)
    if side == 'top': bottom = min(bottom, height / 2)
    if side == 'bottom': top = max(top, height / 2)
    obstacles, shape_ids = [], []
    for shape in context.get('all_shapes', context['shapes']):
        if ('id' not in shape or isinstance(shape['id'], bool)
                or not isinstance(shape['id'], (int, str))
                or not str(shape['id']).strip() or shape['id'] in shape_ids):
            raise ValueError('Every existing shape requires a unique id')
        shape_ids.append(shape['id'])
    if any(shape.get('id') not in shape_ids for shape in context['shapes']):
        raise ValueError('Obstacle IDs must identify existing shapes')
    for shape in context['shapes'] + context.get('reserved', []):
        x, y, w, h = rect(shape)
        if x + w > width or y + h > height:
            raise ValueError('Context rectangles must be within the slide')
        obstacles.append((x - gap, y - gap, x + w + gap, y + h + gap))
    edges = sorted({left, right} | {max(left, min(right, x))
                    for box in obstacles for x in (box[0], box[2])})
    xs = [frame['left']] if 'left' in frame else edges
    if 'width' in frame and 'left' not in frame:
        xs = sorted(set(xs + [x - frame['width'] for x in edges]))
    candidates = []
    for x in xs:
        ends = [x + frame['width']] if 'width' in frame else edges
        for end in ends:
            w = end - x
            if x < left or end > right or w < min_w:
                continue
            # Union blocked vertical intervals for this horizontal span.
            blocked = sorted((max(top, b[1]), min(bottom, b[3])) for b in obstacles
                             if x < b[2] and end > b[0] and b[1] < bottom and b[3] > top)
            free, cursor = [], top
            for lo, hi in blocked:
                if lo > cursor: free.append((cursor, lo))
                cursor = max(cursor, hi)
            if cursor < bottom: free.append((cursor, bottom))
            for lo, hi in free:
                y = frame.get('top', lo)
                h = frame.get('height', hi - y)
                if y < lo or y + h > hi or h < min_h:
                    continue
                candidates.append((w * h, -abs(w / h - 1.6), -y, -x, (x, y, w, h)))
    if not candidates:
        raise ValueError('No collision-free frame satisfies guidance and minimum size; revise the layout or guidance')
    chosen = dict(zip(('left', 'top', 'width', 'height'), max(candidates)[-1]))
    context_hash = hashlib.sha256(json.dumps(context, sort_keys=True, allow_nan=False,
                                             separators=(',', ':')).encode()).hexdigest()
    return {'status': 'PLACEMENT_PLAN_REVIEW_REQUIRED', 'source_sha256': source.lower(),
            'slide_id': context['slide_id'], 'context_sha256': context_hash,
            'chart_type': chart_type, 'rationale': brief['rationale'], 'frame': chosen,
            'frame_role': 'outer_including_labels',
            'provenance': {'chart_type': 'explicit' if explicit_type else 'inferred',
                           **{k: 'explicit' if k in frame else 'available_space' for k in chosen},
                           'side': 'explicit' if side else 'unconstrained'},
            'constraints': {'side': side, 'margin': margin, 'gap': gap,
                            'min_width': min_w, 'min_height': min_h},
            'preserve_shape_ids': shape_ids, 'requires_visual_review': True,
            'requires_native_geometry_readback': True,
            'native_execution': False}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--context', type=Path, required=True)
    parser.add_argument('--brief', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    result = plan(json.loads(args.context.read_text(encoding='utf-8-sig')),
                  json.loads(args.brief.read_text(encoding='utf-8-sig')))
    with args.output.open('x', encoding='utf-8') as stream:
        json.dump(result, stream, indent=2, allow_nan=False)
        stream.write('\n')


if __name__ == '__main__':
    main()
