"""Verify this experiment's exact fixture contract, not arbitrary chart support."""
import json
import math
from pathlib import Path

base = Path(__file__).resolve().parent
evidence = {x['file']: x for x in json.loads((base / 'evidence.json').read_text())}
expected = {'02-controlled-percent.pptx': [25, 75, 100],
            '03-numerator-change.pptx': [25, 75, 200],
            '04-denominator-change.pptx': [125, 75, 200],
            '05-reopened-saved.pptx': [125, 75, 200]}
checks = []
for file, first_column in expected.items():
    item = evidence[file]
    models = [m for m in item['models'] if m['labels']]
    assert len(models) == 1
    model = models[0]
    assert model['version'] == '38775'
    assert model['columns'][0]['values'] == first_column
    assert model['chart_consistency'] == ['1']
    assert len(model['labels']) == 9
    assert model['relative_source_text_variable_counts'] == [0] * 9
    assert not model['zero_width_space_in_model']
    assert all(x['msgraph_rendering'] == '1' and x['suffix'] == '%' and
               not x['prefix'] and x['digits'] == '0' for x in model['labels'])
    chart, = item['charts']
    assert chart['field_count'] == 0
    # Native chart cache reverses the model's three series in these fixtures.
    for col_index, col in enumerate(model['columns']):
        total = sum(col['values'])
        for series_index, val in enumerate(reversed(col['values'])):
            cached = float(chart['series'][series_index]['values'][col_index])
            assert math.isclose(cached, val / total * 100, abs_tol=1e-9)
    assert all(s['show_values'] == ['1'] * 3 and s['show_percent'] == ['0'] * 3
               and len(s['number_formats']) == 3
               and all('%' in fmt for fmt in s['number_formats']) for s in chart['series'])
    checks.append({'file': file, 'result': 'pass', 'target_ratio': first_column[2] / sum(first_column)})
reference = [m for m in evidence['02-controlled-percent.pptx']['models'] if m['labels']][0]
for name in expected:
    current = [m for m in evidence[name]['models'] if m['labels']][0]
    assert current['columns'][1:] == reference['columns'][1:]
    assert current['labels'] == reference['labels']
rows = json.loads((base / 'existing-discovery.json').read_text())
assert len(rows) == 9
assert all(r['relative_text_variable'] is None and r['physical_shapes'] == [] for r in rows)
(base / 'verification.json').write_text(json.dumps({'status': 'pass', 'scope': 'six local Mac fixtures; no Windows or JSON regeneration proof', 'checks': checks}, indent=2))
print('PASS: exact native model values, all nine normalized chart-cache values, stable label bindings, unchanged sibling categories, reopened file, and existing discovery gap')
