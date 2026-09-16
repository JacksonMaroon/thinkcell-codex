# Multiple native sequence charts

Use this bounded experimental route for a local, one-slide presentation with
ordinary `CSequenceChartSE` charts, or one fixed-topology waterfall plus
ordinary siblings, with a current compatible think-cell version,
matching slide dimensions/theme, and internal datasheets. It has passed native
gates for a bar plus stacked column, that pair plus a plain bar, a duplicate
donor, and a four-series/three-category change with the untargeted sibling
model and datasheet streams retained.

It is separate from [existing-slide.md](existing-slide.md), which remains the
route for adding one donor chart to a chartless ordinary slide.

## Update every chart explicitly

Create a complete plan with one selector, name, and canonical data request per
chart. The update must name every native chart on the slide, then generate with
official JSON and pass exact all-chart, strict cache, native-reopen, and render
gates:

```powershell
python scripts/multi_chart_update.py --input input.pptx --expected-sha256 HASH --plan plan.json --output output.pptx --report report.json --execute
```

Set `allow_sequence_count_change` only for a tested simple header plus optional
`100%=` row plus series-row layout. For a blank `100%=` row, leave denominator
cells blank only when the intended denominator equals the calculated category
total; otherwise provide positive explicit denominators and matching
`category_extents`.

A waterfall must retain its exact datasource dimensions and its existing equals
cells, connector topology, and grounds. Its native cache uses an offset
representation, so the route instead checks those semantics and its exact
datasheet. Every ordinary sibling still receives native model/cache parity.

An output is usable only when the report status is `ALL_GATES_PASS`. A timeout,
unknown external link, unsupported sequence layout, unlisted chart, or failed
native check leaves no deliverable.

## Compose one compatible donor

The staged constructor appends one compatible sequence-chart donor to a
receiver that already owns N sequence charts. It closes the ownership graph in
a new native namespace, including when donor and receiver started with the
same names or tags:

```powershell
python scripts/compose_sequence_charts.py --receiver receiver.pptx --receiver-sha256 RECEIVER_HASH --donor donor.pptx --donor-sha256 DONOR_HASH --output candidate.pptx --report candidate-report.json
```

With `--execute`, this command creates the candidate, derives a complete
all-chart plan, runs official JSON, and releases its distinct output only when
the report status is `ALL_GATES_PASS`. The portable-pipeline native canary
passed this end-to-end path. It also remaps and verifies native GUID and hidden
shape-name registries; any duplicate key rejects the candidate. Without
`--execute`, it is preflight only. The
preflight permits at most one fixed-topology waterfall across receiver and
donor. The mixed composition path passed with one waterfall plus two ordinary
charts, including official regeneration, native reopen, exact datasheets and
waterfall topology checks. Ordinary charts retain their exact cache parity gate.

## Move a selected chart without moving siblings

`scripts/chart_geometry.py` supports a selected native chart on a multi-chart
slide when its plot and attributable detached-legend coordinate grids are
exclusive and it has no shared layout constraint. It passed translation and
exact-plot-bounds canaries. The route proves all nonselected model content is unchanged, regenerates
through official JSON, and verifies the selected chart by durable carrier,
owner, tag, and name after native regeneration.

Gantt, other specialized/non-sequence charts, persistent or unknown external links,
and arbitrary feature insertion are outside this route. Those are current
implementation and verification boundaries, not claims about every think-cell
capability.
