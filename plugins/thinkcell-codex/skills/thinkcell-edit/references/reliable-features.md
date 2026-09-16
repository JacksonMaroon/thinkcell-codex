# Reliable formatting and native features

## Series and point colors

For verified ordinary bar/column and stacked 100%-column donors, use `scripts/feature_pipeline.py` to enable the
existing native datasheet-fill setting and run one complete all-chart update.
Colors travel through documented think-cell JSON `fill` properties. Do not
recolor the PowerPoint cache. `multi_chart_update.py` checks the resulting RGB
mapping after generation and native reopen; inspect its preview as well.

Every target retains its normal `selector`, `name`, and `data` contract. Its
optional `appearance` has this form:

```json
{
  "schema": "tc.datasheet-fill.v1",
  "series_fills": {"Revenue": "#45225B", "Cost": "#A6A6AA"},
  "point_fills": {}
}
```

Wrap the complete data plan in a feature plan, using the source hash and chart
tag returned by inspection:

```json
{
  "schema": "tc.feature-pipeline.v1",
  "source_sha256": "SOURCE_SHA256",
  "features": [{"kind": "datasheet_fill_enable", "selector": {"slide_number": 1, "shape_tag": "CHART_TAG"}}],
  "data_plan": {"targets": []}
}
```

Replace the empty example `targets` with every chart's full data contract and
the requested appearance. Run `python scripts/feature_pipeline.py --input
donor.pptx --expected-sha256 SOURCE_SHA256 --plan features.json --output
result.pptx --report result.json --execute`. All paths must be distinct. Without
`--execute`, validation is read-only apart from writing the requested report;
use a new report filename for execution.

Use the controlling deck's actual colors. user-supplied branding does not establish
which color should encode a particular metric or risk level. A series maps by
its unique name; a point override follows the requested category order. All
charts on the working slide must be explicitly covered by the data plan.
Inspection flags slides using datasheet colors. Continue their updates through
`multi_chart_update.py`; the single-target route rejects them rather than
silently resetting the palette.

On enabled charts, a partial appearance request preserves unspecified fills.
An explicit whole-series fill clears that series' old point overrides before
applying new point requests. Retained point fills follow unique category
identities; ambiguous or changed identities fail closed. A single point
override passed a changed-data native canary on the tested 100%-column donor.

## Labels and annotations

Keep native fields linked to the data. The existing `feature_api.py` controls
uniform native font size. Replacing linked labels with literal text is not a
supported shortcut. For the bounded numeric total operation, use
[total-label-precision.md](total-label-precision.md). General label content
controls still require their own proof: metadata readback alone is
insufficient, particularly for cached total labels.

Ordinary slide callouts belong in the presentation assembly plan with explicit
position, dimensions, font and order. Native chart annotations belong to
think-cell. Preserve their donor structure and verify their meaning after data
changes. Do not silently replace a native CAGR arrow or broken axis with an
ordinary shape and call it native support.

## Office coordination

The chart entrypoints share `office_operation_lock.py`. Parallel agents may
prepare data and plans, but only one operation may mutate Office. A timeout
records the still-running child and prevents another operation from colliding
with it. Recover only after confirming that exact child has stopped. Never
quit the user's PowerPoint application or retry an uncertain write blindly.

See [capability-matrix.md](capability-matrix.md) for the current tested boundary.

## Mapping boundary

Ordinary bar/column series map by complete, unique cache value vectors. Null, sparse, ambiguous, secondary-axis and broken-axis color mappings fail closed. The 100%-column route uses its explicit denominator contract. Native evidence covers horizontal stacked bars and 100%-stacked columns; other accepted ordinary subtypes still require their own native readback and preview. Data-only updates of unstyled charts retain their prior support boundary.

See [expanded native controls](expanded-features.md) for secondary axes, line colors, semantic scalar formatting and additional CAGR endpoints/date controls.
