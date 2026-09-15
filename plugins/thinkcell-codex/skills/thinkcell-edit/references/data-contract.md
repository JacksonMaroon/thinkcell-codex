# Data request

`inspect --request-out` derives a request from the selected chart's actual embedded datasheet. Preserve its row slots, optional rows, and matrix dimensions. This avoids guessing from whether the chart appears horizontal or vertical.

The request contains exactly `matrix` and `expected_model`. Matrix cells are null, strings, or finite numbers. Booleans and formulas are not accepted. Use raw numeric values, not display strings such as `$12M` or `15%`. A percent's stored value must follow the source convention.

The matrix drives official JSON generation. The expected model is a separate assertion of the intended categories and values, checked against the result. Update both consistently from the user's data. The wrapper converts cells into think-cell's typed JSON; do not pre-wrap them as `{"number": 1}`.

## Pie

```json
{
  "matrix": [[null, null], ["Product A", 60], ["Product B", 40]],
  "expected_model": {"categories": ["Product A", "Product B"], "values": [60, 40]}
}
```

Pie JSON uses category rows with label/value columns, even when the embedded source datasheet is horizontal. Preserve the header and category count. For a uniformly percentage-based pie or doughnut, the request retains fractions (0.05 means 5%) and the helper emits official percentage points. Fractions must sum to 1 and the source total must be 100. Mixed absolute/percentage cells are rejected. Hole size and exploded-slice settings are preserved. The example illustrates the contract, not a universal two-category template. Inspect labels and segments after a change.

## Sequence charts

Expected fields are `categories`, `series_names`, `series_values`. JSON uses series rows and category columns. The source datasheet may be transposed; the inspector normalizes that orientation. Keep the source's category and series counts. If an explicit numeric `100%=` row exists, the request includes `category_extents`; update that row and expected field together when the denominator changes. A denominator need not equal the segment sum, so do not normalize it silently. Other reserved rows stay fixed. Series labels must be unique. Matrix row order controls the donor slots. Internal model series can appear in a different order, especially for horizontal bars; verification matches their full values by unique label and checks the exact datasheet row order separately. Do not turn a reserved row into a series.

Recognizing `CSequenceChartSE` alone does not establish correct waterfall or Mekko semantics. Read [specialized-charts.md](specialized-charts.md) for their experimental contracts and review requirements.

## Scatter and bubble

Expected fields are `point_labels`, `group_labels`, `x_values`, `y_values`, `size_values`, all with the same nonzero count. JSON uses five columns: label, X, Y, size, group. Preserve the source point count and reserved rows. Scatter sizes are null; bubble sizes are positive finite numbers. Keep the source's size-as-area interpretation. Do not switch scatter/bubble type. Fixed axes can clip a newly moved point even when its numbers are correct, so inspect the render.

## Names and links

The default namer resolves the chart through its native shape tag and data-table owner. It reuses a unique existing automation name or assigns one to the unnamed chart and its table on a local copy. Ambiguous ownership, mismatched names and duplicate names stop the update.

The presence of an internal embedded workbook is normal. The helper checks the actual external-link reference; a generic storage flag alone does not establish a persistent Excel link. Persistent or unknown links, including links on sibling charts, are rejected to avoid detaching them during generation.
