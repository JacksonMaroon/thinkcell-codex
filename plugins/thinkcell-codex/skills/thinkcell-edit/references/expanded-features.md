# Additional native chart controls

Use the chart already on the slide as the formatting authority. These adapters
edit the native model and its bound PowerPoint shapes, then require official
regeneration, native reopen and feature readback. Use fresh outputs and the
shared Office lock. The existing data and composition routes remain available.

## Secondary value axis

`scripts/axis_range_secondary.py` provides `make-plan`, `prepare` and `verify`.
Select `--axis-role secondary` with a maximum and major unit. Ownership comes
from the chart's secondary axis and series references. The primary axis is
unchanged. The default remains the existing primary-axis route.

The dual-axis column/line control passed at 0–150 in steps of 25, including
changed Stock price values. The canonical data reader now distinguishes a blank
reserved row from an explicit percentage denominator; supplied numeric
denominators retain their original meaning.

## Line series colors

`scripts/line_style_production.py` accepts `--input`, `--expected-sha256`,
`--plan`, `--output`, `--report`, `--ppttc` and `--execute`. Its plan follows
the existing feature pipeline. Series RGB colors survived changed line values
and native reopen. Use the existing column route for its verified point fills.

## Scalar label formatting

`scripts/native_label_controls.py make-plan` accepts `--target-json` and
`--controls-json`; `prepare` applies the plan to a new package. Resolve labels
by category and series. Synchronize the selected model number-format key with
its physical dynamic field. The controls support existing-label precision,
prefix/suffix and requested font styling. Preserve all other fields.

The selected West / 2nd Qtr field updated from `$38.6m` to `$40.0m` and then
`$42.3m`, retaining one decimal, 10-point bold styling and the separate dynamic
relative-percentage field. Official regeneration and native reopen passed.
The `--experimental-insertion` mode is a separate research operation.

## CAGR endpoints and typed dates

`scripts/cagr_endpoints/run_endpoint_expansion.py` performs preparation,
regeneration, native reopen and independent readback. Supply `--input`,
`--expected-sha256`, `--data`, `--skill-dir`, `--ppttc`, `--workdir`,
`--source-index` and `--sink-index`.

The data JSON contains a chart `name`, ordered ISO `dates` and positive
`values`. The adapter discovers the chart's native endpoint ownership and
clones the complete annotation structure. The 0→6 control displayed 2.9%,
then recalculated to 5.6% after changing the endpoint. A January 15, 2018 to
July 15, 2024 control displayed 2.7% over 6.5 years. The previous 1→7 route
remains available. Read the saved native value and field, and inspect date
label fit. Stacked-total and cross-chart-family imports remain research.

## Add a CAGR to a chart without one

`scripts/cagr_import/run_import_cagr.py` accepts target and donor presentations,
their SHA-256 hashes, chart names, endpoint indices, the think-cell skill
directory and a fresh working directory. It copies the complete native graph,
arrow and field structure with unique model and relationship identities.
Run the prepared candidate through `cagr_endpoints/run_parameterized_control.py`
and `verify_cagr_result.py`; require native readback and visual review.

The clean no-CAGR target retained exactly one dynamic annotation after import,
recalculated to 5.6% with changed data, and contained no stray donor label text.
Total-owner hypotheses are disabled in the production importer.

## Clean line-family CAGR

`cagr_import/portable_line_family_cagr.py` prepares a complete native CAGR
closure for a clean ordinary line chart. Supply explicit target and donor
presentations, their SHA-256 guards, chart names, endpoint indices, the
installed `--skill-dir`, a fresh `--workdir`, and output/report paths. The
selected-series extension additionally takes `--series-index`, `--series-name`
and requires `--donor-sha256`. The default scalar route remains valid when
`series_index` is omitted. The adapter verifies the actual related chart part
is a line chart and rejects aggregate totals or ambiguous series. Native
regeneration, reopen and changed-data checks remain required for each source.

## Percentage fields

`scripts/percent_labels/prepare_percent_candidate.py` selects a chart label by
category and series and keeps its genuine relative-value field. Use
`--wrapper parentheses` for the verified route. `run_percent_candidate.py`
performs regeneration and native readback. The one-decimal controls correctly
displayed `(12.5%)`, `(15.0%)` and `(10.0%)` as numerator and denominator changed.
The existing zero-decimal route remains available. A direct wrapper-removal
mutation failed native regeneration and is not a supported route. The
invisible U+200B wrapper route below passed selected native controls; its
broader 12-label batch remains failing and excluded from the supported scope.

## Bare percentage display

`percent_labels/prepare_bare_display_candidate.py` is a bounded route for a
bare visible percentage when parentheses are explicitly excluded. It first
uses the semantic parenthesized relative-field preparation, then replaces
only the two physical wrapper runs with U+200B ZERO WIDTH SPACE. The model,
relative field identity, format and placement stay native. The script accepts
`--input`, `--output`, `--report`, `--expected-sha256`, `--category`,
`--series`, optional `--chart-name` and `--digits`.

The script resolves the production scripts rail relative to its installed
`scripts` directory by default; `THINKCELL_PLUGIN_SCRIPTS` may override that
rail with an explicit directory. Do not use a fixture path or a fixed chart or
shape identity in the installed route. The strict grader must use its explicit
zero-width allowance and verify raw U+200B prefix/suffix, one native relative
field, exact changed numerator/denominator, visible bare text and sibling
preservation. The parenthesized route remains the default when bare display is
not requested. The selected controls do not establish general bare-label
support. Any new source or selector requires fresh native regeneration, reopen
and the same grader.

## Existing linked workbook rebind

`external_links/portable_rebind.py` rebinds one existing think-cell Excel link
offline. It requires explicit input/output presentation paths, source and
target workbook paths, SHA-256 guards and a report path. Optional GUID,
link-ID, range-name and expected-range guards narrow the carrier. The adapter
requires the target defined name to match the source sheet and range and
replaces only the UTF-16 workbook path in the moniker. The replacement path
must have the same UTF-16 byte length; source presentation and both workbooks
are read back unchanged. It does not create links or update a datasheet.

## Authentic error-bar donor

`errorbars/reusable_errorbar_route.py` prepares a fresh named copy from an
authentic installed donor with an existing native Min/Max/Marker range. Supply
explicit source, data, output, plan, name, slide and shape selectors. The
offline gate verifies model ownership, exact donor vectors and the signed
DrawingML custom extent carrier. Native same-data and changed-data readback
passed; the changed render uses a 0–11 value-axis range while Marker remains
fixed. The route does not insert error bars into a clean line chart or create
new cap styling. `native_verify_scoped.ps1` accepts optional
`-RenderSlideIndex` (default 1) for target-slide renders.

## Existing native legend translation

`legend_controls/legend_high_adapter.py` prepares a translation plan for an
existing native legend selected by chart and legend identity. It updates the
model anchors and their physical tagged union, while preserving the chart
carrier and rejecting shared-constraint ambiguity. The normalized native
baseline, translated same-data and changed-data controls, and portable
semantic identity all passed. Use a fresh output and require native reopen and
visual review. New legend creation and arbitrary grouped-layout translation
remain outside the tested scope.

## Official scatter trendline donor

`trendline/portable_trendline_implementation.py` prepares a fresh copy from an
authentic scatter donor. Supply explicit `--source`, `--output`,
`--shape-tag`, `--slide-number` and optional installed implementation path.
The adapter resolves the selected chart and preserves its two visible linear
partitions, native polylines and physical freeforms. The native baseline and
changed-data controls passed reopen and independent grading: Higher-end moved
by +2 Y, Lower-end by -1 Y, and unrelated native embeddings stayed equal.
Require native regeneration, reopen, render and changed-data readback for a
new source or selector. Other trendline types and clean-chart insertion remain
outside the tested scope.

## Gantt taskbars

`scripts/gantt/run_paired_gantt_edit.py` accepts `--source`, `--expected-sha256`, `--output`,
`--select-start`, `--select-end`, `--start-date`, `--end-date` and `--report`.
It selects one taskbar by existing dates, reads the native weekly calendar,
and updates both typed dates and the bound model/visible geometry. It rejects
missing or irregular calendar identity. Reopen the prepared output natively
and verify the dates, geometry and non-target records before delivery.
This route does not rely on automatic Gantt reflow. Milestone and dependency
editing require their own adapters.

## Gantt milestones

`gantt/run_milestone_edit.py` is the separate bounded milestone route. It
accepts `--source`, `--expected-sha256`, `--output`, `--report`, `--new-date`,
`--milestone-date`, `--milestone-id` and `--style`. It selects one existing
typed milestone, updates its native date and model marker rectangle, and
updates the bound visible marker. The helper
`gantt/replace_ole_stream.ps1` is resolved beside the adapter, while
`runtime.powershell_env` is resolved from the parent `scripts` directory.

The verified triangle control selected milestone `174` at `2012-02-06` and
moved it to `2012-02-13`. Native reopen and independent readback confirmed the
target date, model bounds, marker style/tag and visible center, and preserved
the other milestone, all taskbars and existing shapes. This route does not add
milestones, reflow dependencies or infer arbitrary marker styles. See
`gantt-milestone-evidence.json` for exact evidence hashes.

## Waterfall and Mekko structure

`scripts/structure_plans/build_structure_plan.py` creates a plan from a canonical
request. Waterfall plans can add a step, intermediate subtotal or series;
Mekko plans can add a category or update widths. Preserve literal equals cells
and compute expected totals from the chart's connector structure. Percent
Mekko widths equal category totals; units Mekko widths remain independent of
segment heights. See [specialized charts](specialized-charts.md).

These are plan builders. Apply the plan through official regeneration, then
verify the saved native model, complete datasheet, topology and render. Added
waterfall subtotal/series and both Mekko types passed changed-data controls.

## Multiple existing breaks

`scripts/multi_breaks/run_multi_break_repeat.py` updates a chart with existing
multiple break owners. The two-break control retained both owners and their
category bindings after changed data; think-cell recalculated automatic gaps
to [45,100] and [0,18]. Verify breaks by owner and crossed categories, not by
unordered gap values. This is a repeat/update route; use the separate break
insertion adapter when adding a break.

## Native series connectors

`scripts/series_connector_adapter.py` prepares a new series connector using
target/donor charts selected by name and series/category labels. It imports
the complete native connector and bound shape with fresh identities. Run the
candidate through the normal official data update and native verification.

The new interior connector and pre-existing connector survived regeneration.
Changed values moved both interior endpoints by the expected stacked boundary
proportions, with consistent plot scaling and preserved data. Inspect the
endpoint geometry and render after updating the chart.

## Runtime behavior

The native verifier releases its owned COM references and closes only its
scratch presentations. It checks source and pre-existing presentation state.
It never quits PowerPoint, changes application visibility or uses the mouse.
Independent file preparation can run concurrently. Native calls share a lock
because independent automation clients can attach to the same PowerPoint
process. Separate windows do not establish process isolation.
