# Capability discovery and execution

Run these commands from the plugin directory. Every route retains its exact
source identity, selector, working-copy and operation-specific verification
requirements. The registry is packaged availability, not proof that a machine
has passed native tests.

```powershell
python skills/thinkcell-edit/scripts/thinkcell.py capabilities
python skills/thinkcell-edit/scripts/thinkcell.py capabilities --id gantt-taskbar
python skills/thinkcell-edit/scripts/capabilities.py check
python skills/thinkcell-edit/scripts/thinkcell.py route gantt-taskbar -- --help
```

`route` passes arguments to the listed adapter without injecting `--execute` or
changing its contracts. It propagates failures. Research-only entries refuse
execution with a precise limitation. `prepare_only` describes an intermediate
copy that needs further gates; it never means a finished native chart.

## Selected-chart geometry

`chart_geometry_cli.py make-plan` requires the exact source SHA-256, a selection
JSON, an outer-frame JSON and a fresh plan output. Optional translation or exact
plot bounds are mutually exclusive. The selection comes from inspection; the
frame contains `left`, `top`, `width` and `height` in points. Only exclusive
native coordinate anchors are permitted.

```powershell
python skills/thinkcell-edit/scripts/thinkcell.py route chart-geometry -- make-plan --input source.pptx --expected-sha256 HASH --selection-json selector.json --frame-json frame.json --plan-out geometry.json
python skills/thinkcell-edit/scripts/thinkcell.py route chart-geometry -- prepare --input source.pptx --plan geometry.json --output prepared.pptx
python skills/thinkcell-edit/scripts/thinkcell.py route chart-geometry -- verify --input native-saved.pptx --plan geometry.json
```

Run official regeneration and native save/reopen between preparation and
verification. The verifier checks the selected plot, legend and containment;
it does not independently attest that native reopen happened.

## Gantt edits

Use `gantt-taskbar` and `gantt-milestone` with their existing date/style selectors
and exact source digest. Both prepare separate copies. Geometry is derived from
the source's bound native calendar and physical shapes. Reversed dates, ambiguous
binding, incompatible marker styles and unsafe outputs fail closed.

`gantt/verify_prepared_candidate.py --source SOURCE --output PREPARED --report
EDIT_REPORT` checks the report-bound dates, model/visible geometry, owning
bindings and untouched package content without opening Office. Its status is
explicitly offline. After native save/reopen, independently compare the saved
target fields and siblings and review the render. The release evidence records
current taskbar and triangle-milestone controls. The separate
`gantt/dependency_adapter.py` inspects actual CFB native graph/date/shape bindings
and prepares experimental authentic range-anchor candidates. Inspection is:

```powershell
python skills/thinkcell-edit/scripts/gantt/dependency_adapter.py --source source.pptx --expected-sha256 HASH --inspect
```

[Native dependency controls](gantt-dependency-evidence.json) distinguish a pristine
typed-date May 28 update reaching native model and XLSB from a rejected guessed
scalar-anchor candidate that skipped the update. Automatic reflow remains
unverified. No guessed scalar-anchor schema is shipped; authenticated scalar
ownership/binding and repeated native event controls remain required. No lag
scheduling or duration preservation is claimed.

## Existing Excel-link batches

The separate `excel-link-batch-rebind` route stages every requested existing link
and publishes only after the complete batch passes. Use a manifest like this:

```json
{
  "schema": "thinkcell-batch-rebind-v1",
  "links": [{
    "source_workbook": "C:/work/source.xlsx",
    "source_workbook_sha256": "SOURCE_WORKBOOK_HASH",
    "target_workbook": "C:/work/target.xlsx",
    "target_workbook_sha256": "TARGET_WORKBOOK_HASH",
    "guid": "INSPECTED_LINK_GUID",
    "range_name": "INSPECTED_RANGE_NAME"
  }]
}
```

```powershell
python skills/thinkcell-edit/scripts/thinkcell.py route excel-link-batch-rebind -- --input-presentation source.pptx --input-sha256 HASH --manifest links.json --output-presentation prepared.pptx --report rebind.json
```

Each request requires at least one exact GUID, link ID or range-name selector,
and each selects a distinct existing carrier. Source workbook paths must match
the binary moniker. Defined-name sheet/range identities must agree. Path changes
retain every other native identity field. Recognized length-prefixed monikers
can rebuild shorter or longer UTF-16 paths; opaque legacy monikers retain the
equal-byte-length fallback.

Native release evidence covers one existing link through the batch wrapper,
with automatic updating disabled and its identity retained after save/reopen.
Two-carrier native controls now passed shorter/longer paths including Unicode
surrogate characters, with complete GUID/moniker/table/grid identities, external
workbook versus model/embedded-XLSB parity, unchanged sibling models and sealed
sources. Independent automatic refresh controls passed first 17/second 10 then
first 17/second 29, with untargeted link, ordinary/sibling dependencies and unrelated
workbook cells/formulas/names/styles retained. See [expanded controls](expanded-links-evidence.json).
New-link preparation and refresh use the separate
commands below. Do not use
the internal-datasheet update route for linked charts. The official
[UpdateBatch documentation](https://www.think-cell.com/en/resources/manual/exceldataautomation)
states that updating Excel-linked elements through that API breaks their links.

## Error-bar readback

The authentic Min/Max/Marker donor route now exposes a read-only model/cache
grader. It selects the named chart on the requested slide, checks exact vectors,
native Min/Max range ownership, signed custom extent semantics and Marker values.
It accepts the documented equivalent positive/negative native carrier orientation.

```powershell
python skills/thinkcell-edit/scripts/errorbars/reusable_errorbar_route.py --verify-native native-saved.pptx --expected-sha256 HASH --data expected.json --name EXACT_CHART_NAME --slide-number 1
```

`ERRORBAR_MODEL_AND_CACHE_GATES_PASS` still needs independent save/reopen evidence
and visual review. The 0.4.1 native controls widened all seven ranges while
retaining Marker values; automatic axis rescaling changes screen positions.
Clean-chart insertion and new cap styling use the separate insertion route below.

## Complete source-slide sequencing

`assemble_native_slides.py` sequences at least two complete one-slide PPTX files
through official think-cell template generation. Its manifest is
`tc.slide-sequence.v1`, with a `slides` array containing unique `id`, source
`path` and SHA-256 `sha256` for each entry. Paths are relative to the manifest
directory or absolute. Sources must have identical slide dimensions and no
external dependencies.

```powershell
python skills/thinkcell-edit/scripts/assemble_native_slides.py --manifest sequence.json --expected-sha256 MANIFEST_HASH --output assembled.pptx --report assembly.json
python skills/thinkcell-edit/scripts/assemble_native_slides.py --manifest sequence.json --expected-sha256 MANIFEST_HASH --output assembled.pptx --report assembly.json --execute
```

Preflight writes nothing. Execution sends task-owned source clones to the
generator, checks complete slide and notes content, inherited theme/layout,
chart data/grammar and dependency closure before and after native reopen, and
withholds output on unexpected changes. Review every slide visually; the
automated preview covers the first slide. Optional per-slide `edits` use exact
ordinary-object selectors returned by `ordinary_slide_edits.py --input SOURCE
--expected-sha256 HASH`. An edit includes `shape_id` and `sha256`, then
the supported `op`, `text_runs`, `bounds_emu` or removal action documented by that
helper's contract. Native-owned, tagged and linked objects remain protected;
inherited placeholders cannot be removed. The edited source becomes the
expected saved-output snapshot, while native models and dependency closures
must retain the original state. A current native control passed text, geometry,
deletion and explicit textbox/rectangle addition with complete native/cache,
physical and dependency gates and visual review. Earlier assembly evidence
belongs to its separate historical adapter. Additions require the explicit
`op: add`, new `shape_id`, `name`, `kind`, `bounds_emu`, `text`, `font`, `fill` and
`line` contract; font size uses hundredths of a point and bounds use EMUs.

## Installed-template creation

```powershell
python skills/thinkcell-edit/scripts/native_library_create.py list --type clustered-column
python skills/thinkcell-edit/scripts/native_library_create.py create --type clustered-column --data-json data.json --output-directory new-chart --execute
```

Discovery reads actual native and visible subtype identities. Creation extracts
a whole native slide from a locally installed template and populates it through
official JSON/native checks; optional `--template`, `--slide-number`,
`--library-root`, `--style-file` and `--require-feature` constrain selection.
No vendor assets are bundled or downloaded. The current clustered-column
control changed a two-series/five-category donor into three series/three
categories with exact model/datasheet/cache parity and retained native subtype.
Its inherited thousands scaling rendered small supplied values as zero labels,
so that output requires label/layout revision and visual review before client
use. This proves the bounded template route, not arbitrary direct construction.

## Native feature insertion

```powershell
python skills/thinkcell-edit/scripts/native_insertions.py --help
python skills/thinkcell-edit/scripts/native_insertions.py verify --help
```

Preparation requires `--source`, `--donor`, distinct `--output`, exact JSON
`--selector` and `--donor-selector`, both source/donor hashes, and `--feature`
chosen from `legend`, `errorbar-range`, `errorbar-caps`, `linear-trendline`, `power-trendline`, `exponential-trendline` or
`logarithmic-trendline`, `quadratic-trendline`, `cubic-trendline` or
`quartic-trendline`.
Optional target-series and donor-feature IDs resolve bounded graph selection.
The prepared graft is not a finished native result. Run official generation,
native save/reopen, then `verify` with the saved artifact's digest, exact selector
and canonical data request. Verify source/sibling preservation and a second
changed-data control before promotion; inspect the preview independently.

Native range and endpoint-cap insertion have passing bounded native canaries;
dash-cap changed-data extent 5 to 12 also passed canonical signed-range ownership,
all nine cap markers on both endpoints and visual review. A linear scatter insertion passed
native regression slope/intercept and authentic clipped forecast-domain grading;
its repeat native control passed selected-series Y +2 with unchanged slope,
intercept increasing by exactly 2, retained other series and visual review. The area-legend candidate lost its native graph
after reopen and is rejected. A bounded regular-bar insertion passed zero to one
native legend, matching swatch, Product B to Product B revised and first bar
5.1 to 6.8 repeat controls, with Product A sibling preserved, strict
ALL_GATES_PASS, native save/reopen and visual review. Its source required
explicit duplicate-axis-GUID repair and official unchanged-data schema
normalization before fresh graft; those preflight requirements remain explicit.
See [legend controls](legend-insertion-evidence.json). Broader families remain
unverified. Do not treat a
visible cached legend or an existing donor feature as proof of retained native
insertion.

Power scatter preparation uses `--feature power-trendline` with an authentic
native enum1 donor closure and finite positive-X/Y selected data. Native baseline
and selected Y times 2 repeat retained a native partition and physical power trendline. Independent
log-OLS/plot-bounds grading passed the changed fit-limited upper endpoint
5.764033425759087 versus reference 5.764033425759088. The reference coefficient
approximately 1.15122 to 2.302449 and exponent approximately 0.838422 are fitted
calculations, not serialized native coefficient fields. The baseline axis limit
9 alone does not validate the fit. Native power curve points are not serialized
(CPPTPolyline length 0, placed 0); the changed visible-Y-limit clipping supplies
one bounded numeric fit control. Other types are outside this route;
each candidate requires native, data, sibling and visual gates.

Exponential and logarithmic preparation use `--feature exponential-trendline`
and `--feature logarithmic-trendline` with actual saved native enum2/enum3 donors.
Require finite full-rank selected data, positive Y for exponential and positive X
for logarithmic. Independent fits use log(Y) on X and Y on log(X), respectively;
coefficients are reference calculations, not serialized native fields. These
verifiers require a finite increasing fitted curve, unique linear X/Y axes and
explicit bounds. Observed
exponential donor baseline endpoint 4.77347400977944 matches the reference;
logarithmic baseline endpoint9 is axis-limited and does not validate the fit.
Public clean-chart graft and native repeat controls passed exact cache/request,
source and Higher-end sibling preservation, with current visual review. Exponential
Y times 2 moved the fit-limited endpoint to 3.8960719529819787, exactly matching
reference. Logarithmic Y times 2 remained axis-limited; additional selected
X times 0.5 plus Y times 2 yielded discriminating endpoint 4.808574822433828,
exactly matching reference. This establishes one bounded numeric log-fit control.
See [nonlinear controls](nonlinear-insertion-evidence.json). Guessed enum4/5
candidates lack a native physical trendline and are rejected hypotheses; polynomial types have separate routes.

Polynomial preparation uses `--feature quadratic-trendline`, `cubic-trendline`
or `quartic-trendline` only from actual saved native enum6/7/8 donors with exact
physical polynomial order2/3/4. Require finite full-rank selected data. Independent
polynomial OLS and closest-domain root calculations grade forecast clipping;
reference coefficients are not serialized native fields. Require distinct X
count at least degree + 1, numerical conditioning guards, unique linear X/Y
axes with explicit min/max and no forced intercept. Empirical native donor
baseline controls matched clipped roots. Public zero-feature graft baseline and selected Lower-end Y plus 0.5X native
repeat controls passed all three degrees, exact model/data/cache/reopen, sealed
source and Higher-end preservation, and current visual review. Actual fit-limited
endpoints matched independent closest-domain references within 1.2e-11.
See [nonlinear controls](nonlinear-insertion-evidence.json); reference coefficients
remain fitted calculations. Preexisting template instruction-box crop is outside
this visual scope. Forecast clipping plus visual shape is bounded evidence,
not an exact curve-point certificate or general engine certification. Guessed
enum4/5 hypotheses remain rejected.

The authenticated power, exponential, logarithmic and polynomial graft profiles
start at the selected series minimum X. Their verifiers require an absent or
exactly zero backward extension; nonzero, duplicate and nonfinite extensions
are rejected. Polynomial roots use scaled coordinates and residuals, exclude
isolated tangent points and retain visible reentry intervals. Endpoint tolerance
is relative to the X-axis span, including for small numeric domains.
Extreme subnormal coefficients or coordinate transformations that lose nonzero
terms through underflow or overflow are rejected.

## Persistent-link preparation and refresh

```powershell
python skills/thinkcell-edit/scripts/external_links/new_link.py --help
python skills/thinkcell-edit/scripts/external_links/refresh_link.py prepare --help
python skills/thinkcell-edit/scripts/external_links/refresh_link.py validate-plan --plan refresh-plan.json
python skills/thinkcell-edit/scripts/external_links/refresh_link.py execute --plan refresh-plan.json
python skills/thinkcell-edit/scripts/external_links/refresh_link.py verify --plan refresh-plan.json --presentation native-saved.pptx
```

New-link preparation needs hash-bound source and authentic donor presentations,
a source workbook, exact target/donor OLE parts and slide numbers, and fresh
presentation/workbook/report outputs. It preserves the external adviser/range
closure and stages a separate workbook. Refresh stages explicit cell changes,
checks workbook values/formulas/names/styles outside those edits, binds the
copied workbook, then executes native refresh under the shared Office lock.
Follow the strict request/plan schema shown by the adapter. Native changing
values alone is not enough: saved link/workbook/range identity, exact cell/model
readback and untouched sources/other presentations must all pass. A newly
authored persistent link passed repeated B2 changes 10 to 15 to 20 with exact
model/datasheet/cache strict parity, selected ordinary objects and unselected slide models/text/dependencies,
unrelated workbook cells/formulas/names/styles, source preservation and other
open presentation/workbook state. The local adviser numeric ID renumbered from
187 to 2 while its table/grid/registration closure retained semantic identity.
See [persistent-link evidence](persistent-link-evidence.json). These controls
cover the tested link/cell profile rather than arbitrary formula/range changes.

## Portable regression coverage

```powershell
python skills/thinkcell-edit/scripts/run_portable_tests.py --report C:/work/portable-tests.json
python skills/thinkcell-edit/scripts/run_portable_tests.py --fixture-root C:/authorized/historical-fixtures --report C:/work/fixture-tests.json
```

The runner includes standard tests, embedded appearance/render-tag tests and the
standalone relationship allocator regression. It reports unavailable historical
fixture profiles as explicit skips and propagates failures/timeouts. Source
fixtures are supplied locally and never included in this public distribution.
Portable passes cannot substitute for native regeneration, readback or visual
review.
