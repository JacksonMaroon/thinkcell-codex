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
current taskbar and triangle-milestone controls; dependency creation and automatic
reflow remain unimplemented.

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
retain the same UTF-16 byte length and every other native identity field.

Native release evidence covers one existing link through the batch wrapper,
with automatic updating disabled and its identity retained after save/reopen.
Multiple-carrier batches and surrogate-character paths have portable controls
only. Arbitrary-length rebinds and new-link creation remain research. Do not use
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
Clean-chart insertion and new cap styling are outside this donor route.

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
automated preview covers the first slide. This public route preserves existing
source content. The earlier ordinary-object editing assembly remains historical
evidence for its separate implementation.

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
