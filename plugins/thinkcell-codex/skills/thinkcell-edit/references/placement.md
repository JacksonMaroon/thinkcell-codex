# Placement in a finished slide

`scripts/inspect_layout.ps1` reads a saved source through a hidden, read-only
native staging copy. Supply `-InputFile`, `-ExpectedSha256`, `-SlideNumber`,
optional matching `-SlideId`, and a new `-OutputDirectory`. Run with Windows
PowerShell 5.1. It returns `layout.json` and `slide.png` without saving a deck.
The report's `shapes` are clipped on-slide occupancy candidates; `all_shapes`
retains every source shape for preservation, including invisible helpers and
offcanvas content. Native visibility and original rotated bounds are recorded.
Inherited master/layout shapes are reported separately because their actual
visibility and occupancy need visual review. Inspect the PNG and add occupied
inherited title/footer/source rails to `reserved` before planning. Design and
layout names are metadata, not proof of template equivalence.

Inspect the actual slide, its data, and the user's guidance. Retain every explicit
chart choice, position, dimension, and other requested feature. Decide only what
is missing. Choose the chart to support the slide's argument and the real data;
do not classify by title keywords. Record the reason for the choice. Resolve a
requested but unsupported feature explicitly rather than silently substituting.

`scripts/placement_plan.py` finds a conservative empty rectangle from observed
slide geometry. It does not select a chart, inspect a deck, generate data, or
execute native insertion. Its frame is the **outer chart footprint including
labels**, not just the plot. Inset the plot by the donor's measured surrounding
labels and annotations, then verify the returned native geometry and render.
Changed data can change label extents, so recheck the final chart's outer bounds.

Call `plan(context, brief)` or run:

```text
python scripts/placement_plan.py --context context.json --brief brief.json --output plan.json
```

`context` contains `source_sha256`, `slide_id`, slide `width` and `height` in
points, `shapes`, and optional `reserved` rectangles for title/source/footer
rails. Every rectangle uses `left`, `top`, `width`, `height`; every shape also
has a unique `id`. Supply all occupied content bounds. Do not omit content to
make a chart fit. Treat genuine background decorations separately when preparing
the context, based on visual inspection, and preserve them during insertion.

`brief` supplies `requested_chart_type` or `inferred_chart_type`, plus `rationale`.
An explicit requested type takes precedence. Optional `frame` may contain any
subset of the four rectangle coordinates. Optional `side` is `left`, `right`,
`top`, or `bottom`, meaning the corresponding half of the slide. Optional
`margin` (default 12 points), `gap` (6), `min_width` (72), and `min_height` (54)
control clearance and minimum outer dimensions. Explicit positive frame sizes
replace default minimums on their axis; an explicit minimum remains binding.
Default margins also yield to explicit canvas coordinates or dimensions;
an explicitly supplied margin, actual slide bounds, and occupied content remain binding.
Raise minimums for a dense chart
or donor whose labels need more room. The defaults are geometric lower bounds,
not a readability guarantee. Unknown brief fields fail rather than disappear.

The planner preserves explicit coordinates, considers the largest empty area,
and uses a moderate landscape ratio only to break area ties. Collision,
contradictory constraints, invalid geometry, and insufficient space fail without
moving or shrinking existing content. Read back source identity before executing
the plan. The plan records a context digest and choice provenance and always
requires native geometry readback and visual review.
