# Waterfall and Mekko contracts

The data-update route preserves the native donor's category/series slots. The structure-plan route below prepares supported changes to those slots. Both use an internal embedded datasheet and require native save/reopen plus visual review. Native-cache parity is **not implemented** for these specialized charts: waterfall Office caches encode offsets, and Mekko charts render as native shapes. Structural, model, datasheet, source, theme and notes checks remain active.

## Waterfall

The request matrix retains literal `"e"` cells, while `expected_model.series_values` contains the independently calculated numbers expected in those slots. Do not replace `e` with a fixed number just to satisfy a validator. Existing equals positions, group grounding and connector endpoints must stay unchanged. The reader resolves endpoint identity by category and group, so native object-ID regeneration does not hide topology changes. Unrecognized endpoint layouts stop with a specific error.

For a standard forward bridge, 100 + 20 - 5 gives a total of 115. For a backward build-down with an equals total first, infer that total from the signed steps and their actual connector direction. For an intermediate subtotal, do not add the subtotal again to the running total. Compute each expected value from the specific donor's connector structure; do not use a universal cumulative-sum formula on arbitrary waterfalls.

Inspect the render for grounded totals, connector direction, signed steps and mixed-sign stacks. The existing-data route preserves connectors and equals positions. Use a structure plan for adding a step, subtotal or series.

## Mekko with percent axis

The matrix uses nonnegative absolute segment values. Each expected column width is the sum of its category's segments. Include `column_widths` in `expected_model`. For values [30,70], the column total/width is 100 and shares are 30%/70%. The existing reserved row must be empty; typed percentage-input and populated `100%=` modes are not implemented.

## Mekko with units

The matrix's second row is the native X extent row. Change this row and `expected_model.column_widths` together. Widths must be positive. Heights and widths are independent: increasing width must not silently increase the segment heights. Nonnegative absolute segment heights are supported.

For both Mekko types, check displayed width ratios and segment shares/heights after generation. The structure-plan route supports category additions and width changes. Zero/negative widths, signed heights and chart-type conversion still need a separate contract.

## Structure plans

`scripts/structure_plans/build_structure_plan.py --request <canonical.json>
--operation <operation> --output-plan <fresh.json>` builds a typed plan. Use
`--help` for the operation-specific names, values, widths and subtotal range.
`build_add_step_job.py` also builds the official job for a waterfall step.
These helpers prepare inputs; they do not by themselves certify a presentation.

Supported controls include waterfall step addition, an intermediate subtotal
over a prefix range without an existing equals cell, and series addition while
retaining the reserved row. Literal `e` cells stay literal; expected values
are independently computed. Subtotal and series additions passed subsequent
data changes with totals, connectors and grounds retained.

Percent Mekko category additions calculate widths from the supplied segment
totals. Units Mekko category additions and width edits preserve independent
heights. Both passed changed-data model/datasheet readback and native reopen.
Use the plan's `request` object as the transformed canonical data request.
Run that request through official regeneration and the specialized
contract, then review the final chart before delivering it.

Official semantics: [waterfall](https://www.think-cell.com/en/resources/manual/waterfall), [Mekko](https://www.think-cell.com/en/resources/manual/mekko).
