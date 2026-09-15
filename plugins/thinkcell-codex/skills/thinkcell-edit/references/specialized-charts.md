# Experimental waterfall and Mekko contracts

These routes preserve a native donor's structure. Each requires a single chart on the slide, an internal embedded datasheet, fixed category/series slots, and native save/reopen plus visual review. Native-cache parity is **not implemented** for these specialized charts: waterfall Office caches encode offsets, and Mekko charts render as native shapes. The report identifies this limit rather than reporting generic parity as passed. Structural, model, datasheet, source, theme and notes checks remain active.

## Waterfall

The request matrix retains literal `"e"` cells, while `expected_model.series_values` contains the independently calculated numbers expected in those slots. Do not replace `e` with a fixed number just to satisfy a validator. Existing equals positions, group grounding and connector endpoints must stay unchanged. The reader resolves endpoint identity by category and group, so native object-ID regeneration does not hide topology changes. Unrecognized endpoint layouts stop with a specific error.

For a standard forward bridge, 100 + 20 - 5 gives a total of 115. For a backward build-down with an equals total first, infer that total from the signed steps and their actual connector direction. For an intermediate subtotal, do not add the subtotal again to the running total. Compute each expected value from the specific donor's connector structure; do not use a universal cumulative-sum formula on arbitrary waterfalls.

Inspect the render for grounded totals, connector direction, signed steps and mixed-sign stacks. This release does not add or reroute connectors, change which segments are equals, or certify arbitrary waterfall topology.

## Mekko with percent axis

The matrix uses nonnegative absolute segment values. Each expected column width is the sum of its category's segments. Include `column_widths` in `expected_model`. For values [30,70], the column total/width is 100 and shares are 30%/70%. The existing reserved row must be empty; typed percentage-input and populated `100%=` modes are not implemented.

## Mekko with units

The matrix's second row is the native X extent row. Change this row and `expected_model.column_widths` together. Widths must be positive. Heights and widths are independent: increasing width must not silently increase the segment heights. Nonnegative absolute segment heights are supported.

For both Mekko types, check displayed width ratios and segment shares/heights after generation. Zero/negative widths, signed heights, category/series count changes and chart-type conversion require another contract. The same generic sequence-family label does not imply support for these cases.

Official semantics: [waterfall](https://www.think-cell.com/en/resources/manual/waterfall), [Mekko](https://www.think-cell.com/en/resources/manual/mekko).
