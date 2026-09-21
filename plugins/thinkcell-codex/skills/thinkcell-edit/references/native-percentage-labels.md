# Native percentage labels

Use native direct-rendered labels when the input chart already has them. They display dynamic percentages without parentheses, invisible-character wrappers or overlay text. Inspection reports `native_percentage_labels`; the update command automatically checks those labels again after official JSON regeneration and native save/reopen, before delivering the output.

## Existing chart

Generate the canonical request using `inspect --request-out`, change both the matrix and expected model, and run the normal `update` command. Add `--require-native-percent` to reject an input that lacks verified native percentage labels. Exact target, source hash, data, sibling and native checks still apply. Review the rendered slide.

## New percentage chart from a donor

Choose an authorized donor whose chart type, axes, geometry and formatting match the intended exhibit. The selected slide must contain exactly one chart. Generate its canonical request, then:

```powershell
python scripts/thinkcell.py create --input donor.pptx --expected-sha256 HASH --slide-number 1 --native-percent --data-json request.json --output-directory new-percent-chart --execute
```

This copies the whole native slide, preserves its direct percentage-label format, updates data through official JSON, saves/reopens and verifies the result. The returned output is `populated.pptx`; review its preview and update inherited titles/notes as needed. Without `--execute`, request and donor checks run without writing files. Without `--data-json`, the command copies and checks the donor only.

## Scope

The tested profile is a three-series, three-category native 100% stacked-column chart with whole-number percentage labels. Numerator and denominator updates produced 50%, 67%, 50% after save/reopen. Discovery requires exact chart ownership, native percentage-axis semantics, direct precision and unambiguous numerical cache agreement. Ambiguous direct percentage candidates are rejected rather than silently sent through a workaround.

## Convert an existing absolute-label chart

The experimental `convert-percent` command binds the existing scalar label to its own native relative source, using a genuine `CTextVariable` and a coherent native physical field. It retains the absolute axis and native plot; it does not change the chart into a 100% stacked chart. No parentheses, invisible wrapper characters or static text overlay are used.

```powershell
python scripts/thinkcell.py convert-percent --input absolute.pptx --expected-sha256 HASH --all-labels --output-directory converted-chart
```

Use `--category-index 0 --series-index 2` instead of `--all-labels` for one selected label; indices are zero-based. This command executes the full conversion, writes `converted.pptx` only after native verification, and returns the render for visual review. All-label conversion prepares the nine bindings together and regenerates once. Subsequent ordinary `update` calls require the same native relative-source identities, precision, exact displayed ratios and absence of wrapper text.

The currently guarded conversion profile is one internal, consistent ordinary stacked-column chart; three categories and three series; nonnegative values; positive category totals; an ordinary absolute axis without breaks; zero-decimal percentage labels; supported model versions 38775/38772; and native segment fonts using the tested `bg1` color. The Windows conversion evidence uses model 38775. Unknown profiles are rejected before native generation. The source remains unchanged.

Selected-label conversion preserves the original cache frame as well as the native axis and plot. Converting every label lets think-cell recompute the chart-cache container; the native plot coordinates, data, axis and visible segment positions remain the governing geometry checks. This cache-container change is reported explicitly.

Combined labels, custom denominators, negative values, other chart families and migration of existing wrapper labels are not certified by these tests. Keep the earlier bounded field-backed routes for those configurations where they have their own evidence.
