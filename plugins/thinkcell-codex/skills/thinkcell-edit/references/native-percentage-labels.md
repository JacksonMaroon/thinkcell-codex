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

This is not an absolute-to-percentage conversion adapter. An ordinary stacked chart can also display percentage labels while retaining an absolute axis; replacing it with a 100% stacked donor would change its meaning. Retain the existing bounded field-backed route for such charts until same-chart conversion is validated. Combined labels, custom denominators and arbitrary chart-family conversion are not certified by these tests.
