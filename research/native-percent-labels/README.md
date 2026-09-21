# Native percentage-label experiment

Date: 2026-09-21. Mac think-cell 14.0.38775, PowerPoint 16.113.1. Synthetic native charts only. The Windows continuation adds read-only discovery; percentage conversion remains unsupported.

Start with [the Windows handoff](WINDOWS-HANDOFF.md) to continue this work. The numbered decks and extracted structures are committed directly so a normal clone includes the evidence.

## Finding

A freshly inserted native 100% stacked-column chart renders bare dynamic percentage labels without parentheses, zero-width wrappers, separate relative CTextVariable objects, or standalone PowerPoint label shapes. The labels survived two deliberate data changes and native close/reopen.

This supplies a concrete example for the direct-precision rendering form already mentioned in the repository's research notes. It does not prove a new general conversion algorithm.

The original reviewed public package's `percent_labels/discover_percent_semantics.py` discovers nine scalar labels but reports null relative text variables and no physical shapes for all nine. Its selector rejects Series 3 / 2024 with `selected label tag does not resolve exactly one physical shape`. See `existing-discovery.json` and `existing-selector-result.json`, which records the inspected source hash.

## Native structure

- Model version is 38775. Nine `CSequenceChartDataScalarLabel` objects have `m_bMSGraphRendering=1` and direct `m_prec` with empty prefix, `%` suffix, and zero decimal places.
- Their relative sources have empty `m_ctextvar` collections. These scalar labels are still dynamic.
- The native chart part contains normalized values on a 0Ã¢â‚¬â€œ100 scale. Its data labels use `c:showVal=1`, `c:showPercent=0`, and number formatting with a literal percent suffix. No `a:fld` appears in that chart part.
- Example: after the final edit, the native model stores 125, 75, 200. The chart cache stores 50, 18.75, 31.25 in reverse series order. The native screen displays 50%, 19%, 31% in bottom-to-top order.
- The ordinary absolute-label control also uses MSGraph rendering. Therefore this flag alone, a percent suffix alone, or the absence of a field cannot establish dynamic percentage semantics. Classification must include native ownership, axis/data semantics, and numerical agreement.

## Fixtures and controls

| File | First-column model values, top-to-bottom series order | Target Series 3 label |
| --- | --- | --- |
| 00-native-absolute.pptx | 3.5, 5.1, 14.9 | Absolute control |
| 01-native-percent.pptx | 4.37, 5.98, 12.65 | 55% |
| 02-controlled-percent.pptx | 25, 75, 100 | 50% |
| 03-numerator-change.pptx | 25, 75, 200 | 67% |
| 04-denominator-change.pptx | 125, 75, 200 | 50% |
| 05-reopened-saved.pptx | 125, 75, 200 | 50%, after reopening and saving |

Files 00 and 01 were independently inserted through the chart gallery. They are not a controlled single-setting conversion pair. Files 02Ã¢â‚¬â€œ05 are successive versions of the same percentage chart. Labels on the other two categories and their underlying values remain unchanged. The final screenshot is `04-reopened.png`.

All values were changed through the native Excel datasheet. Inspection verifies exact saved model values, nine chart-cache ratios, native consistency, stable label bindings, unchanged sibling categories, and persistence in the reopened file. Embedded workbook bytes are preserved in the PPTX files, but this experiment did not independently parse saved XLSB cells.

## Reproduce the read-only checks

Install `olefile` and `lxml` into a Python environment, then run:

```sh
python inspect_fixtures.py
python verify_fixtures.py
```

`evidence.json` includes fixture hashes, package-part hashes, model summaries and chart formats. `extracted/` holds readable native models and chart XML. `verification.json` records the tested contract. The verification script intentionally applies only to these exact fixtures.

## Implementation status

Read-only discovery now reports direct precision, exact owning-chart identity, native percentage-axis evidence and a unique numerical cache mapping. It fails closed for ambiguous mappings and bounds permutation search to 100,000 combinations; it does not assume reverse order. Field-backed mutation selection is unchanged. See [Windows results](windows/README.md) for native reopen and independent JSON regeneration results.

## Further work

1. Add a real field-backed deck fixture to complement the synthetic selector guard tests. No such distributable fixture is present in this public repository.
2. Establish a same-chart UI conversion pair and reproduce it on Windows before attempting a production mutation adapter.
3. Repeat native datasheet UI changes on Windows independently from the completed JSON regeneration tests.

## Remaining limitations

The floating Label Content toolbar could not be captured reliably through the available Mac computer-use interface. Therefore absolute-to-percent conversion, combined labels, custom formats, and removal of the existing workaround remain unproven. This experiment used default native percentage labels instead. The initial Mac experiment did not test Windows COM or `.ppttc` regeneration; the subsequent [Windows run](windows/README.md) covers these exact fixtures and build only. Do not patch production labels by copying chart-cache formatting or removing validation checks.

Official background: https://www.think-cell.com/en/resources/manual/textlabels

## Conversion continuation

A subsequent same-chart conversion adapter now passes native save/reopen and changed-data checks for the guarded absolute 3x3 stacked-column profile. It converts one or all nine labels to genuine bare relative fields, without changing the native plot or absolute axis. See [conversion evidence and real field-backed fixtures](conversion/README.md). The earlier missing-conversion findings above describe the initial direct-label experiment, not this later adapter.
