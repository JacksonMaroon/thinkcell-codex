# Native percentage-only labels, bounded route

`scripts/run_percent_only_parentheses.py` converts an existing dual label
`value<line-break>(percent%)` to `(percent%)` by deleting only the absolute
field and separator through PowerPoint. The surviving percentage is a genuine
native relative field, and remains responsive to numerator and denominator
changes. The parentheses are required by this verified route.

## Exact supported profile

One model-version-38764 ordinary stacked percent-axis chart; four series
`Not Mapped`, `High risk`, `Medium Risk`, `Low risk`; three categories
`Gross Profit`, `Revenues`, `Customer Count`. The selected label is the
Low risk / Gross Profit scalar 191, tag `tbUpiE_yCia4NQkLBVWFCIA`, in
`ppt/embeddings/oleObject13.bin`. Both original fields must have genuine native
bindings and nonnegative integer display text. Percentage precision is zero
decimal places. Changed values within this profile are supported.

This is a donor-profile adapter, not a general label-content command. It does
not cover arbitrary scalar IDs, all labels at once, other chart families or
percentage formatting without parentheses. Unsupported profiles must be
investigated on copies under the native-feature research workflow.

The reusable original donor is
`assets/feature-fixtures/percent-dual-38764.pptx`, SHA-256
`AED1B65F82F8C8F9CF0FCC78778549DA60FE4BFFC2B1C76F0C85C42A34753A1B`.
Inspect the actual input hash and provide a complete native update plan.
The fixture plans under `references/percent-only-evidence/` are test inputs,
not defaults for user data.

## Run

Create a new empty work directory. From the skill root:

```powershell
python scripts/run_percent_only_parentheses.py `
  --input dual.pptx --expected-sha256 <current-source-hash> `
  --plan requested-data.json --expected-display '15%' `
  --work-dir fresh-work --output percent-only.pptx `
  --report percent-report.json --evidence percent-evidence.json
```

`--expected-display` must match the ratio calculated from the requested data.
The runner independently reads the saved model and embedded datasheet,
requires the same actual ratio, checks native field ownership/precision and
source preservation, and copies the output only after the feature gate passes.
It serializes Office work without selection, activation or mouse control.
The existing official update runs in-process to share the operation lock.

Review the preview referenced by the official report. Routine subsequent data
updates use the ordinary native update route, preserving the converted label.
Do not reconvert a label that is already percentage-only.

## Evidence

Native updates and reopen verified 25/200 → `(13%)`, 30/200 → `(15%)`, and
25/250 → `(10%)`; 12.5% rounds to 13% at zero decimal places. Independent review
passed. The packaged command also converted a changed input showing `25 (13%)`
and updated it to `(15%)`, with actual native ratio 30/200. See
[percent-only-evidence.json](percent-only-evidence.json).
