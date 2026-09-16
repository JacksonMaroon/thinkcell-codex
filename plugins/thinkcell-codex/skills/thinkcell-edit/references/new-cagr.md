# Native dynamic CAGR insertion

This experimental route adds a second native CAGR annotation to the verified
single-series, eight-category column donor. It copies the existing annotation's
native graph, arrow shapes and label structure, gives the new objects separate
identities, and binds the new label to the new native CAGR calculation.

It does not require mouse clicks. The result remains think-cell content, with
an embedded datasheet and a real numeric percentage field. The original donor
annotation remains present. If its label was manually overridden, this route
preserves that override; it does not silently convert the original label.

## Verified calculation

| Source | Endpoint | Dates | New label |
|---|---|---|---|
| 371.1 | 600 | 2019-01-01 to 2025-01-01 | 8.3% |
| 371.1 | 700 | 2019-01-01 to 2025-01-01 | 11.2% |
| 371.1 | 700 | 2019-01-01 to 2031-01-01 | 5.4% |

These controls passed official JSON regeneration, native save/reopen and
independent inspection of the saved model, datasheet, linked field and preview.
The label has one decimal place. Category dates must be genuine date cells;
year-looking strings do not establish the elapsed period.

## Scope and verification

Use a user-supplied donor that matches the tested profile. Work from a copy and
verify its current SHA-256 before preparation.

The bounded profile is model 38764, one slide, one ordinary column chart, one
series, eight categories, an internal datasheet and an existing native CAGR
with the verified ellipse-label and two-line structure. The new annotation
uses the same endpoint category indices, 1 and 7 (zero based), as its source.
The initial date contract uses January 1 dates and positive values. Broader
chart families, different endpoint selection and fractional-year periods
require separate native evidence.

Keep input, working candidates and delivered output separate. Verify the source
hash before preparation. Withhold the delivered file until all of these pass:

- Exact requested values and typed dates in both native model and datasheet.
- Current chart consistency and no fallback to a last-consistent datasource.
- Separate new graph, native arrows, label and linked percentage field.
- Correct endpoint ownership, numeric CAGR, precision and visible percentage.
- Exact field-format binding after native normalization.
- Original annotation identity, text and style retained, source unchanged and
  unrelated open presentations untouched.
- Native save/reopen and visual review of annotation placement.

## Commands

From the skill directory, prepare a candidate with
`scripts/experimental_cagr/insert_native_cagr.py`. It requires `--input`,
`--expected-sha256`, `--workdir`, `--output` and `--report`. The work directory
must already exist; output and report must be new direct children of it.
`PREPARED_NATIVE_REGENERATION_REQUIRED` is not a deliverable result.

Run `scripts/experimental_cagr/native_cagr_gate.py` with the prepared candidate,
its inspected hash, `--prepared-report`, a complete `--data` JSON, `--ppttc`
pointing to the installed executable, and fresh output/report paths under the
work directory. Its success status is `NATIVE_DYNAMIC_CAGR_GATES_PASS`.
For later changes, use the previous output as input and its verified
`--previous-report` in a fresh work directory.

The complete data JSON has exactly these keys:

```json
{
  "name": "CAGR_CONTROL",
  "dates": ["2018-01-01", "2019-01-01", "2020-01-01", "2021-01-01", "2022-01-01", "2023-01-01", "2024-01-01", "2025-01-01"],
  "values": [360.6, 371.1, 381.9, 393.0, 404.4, 416.2, 428.3, 600]
}
```

Dates must be strictly increasing January 1 dates; values must be finite and
positive. Keep every requested category and value explicit. Review the native
preview before delivering the output.

Native regeneration can move arrows and labels to accommodate the chart. Check
the final geometry and readable placement rather than forcing stale donor
coordinates. A working feature fixture is not a finished slide: surrounding
template placeholders and date-label fit still need normal slide assembly.

## Why earlier attempts failed

The failed foreign-label clone mismatched its physical PowerPoint shape and
native model. The chart became inconsistent and skipped the entire data update.
That was not evidence that think-cell cannot recalculate a new CAGR label.
The successful route preserves the native label structure and updates its
numeric field coherently. See [native-feature-findings.md](native-feature-findings.md).

See [expanded native controls](expanded-features.md) for secondary axes, line colors, semantic scalar formatting and additional CAGR endpoints/date controls.
