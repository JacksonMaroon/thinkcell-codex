# Independent review: COM partial-delete percent labels

## Result

The three changed-data outputs are evidence-backed native percent-only label conversions for the tested scalar. The operation deletes only the absolute `71` field and preserves the surrounding parentheses; the saved output retains a single visible native `a:fld` carrying the relative percentage value.

## Changed-data proof

| Saved output | Numerator | Denominator readback | Expected ratio | Visible native field |
|---|---:|---:|---:|---|
| `com-relative-parentheses-25of200-native.pptx` | 25 | 200 | 12.5% | `13%` |
| `com-relative-parentheses-30of200-native.pptx` | 30 | 200 | 15% | `15%` |
| `com-relative-parentheses-25of250-native.pptx` | 25 | 250 | 10% | `10%` |

The numerator is read from `CVariableSource` 206 (`m_varval`), while the denominator is read from the validated embedded datasheet category extent. The 25/250 case changes the first category extent to 250 while the other categories remain 200, so numerator and denominator vary independently. The 25/200 result rounds 12.5% to 13% under the donor's zero-decimal percentage precision; this is consistent with the 15% and 10% results.

## Native ownership and topology

All three compact traces identify scalar 191 and `CSequenceChartSE` owner 7 with `m_ect val="0"`. The scalar retains both absolute and relative source objects. The absolute source (206) owns the changed raw numerator and has an empty `m_ctextvar`; the relative source (208) owns `CTextVariable` 212, whose format and precision suffix are `%`. The visible shape contains exactly one `a:fld`, with percent format and rendered text matching the relative source. Literal siblings `(` and `)` remain around that field. This is the relative-owned physical field topology; the unbound absolute source is retained as model data and is not displayed.

No additional scalar-label content selector, `m_bMSGraphRendering` change, or owner mode mutation is required by these outputs. The owner and label `reqver`/`endver` metadata remain present, and the traces show `m_bMSGraphRendering val="0"`.

## Integrity checks

Each per-case report is `ALL_GATES_PASS`, has one native chart, and records `native_reopen_pass: true`. The reports record `source_unchanged: true`, `other_presentations_unchanged: true`, `theme_and_notes_preserved: true`, and unchanged non-target OLE/ZIP streams. The visible field text is read from the saved output traces, not supplied as a literal string.

## Remaining limitation

This proves the tested partial-delete mutation and its changed-data behavior for one scalar in a stacked percent-axis chart. It does not establish a general-purpose API command or certify every chart type. The earlier Format Library scan was scoped to the absolute-owned-field/relative-numeric split and should not be read as excluding this relative-owned-field topology.
