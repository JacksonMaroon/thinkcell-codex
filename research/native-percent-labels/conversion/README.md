# Same-chart absolute-to-percentage conversion

The existing absolute stacked-column fixture now converts to bare dynamic percentage labels while retaining its ordinary absolute axis. This is an experimental native-model/physical-field adapter followed by official JSON regeneration and native save/reopen. It is not a captured toolbar operation or an official think-cell label-content API.

## Verified result

Starting from `../00-native-absolute.pptx`, the packaged `convert-percent` command converts one selected label or batches all nine labels into one regeneration. The original source is unchanged. A new CTextVariable binds each selected scalar's own relative source to one native physical field, with no parentheses, zero-width characters or literal overlay.

| Fixture | Series 1 / 2024 | Series 3 / 2024 | Category total | Target label |
| --- | ---: | ---: | ---: | ---: |
| 00-one-native-percent.pptx | 3.5 | 14.9 | 23.5 | 63% |
| 01-all-native-percent.pptx | 3.5 | 14.9 | 23.5 | 63% |
| 02-numerator.pptx | 3.5 | 25 | 33.6 | 74% |
| 03-denominator.pptx | 13.5 | 25 | 43.6 | 57% |

Every all-label output passes the nine-field retention contract, exact embedded-datasheet/model parity, native-cache consistency, source preservation and native save/reopen. Relative-source GUIDs and semantic label positions persist; displayed percentages equal the new ratios. Visual review confirmed white labels, correct percentages and no clipping. Other category values remain unchanged.

The independent geometry audit finds unchanged native owner/table/sequence/plot-area identity, axis, type, orientation, gap and plot coordinates. For all-label conversion only, think-cell recomputes the outer chart-cache frame and compensates its manual-layout cache. Exact-color component bounds match for all three colors across all three categories. The final gate retains exact native plot/axis checks and permits this reported cache-frame adjustment only when all nine native fields pass. Single-label conversion retains the original frame too.

`results.json` records hashes, label bindings, ratios and the independent geometry findings. The four PPTX files are synthetic public regression fixtures, including a real field-backed fixture for the original selector.

## Reproduce

```powershell
python plugins/thinkcell-codex/skills/thinkcell-edit/scripts/thinkcell.py convert-percent --input research/native-percent-labels/00-native-absolute.pptx --expected-sha256 6F4107910C797192E24CCA9B215A37C46E9C0674C318EE35F730A7EF41C96968 --all-labels --output-directory C:/Temp/native-percent-converted
python -m unittest discover -s plugins/thinkcell-codex/tests -q
```

Use a fresh directory. Run ordinary `inspect --request-out` and `update` on the returned `converted.pptx` for later data changes; the new relative-field contract runs automatically before output delivery. The current portable suite passes **91 tests**.

## Profile and failed alternatives

Tested on Windows 11 Enterprise 10.0.26200, PowerPoint 16.0.20326.20140, think-cell 14.0.38.772 and native model 38775. The adapter accepts the explicitly checked model 38775/38772 schema, one ordinary consistent absolute stacked-column chart, internal data, 3x3 dimensions, nonnegative values, positive calculated totals, no axis breaks and the tested bg1 segment text styling. Other families, custom denominators, negative values and wrapper-label migration remain outside this certification.

PowerPoint's `DataLabel.ShowPercentage` rejected this stacked chart with “This property is not valid for the current chart type.” No presentation outside the experiment changed. The successful route instead establishes a coherent native relative-field binding and verifies its changed-data behavior. The existing working direct-native percentage route remains available for charts already using that representation.
