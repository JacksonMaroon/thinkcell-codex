# Validation for 0.2.1

Validation date: September 15, 2026. This is an experimental Windows desktop release. Version 0.2.1 changes only repository and companion documentation; the recorded native chart checks below are the 0.2.0 core-behavior evidence.

## Native execution

Nineteen changed-data examples completed official JSON generation, exact model/datasheet checks, native PowerPoint save/reopen, source-preservation checks and rendered-preview review:

| Charts | Executed check |
|---|---|
| Four waterfalls | Ascending total 160; backward build-down total 105; subtotal 130 and final total 65; negative-step total 145 |
| Percent Mekko | First segment 36 to 41; first category width/total 124 to 129; relative shares updated |
| Units Mekko | First width 300 to 330 independently of segment height 20 to 25 |
| Line and stacked area | Changed a numeric series point; native model/datasheet/cache checks passed |
| Clustered and stacked columns | Changed one segment; native model/datasheet/cache checks passed |
| 100% stacked column | Changed a segment and its explicit 100%= denominator together; normalized cache matched that denominator |
| Single-axis and dual-axis column/line combinations | Changed one series point; native model/datasheet/cache checks passed |
| Bubble and scatter | Increased one size by 10%; moved one X coordinate by 0.01 |
| Horizontal bar and stacked horizontal bar | Changed one value; handled internal model order separately from exact datasheet row order |
| Pie and doughnut | Changed one absolute pie value; transferred one percentage point between doughnut slices, retaining hole size and input mode |

Ordinary chart examples passed full native-cache parity. **Waterfall and Mekko do not have that generic parity check.** Their Office representations need a different verifier. Those six examples instead passed exact data/model checks, native reopen and visual semantic review. Waterfall equals positions, connector endpoints and group grounding were checked for preservation. Mekko widths were checked in the native model and reviewed visually. This is bounded evidence for these donors, not certification of every topology, label layout, count, or signed-data case.

Native style defaults were loaded through `LoadStyle` and the expected style name was read back through `GetStyleName` after save/reopen. Existing chart formatting came from branded donors. This does not prove retroactive restyling of existing elements.

The `create` command was executed for bar, pie, doughnut, stacked-column, stacked-bar and Gantt donor slides. Gantt dates/tasks/dependencies were not changed. The local companion also completed a fresh Mekko donor copy with source and other-presentation state unchanged. Donor creation preserves example text and layout; a copied example is not a finished client slide.

Successful update runs in this session took approximately 9 to 17 seconds each, excluding donor selection, initial extraction, style loading and visual review. These observations are not a cross-machine benchmark.

## Fixes exercised

- Small legacy BIFF datasheets now get a valid temporary compound-file reader wrapper.
- Template donors retain their file type during staging and are saved natively as PPTX.
- Waterfall `e` inputs are distinct from their calculated numeric model values.
- Mekko targets resolve through genuine scalar-owned shapes rather than assuming a PowerPoint chart object.
- Explicit 100%= denominators are separate from data values; a first series row is not mistaken for a denominator.
- Horizontal chart series are matched by unique label while exact datasheet row order is checked independently.
- Uniform percentage pies/doughnuts retain fractional input and emit official percentage-point JSON.
- Independent review found boolean expected totals and formula-like numeric values slipping through preflight; both are now rejected and covered by regression tests.

## Portable checks and limits

The 26 portable tests pass without Office or proprietary files. Plugin and skill validators pass. The public archive excludes proprietary templates, style files, internal data and local account paths. Original source helpers in the installed presentation skill were checked for unchanged hashes.

```powershell
python -m unittest discover -s tests -v
```

Recorded runtime: Windows, PowerPoint 16.0.20228.20186, installed think-cell 14.0.38.764, Python 3.12.14. The installed file version does not independently prove which DLL an already-running Office process loaded.

Not implemented: arbitrary chart insertion onto a blank slide, changing category/series/point counts, general geometry changes, adding/removing chart annotations or scale breaks, Gantt data changes, percent-input Mekko, mixed absolute/percentage pie inputs, persistent external-link updates, or atomic coauthoring conflict protection. Native reopen is distinct from manually double-clicking every datasheet in the UI. These limits appear in the skill and capability guide.
