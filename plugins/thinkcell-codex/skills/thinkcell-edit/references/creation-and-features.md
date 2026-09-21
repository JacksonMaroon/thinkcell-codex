# New charts, styles, and chart features

## Create a native chart slide

Choose a real native donor supplied or authorized by the user. Match chart type, category and series capacity, label placement, axes, annotations and dimensions before generation. Copying a full native slide retains think-cell editability and dependencies. This command supports `.pptx` and `.potx` donors:

```powershell
python scripts/thinkcell.py create --input donor.pptx --expected-sha256 HASH --slide-number 1 --output-directory new-chart --execute
```

Use `--native-percent --data-json request.json` to create, populate and verify a single native percentage donor chart in one operation. See [native percentage labels](native-percentage-labels.md). The data request uses the ordinary canonical matrix and expected-model contract.

Add `--style-file company-style.xml` to bind native style defaults. Read the returned output path, inspect that new file, generate its request, update the data, and review the preview. Creation includes the donor slide's existing text and notes: revise ordinary titles and source text to match the user's content using the presentation workflow. Do not treat inherited donor copy as a finished client slide.

The donor determines the type and starting geometry. For a chart added to a finished chartless slide, use the separate [existing-slide composition route](existing-slide.md), which includes experimental movement and resizing. Use a donor with the required point/series counts for the single-target route. The separate [multi-chart route](multi-chart.md) passed a four-series/three-category ordinary sequence update and a fixed-topology waterfall with an ordinary sibling; other count changes or waterfall topologies still require equivalent readback and visual evidence. Gantt/timeline and other native elements can be copied, but their dates, tasks and dependencies cannot be edited through this JSON wrapper.

Use only donors and style files that the user supplies or is authorized to use. This package does not bundle proprietary templates or branded style references.

## Style support

The native `LoadStyle` API can load an XML style into a master or layout. `load_style.ps1 -InputFile <pptx> -OutputDirectory <new-folder> -StyleFile <xml>` binds the style on a copy, saves/reopens it and verifies `GetStyleName`. `create --style-file` calls this helper automatically.

**Loading a style sets defaults for newly inserted elements. It does not retroactively restyle existing charts.** Use a correctly branded donor to preserve existing colors, fonts, label formats and geometry. For supported ordinary bar/column and stacked 100%-column donors, use the verified datasheet-fill pipeline for series colors. Other styling uses a compatible donor. Do not recolor PowerPoint caches.

Official JSON accepts typed cell fills. The earlier theme-color failure was resolved by enabling the existing native datasheet-fill setting on a guarded copy. Four exact RGB fills survived a second update with changed totals. Use [reliable-features.md](reliable-features.md) and its tested topology limits.

### Native feature route

`feature_api.py` is an experimental route for a one-slide local copy. It requires every native chart to be explicitly targeted and a non-null 1–72 point font size. It sizes existing tagged native text fields, regenerates a complete all-chart baseline, then requires native verification and retention readback before releasing the output.

The identity-durability canary passed a 9-point fonts-only multi-chart run and a subsequent changed-data update. This is the supported formatting route.

A native series fill survived an unchanged-data refresh but reverted after the data changed. That first refresh was insufficient evidence of durability. A label-text edit also removed the label during regeneration. Direct cache recoloring and literal label replacement remain unsupported. The separate datasheet-fill route now resolves series colors for certified donors.

## Features beyond data

| Feature | Current route |
|---|---|
| Data, category/series labels, bubble size | JSON update, within the checked contract |
| Waterfall equals totals/subtotals and connector topology | Existing donor topology retained; calculated results checked against explicit expectations |
| Mekko width and composition | Specialized absolute-input contracts for percent and units variants |
| Existing native text size | Experimental `feature_api.py` route with exact source hash, every chart explicit, complete data baseline, official regeneration, native verification and retention readback |
| Primary absolute axis maximum and major unit | Experimental `axis_range.py` route only for one selected exclusive zero-minimum axis; no axis break, secondary axis, negative values, or shared series group |
| Colors, number formats, label placement, legend, axes | Preserve the donor; optional style defaults for future inserted elements |
| Series RGB fill requests | Verified for ordinary horizontal stacked-bar and stacked 100%-column donors through the datasheet-fill pipeline; point overrides await native certification |
| Existing native axis-break position | Use `axis_break_position.py` for the verified one-break profile; see [broken-axis.md](broken-axis.md). |
| Existing numeric total decimal places | Use `total_label_precision` with its one-category contract and exact dynamic-field verification; see [total-label-precision.md](total-label-precision.md). |
| New CAGR arrows and axis breaks | Bounded insertion routes are verified: [native dynamic CAGR](new-cagr.md) and [new axis break](new-axis-break.md). Respect each route's exact profile and native gates. |
| Difference arrows, error bars, trendlines, series connectors, reference lines | Prefer an authorized donor containing the feature. Other new insertion is not yet verified for delivery; user-authorized extension work follows [native-feature-research.md](native-feature-research.md). |
| Label text insertion or editing | Not supported; edited native label text did not survive official regeneration |
| Gantt dates, task bars, milestones and dependencies | Native donor copy only; specialized editing is not implemented |
| Excel/Tableau links | Persistent or unknown external links rejected by the update route |
| Move/resize a one-chart native donor | Experimental exclusive plot/legend coordinate adapter, official regeneration and native readback in the existing-slide route |
| Type conversion, shared/grouped constraints, general element insertion | Not implemented; choose a compatible donor or a separately supported native route |

Presence is not correctness: after a data update, review retained annotations, automatic labels, ranges, scale breaks and clipping against the requested result. Feature inventory is slide-level and does not prove each feature belongs to the selected chart or remains meaningful.

Official references: [API](https://www.think-cell.com/en/resources/manual/api), [style files](https://www.think-cell.com/en/resources/manual/style-files), [JSON automation](https://www.think-cell.com/en/resources/manual/jsondataautomation).
