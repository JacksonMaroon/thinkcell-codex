# New charts, styles, and chart features

## Create a native chart slide

Choose a real native donor supplied or authorized by the user. Match chart type, category and series capacity, label placement, axes, annotations and dimensions before generation. Copying a full native slide retains think-cell editability and dependencies. This command supports `.pptx` and `.potx` donors:

```powershell
python scripts/thinkcell.py create --input donor.pptx --expected-sha256 HASH --slide-number 1 --output-directory new-chart --execute
```

Add `--style-file company-style.xml` to bind native style defaults. Read the returned output path, inspect that new file, generate its request, update the data, and review the preview. Creation includes the donor slide's existing text and notes: revise ordinary titles and source text to match the user's content using the presentation workflow. Do not treat inherited donor copy as a finished client slide.

The donor determines the type and starting geometry. This is template-based chart creation, not an API that inserts any chart onto an arbitrary blank slide. Use a donor with the required dimensions and point/series counts. Counts cannot be changed through this version's update contract. Gantt/timeline and other native elements can be copied, but their dates, tasks and dependencies cannot be edited through this JSON wrapper.

No proprietary donors or company style files are bundled. A local company profile can reference authorized local files without redistributing them.

## Style support

The native `LoadStyle` API can load an XML style into a master or layout. `load_style.ps1 -InputFile <pptx> -OutputDirectory <new-folder> -StyleFile <xml>` binds the style on a copy, saves/reopens it and verifies `GetStyleName`. `create --style-file` calls this helper automatically.

**Loading a style sets defaults for newly inserted elements. It does not retroactively restyle existing charts.** Use a correctly branded donor to preserve existing colors, fonts, label formats and geometry. For an existing chart requiring restyling, the native think-cell toolbar or a preformatted replacement donor is still needed. Do not recolor PowerPoint caches or patch private model styles.

## Features beyond data

| Feature | Current route |
|---|---|
| Data, category/series labels, bubble size | JSON update, within the checked contract |
| Waterfall equals totals/subtotals and connector topology | Existing donor topology retained; calculated results checked against explicit expectations |
| Mekko width and composition | Specialized absolute-input contracts for percent and units variants |
| Colors, fonts, number formats, label placement, legend, axes | Preserve the donor; optional style defaults for future inserted elements |
| Difference arrows, CAGR arrows, error bars, trendlines, axis breaks, series connectors, reference lines | Choose a donor already containing them; adding/removing/repositioning them is not implemented by this package |
| Gantt dates, task bars, milestones and dependencies | Native donor copy only; specialized editing is not implemented |
| Excel/Tableau links | Persistent or unknown external links rejected by the update route |
| Arbitrary move/resize/type conversion | Native UI or suitable donor required; no generic API route is claimed |

Presence is not correctness: after a data update, review retained annotations, automatic labels, ranges, scale breaks and clipping against the requested result. Feature inventory is slide-level and does not prove each feature belongs to the selected chart or remains meaningful.

Official references: [API](https://www.think-cell.com/en/resources/manual/api), [style files](https://www.think-cell.com/en/resources/manual/style-files), [JSON automation](https://www.think-cell.com/en/resources/manual/jsondataautomation).
