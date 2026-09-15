# Add a chart to a finished slide

Use this experimental route when the target contains **no think-cell elements** and the user wants a new slide copy with a native chart added. It retains a complete donor chart and its native model, copies the target's ordinary content natively, applies the target layout, and preserves the original. It does not transplant a raw think-cell shape range or merge two think-cell models.

## Follow the guidance that exists

Honor explicit chart type, data, labels, units, features and placement. Infer only missing choices from the actual slide, neighboring context and available data. Do not classify charts by title keywords alone, substitute a convenient donor for an explicit chart request, or manufacture values. A donor must already contain the requested features and compatible category/series capacity.

When chart type is unspecified, choose the exhibit that answers the slide's question: comparisons often suit bars, changes over ordered time suit lines, contributions to a total change suit waterfalls, and composition with meaningful category widths can suit Mekko. A pie can suit a small number of parts of one whole. Match the user's purpose rather than applying this as a fixed lookup table. Use their controlling template and a matching native donor. Company assets stay local.

Choose available space after viewing the slide and reading its native geometry. Preserve title, commentary, source/footer rails, tables, pictures, grouped shapes and notes. The frame is the available **outer region including labels and legend**, not just the plot. Pies retain a square plot. If explicit guidance overlaps protected content, explain the specific conflict; do not silently move existing objects or shrink their text. Missing material data is a reason to ask, while routine chart and layout choices are for the agent to resolve.

## Execute

1. Extract the selected target and donor as separate one-slide files with `thinkcell.py create` or `extract_slide.ps1`. These are working copies. Inspect the source identities again after extraction. The output of this route is a distinct one-slide PPTX; returning it to a live deck is a separate whole-slide handoff.
2. Inspect the target natively:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File scripts/inspect_layout.ps1 -InputFile target.pptx -ExpectedSha256 HASH -SlideNumber 1 -OutputDirectory target-layout
   ```

   Review `slide.png`, inherited layout/master shapes and off-canvas objects. Add occupied inherited regions to `reserved` in the returned context JSON. Keep its source identity, dimensions, shape records and preview hash intact. See [placement.md](placement.md).
3. Write a brief with `requested_chart_type` when specified, otherwise `inferred_chart_type`, and a short `rationale`. Optional `frame` members and `side` express the actual placement guidance. Fill unspecified coordinates automatically. Supply a real canonical data request derived from the selected donor, as in [data-contract.md](data-contract.md).
4. Execute the authorized copy operation:

   ```powershell
   python scripts/insert_chart.py --target target.pptx --target-sha256 HASH --donor donor.pptx --donor-sha256 HASH --context target-layout/layout.json --brief brief.json --data-json data.json --output-directory result --execute
   ```

   Omit `--execute` for read-only preflight. The default experimental namer remains enabled. A previously unnamed donor first receives the normal naming and native generation pass. Named donors skip that extra pass.
5. Review the final `composition/after.png` against the slide's message and data. The command checks native plot coordinates, the complete visible chart footprint, exact datasheet/model values, native reopen, original shape signatures, original note text and font runs, and pixel equality outside the insertion region. Note font checks cover font name, size, bold, italic, underline and color; paragraph bullets, indents and spacing are not independently checked. It withholds `slide-with-chart.pptx` after a failed check. A passing automated report still requires review of internal chart labels and feature meaning, especially waterfall/Mekko semantics.

Example brief when the user specified only a chart type:

```json
{"requested_chart_type":"line","rationale":"The user requested a line chart for the supplied monthly series."}
```

Example with partial placement guidance:

```json
{"inferred_chart_type":"column-clustered","rationale":"Compare the supplied categories across two series.","frame":{"left":400,"width":480}}
```

The planner accepts semantic chart labels, but execution checks them against the donor. Current execution labels are `pie`, `doughnut`, `bar`, `bar-stacked`, `bar-100`, `column-clustered`, `column-stacked`, `column-100`, `line`, `area`, `column-line`, `column-line-dual-axis`, `scatter`, `bubble`, `waterfall`, `mekko-percent`, and `mekko-units`. For waterfall, inspect the donor's actual equals cells and connector structure against any requested subtype; the subtype is a topology requirement, not a type conversion. An unrecognized subtype label is rejected rather than silently mapped to another donor.

## Geometry mechanism and limits

The adapter extends the tested pie approach: on a separate copy it changes only exclusive plot-coordinate grids, the outer coordinate grids required by horizontal sequence charts, and, when present, the one detached legend's exclusive coordinate grids. It proves that all other model content, embedded data streams and package parts remain unchanged, then regenerates through official think-cell JSON and reads back the native result. This is an experimental compatibility adapter, not an official geometry API.

Supported donors have one chart, one carrier, current native model structure, no grouped chart, no shared layout constraints, and at most one detached legend. Multiple-chart donors, Gantt, existing think-cell target slides, arbitrary new annotations, and persistent or unknown external links are outside this route. The chart's category/series/point counts remain bounded by the existing data contract. Other chart-owned annotations must survive regeneration and fit the final frame; no claim is made that arbitrary features can be added or edited.

Target and donor slide dimensions and theme colors/fonts/effects must match. The target route currently rejects custom slide backgrounds, animations/transitions, comments, slide links, OLE objects and custom note objects because their preservation has not been implemented here. An empty ordinary slide currently uses the donor-creation route. It does not flatten unsupported content or silently remove it. A failed geometry or preservation gate means inspect the candidate and choose a suitable donor/frame before a distinct retry.

Native PowerPoint calls and clipboard use are serialized. Close only task-owned presentations. A timeout records the running helper PID and retains the work; inspect it before retrying. This route does not require desktop clicks or change another presentation skill.

The available outer frame remains binding. Native horizontal sequence charts may contract the derived plot's right/bottom edges by up to 4 points during label reflow; the report records the signed adjustment. Left/top movement and outward growth retain the 1.01-point tolerance, and the complete visible bundle must still fit. This allowance does not apply to other families.
