# Validation for 0.3.0

Validation date: September 15, 2026. This is an experimental Windows desktop release.

## Native data execution retained from 0.2.0

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

## New finished-slide workflow

The 0.3.0 route follows the previously tested pie coordinate approach: update exclusive native coordinate anchors on a copy, regenerate through official JSON, retain the complete donor chart and model, and copy the target's ordinary content onto that slide through native PowerPoint. No raw think-cell shape-range transplant is used.

Seventeen one-chart donor cases passed placement/composition checks: pie; compact bar and stacked bar; clustered, stacked and 100% columns; line; area; single-axis and dual-axis column/line; four waterfall topologies; percent and units Mekko; and bubble. All returned a separate native one-slide PPTX after exact data/model checks, native save/reopen, full visible chart-footprint containment, source hash checks, original-content signatures, and pixel equality outside the insertion frame. Their previews were reviewed for fit. The matrix uses illustrative technical fixtures; its generic titles do not substantiate a business conclusion for each donor's data.

A separate coherent example supplied five services with shares 33%, 27%, 20%, 15%, and 5% and a finished slide saying the first two contribute 60%. The agent selected a pie from this context and the planner filled the empty region. A second run supplied only the pie choice and horizontal placement (left 400, width 480 points); it inferred the remaining coordinates. Both retained the existing text, table, SVG picture, native group, hidden/off-slide object, slide number 7, note text and note font runs. The source-native render and every pixel outside the insertion frame matched. Updating the inserted chart again to 33/27/20/14/6 retained its geometry and passed the normal native data route.

The rich example and final compact/stacked-bar runs compare note font runs to the original source snapshot, including font, size, bold, italic, underline and resolved color. Earlier matrix runs checked note text and imported-target font runs after reopen. Paragraph spacing, indentation, bullets and notes-page layout are not independently certified. Existing content remains native; the rendered comparison is an additional preservation check, not the only one.

### Native geometry findings

- Pie plot dimensions remain square, with extra space for outside labels.
- Detached legends have separate exclusive anchors and are repositioned when supported.
- Horizontal sequence charts also require their outer anchors. The compact bar then contracted its derived plot's right edge by 3.125 points during native reflow. The adapter permits only inward right/bottom changes up to 4 points for this family, records the adjustment, preserves position, and still requires the entire visible bundle to fit the requested outer region.
- The tested scatter donor's decorative background group stayed at its original position. The footprint gate rejected it and withheld a final output. Scatter data updates still pass; this decorated donor is not certified for composition. A suitable alternative donor needs its own native test. The workflow never silently removes the decoration.
- The local doughnut donor contains two charts, so it is rejected by this one-chart composition route. Its prior copy/update support remains intact. Gantt is still native-copy-only.
- PowerPoint can briefly create hidden unnamed presentations. Full before/after snapshots are retained; the comparison excludes only saved, hidden, unnamed transients and waits up to five seconds after an actual mismatch. Visible, named or unsaved work remains in the comparison. No application-wide quit is used.

This is bounded donor evidence, not support for every annotation, chart topology or slide feature. Explicit user chart/placement instructions take precedence. Inference is performed from the actual slide and supplied data; portable geometry code does not choose chart semantics from keywords. Automatic reports still require visual and semantic review.

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

The 64 portable tests pass without Office or proprietary files. Plugin and skill validators pass. The public archive excludes proprietary templates, style files, internal data and local account paths. Original source helpers in the installed presentation skill were checked for unchanged hashes.

```powershell
python -m unittest discover -s tests -v
```

Recorded runtime: Windows, PowerPoint 16.0.20228.20186, installed think-cell 14.0.38.764, Python 3.12.14. The installed file version does not independently prove which DLL an already-running Office process loaded.

Not implemented: arbitrary chart generation from scratch, changing category/series/point counts, geometry outside the bounded donor route, adding/removing chart annotations or scale breaks, Gantt data changes, percent-input Mekko, mixed absolute/percentage pie inputs, persistent external-link updates, or atomic coauthoring conflict protection. Native reopen is distinct from manually double-clicking every datasheet in the UI. These limits appear in the skill and capability guide.
