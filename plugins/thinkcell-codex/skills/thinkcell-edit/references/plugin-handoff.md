# Connected PowerPoint handoff

This optional route needs an independently connected PowerPoint plugin exposing `run_officejs`, `list_slides` and `edit_slide_ooxml`. Discover its current schemas and bind the exact presentation session. These are not tools supplied by this package. If unavailable, produce a verified local slide and say that live return was not performed.

## Export

Record the exact slide ID, current index, displayed slide number and relevant notes/theme/content. Export the whole native slide through Office.js:

```javascript
const result = ctx.presentation.slides.getItem(slideId).exportAsBase64();
await ctx.sync();
return result;
```

In the tested connector sandbox, the serialized result used `m_value`; returning `result.value` failed. Read the actual response structure. Decode its base64 into a local PPTX without printing the payload into the conversation. Save the export as an immutable source snapshot, then use `thinkcell.py inspect` and `update`. Preserve the displayed page number when judging render parity.

## Return

After exact data checks, native reopen and visual review:

```powershell
python scripts/thinkcell_no_click/implementation/prepare_plugin_return.py --candidate "D:/work/updated.pptx" --expected-sha256 "<verified-hash>" --slide-id "<bound-plugin-id>" --slide-index 2 --output "D:/work/return-plan.json"
```

The helper splits the complete PPTX into auxiliary VFS chunks to avoid the connector's script-size limit. Execute the `staging_args` through the bound session's `edit_slide_ooxml`. Staging does not contain a final replacement operation.

Immediately before replacement, fetch `list_slides` again and re-export the target. Compare its actual content, notes, data, formatting and theme dependencies with the original snapshot. If the target changed, regenerate from the fresh export and restage. Unrelated-slide changes do not require restarting an unchanged target. Ask about a conflict only when the user's intended content is ambiguous.

Save the fresh list object with `slides` entries containing `id` and `slideIndex`, then finalize:

```powershell
python scripts/thinkcell_no_click/implementation/prepare_plugin_return.py --plan "D:/work/return-plan.json" --slides-json "D:/work/fresh-slides.json" --output "D:/work/final-call.json"
```

Execute only the finalized `replacement_args`, promptly in the same session. The helper re-resolves identity to the current index and rejects missing/ambiguous targets. It does not itself fetch the source or protect against simultaneous server edits.

Read the replacement's new slide ID from the tool response, export it, and check it using `thinkcell.py verify --prepared <work/generated_thinkcell_work/prepared.pptx> --input <returned-slide.pptx> --automation-name <name-from-report> --data-json <request.json>`. Inspect its render too. Preserve relevant comments, sections, custom shows and internal links: the earlier local handoff test did not establish their preservation. If those dependencies are present and preservation cannot be verified, deliver a local candidate until a supported return method is available. Shared-file writes require the user's authorization and fresh saved readback; this is not an atomic coauthoring lock.

Never replace only the visible chart cache, chart-tagged shapes or model stream. Transport the complete native slide. After an uncertain write, inspect current state before retrying.
