# Command reference

Commands below run from this skill directory with a configured Python interpreter. Use absolute file paths when working across directories. Put outputs in the user's task folder, outside the installed skill. Commands refuse existing outputs.

## Inspect and prepare a request

```powershell
python scripts/thinkcell.py doctor
python scripts/thinkcell.py inspect --input "D:/work/slide.pptx"
python scripts/thinkcell.py inspect --input "D:/work/slide.pptx" --slide-number 1 --shape-id 21 --data --request-out "D:/work/update.json"
```

Use the actual shape ID from inspection. A think-cell tag (`--shape-tag`) is an alternative; if multiple selectors are supplied, all must agree. `--slide-id` selects a stable native slide ID instead of a one-based `--slide-number`. IDs may change on extraction or plugin return, so re-inspect.

The request starts with current data. Change it using [data-contract.md](data-contract.md). Do not treat example IDs or numbers as the user's chart.

## Update

```powershell
python scripts/thinkcell.py update --input "D:/work/slide.pptx" --expected-sha256 "<hash-from-inspect>" --slide-number 1 --shape-id 21 --data-json "D:/work/update.json" --output "D:/work/updated.pptx" --execute
```

Omit `--execute` for a read-only preflight. With `--execute`, the command automatically uses the experimental namer when needed, generates the chart, checks its model and datasheet, applies the reported native-cache or specialized semantic checks, saves and reopens a native output, checks it again, and writes `updated.pptx`, `updated.report.json` and `updated.work/`. Inspect the preview listed in the report before delivery. A report status ending in `VISUAL_REVIEW_REQUIRED` means automated checks passed but a person or image-capable agent still needs to inspect the preview.

Add `--named-only` to require an existing automation name. Add `--report <new.json>` for a custom report path. Outputs and work folders must be new. Failed runs retain diagnostic files; never blindly repeat the same operation after a timeout.

This wrapper performs native verification once per update and reuses the same generated copy for the final exact-data check. It does not certify a full family on every invocation.

## Extract a saved slide

Prefer a connected plugin export when the deck has unsaved edits. For a saved local file:

```powershell
& "$env:SystemRoot/System32/WindowsPowerShell/v1.0/powershell.exe" -NoProfile -ExecutionPolicy RemoteSigned -File scripts/extract_slide.ps1 -InputFile "D:/work/deck.pptx" -SlideNumber 3 -OutputDirectory "D:/work/extracted"
```

This opens only a copy, removes other slides in that copy, retains the original displayed page number, and renders before/after images. Compare those images and inspect exact chart data before updating. Cross-slide links may no longer make sense in a standalone slide. A local extraction does not replace any slide in the source deck.

## Lower-level helpers

Use `scripts/thinkcell_no_click/implementation/prepare_plugin_return.py` only for the connected handoff described in [plugin-handoff.md](plugin-handoff.md). The naming and JSON helper modules are implementation details. Prefer the wrapper so data checks and native verification are not accidentally skipped.
