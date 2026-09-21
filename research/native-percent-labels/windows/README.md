# Windows validation, 2026-09-21

This run validates the existing synthetic percentage chart. It does not establish absolute-to-percentage conversion or general chart-family support.

Environment: Windows 11 Enterprise 10.0.26200, PowerPoint 16.0.20326.20140, think-cell 14.0.38.772, Python 3.12.14. The Mac fixture model version is 38775; this Windows build successfully reopened it.

Both offline fixture inspection and verification scripts passed. Separate working copies of fixtures 02 through 05 were opened, saved, reopened and rendered using `native_verify_scoped.ps1`, under the serialized Office lock. All four passed native reopen, source preservation and preservation of other open presentations. Visual review confirmed Series 3 / 2024 labels of 50%, 67%, 50% and 50%.

The official JSON workflow was tested separately, using `thinkcell.py update --execute` and its `.ppttc` regeneration route. Each request came from `thinkcell.baseline_request`, with exact source SHA-256, slide and shape-tag targeting. Starting from the Windows-saved copy of 02:

| Step | Series 1 / 2024 | Series 3 / 2024 | Rendered target |
| --- | ---: | ---: | ---: |
| Unchanged data | 25 | 100 | 50% |
| Numerator change | 25 | 200 | 67% |
| Denominator change | 125 | 200 | 50% |

Each step used the preceding verified output. All three passed exact model-data checks, native/cache parity, source preservation and native save/reopen. Visual inspection of the three committed previews confirmed the target values, unchanged sibling-category labels (82% and 72%) and no clipping. `results.json` records output hashes and scoped outcomes. Raw reports and binary working copies stay local because the reports include machine paths and unrelated open-presentation identities.

## Still outstanding

- Windows native **datasheet UI** edits were not performed. JSON regeneration is independent evidence and does not substitute for that route.
- A same-chart before/after pair made through the native **Label Content toolbar** is still missing. This session exposes browser computer control only, so native toolbar interaction cannot be verified here.
- No production conversion adapter is added. Existing field-backed mutation guards remain in place; read-only direct-label recognition does not grant permission to mutate those labels.

These are narrow fixture/build results, not proof for arbitrary formatting, combined labels, custom denominators or other chart families.

## Reproduce on Windows

From the repository root, with the public plugin dependencies installed:

```powershell
python research/native-percent-labels/run_windows.py --output-dir C:/Temp/native-percent-new-run
python -m unittest discover -s plugins/thinkcell-codex/tests -q
```

Use a new output directory. Office writes run sequentially under the existing lock. Raw reports can name other open presentations; review and sanitize before publishing. Review all seven previews after the run.

The public portable suite passes 72 tests, including eight new discovery tests. Field-backed selection has synthetic guard coverage only; a real distributable field-backed PPTX regression fixture remains outstanding.
