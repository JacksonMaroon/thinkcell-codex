# Continue on Windows

This branch includes conservative read-only discovery and Windows fixture validation. It does not add percentage conversion support. See [Windows results](windows/README.md).

## First: reproduce the evidence

From the repository root in PowerShell:

```powershell
py -m venv .venv-label-lab
.\.venv-label-lab\Scripts\python.exe -m pip install olefile lxml
.\.venv-label-lab\Scripts\python.exe research/native-percent-labels/inspect_fixtures.py
.\.venv-label-lab\Scripts\python.exe research/native-percent-labels/verify_fixtures.py
```

These commands read the PPTX files and regenerate the JSON/XML evidence alongside them. They do not require Office. `verify_fixtures.py` checks this exact corpus and the saved discovery snapshot, not a fresh run of the current production selector. Do not use Python optimization (`-O`), which disables its assertions.

- [x] Record Windows, PowerPoint, think-cell, and Python versions.
- [x] Run the two checks above and record the result.
- [x] Copy decks 02 through 05 to a separate working folder before opening them in PowerPoint with think-cell enabled.
- [x] Confirm the Series 3 / 2024 label reads 50%, 67%, 50%, 50% respectively.
- [ ] In a working copy of 02, change Series 3 from 100 to 200 through the native datasheet. Confirm 67%. Then change Series 1 from 25 to 125. Confirm 50%.
- [ ] Save, close, reopen, and verify again. Preserve the new Windows decks separately from the Mac originals.

## Then: extend read-only discovery

Relevant code: `plugins/thinkcell-codex/skills/thinkcell-edit/scripts/percent_labels/discover_percent_semantics.py`.

The saved `existing-discovery.json` contains nine labels with null relative text variables and empty physical-shape lists. `existing-selector-result.json` records the reviewed source hash and selector failure. Compare that hash to your checkout before treating the snapshot as current. The recorded selector call used numeric category `2024.0`; passing a string category through the CLI can fail earlier for a different reason.

- [x] Add explicit discovery for direct-rendered native labels, reporting direct precision, owning chart, and category/series identity.
- [x] Keep field-backed selection guards intact. A missing relative text variable is not sufficient to classify a label as static.
- [x] Check native ownership, axis/data semantics, and numerical agreement together. Neither the percent suffix nor the MSGraph-rendering flag proves dynamic percentages on its own.
- [ ] Test the six fixtures and an existing field-backed fixture. Do not assume the reverse series order in this corpus applies to every chart.
- [x] Keep discovery support separate from permission to mutate a label.

## Finally: investigate conversion

- [ ] On one ordinary stacked chart, save a before deck, use the native floating toolbar to change label content, then save an after deck. Change only the target setting.
- [ ] Repeat numerator and denominator edits and save/reopen checks on the converted chart.
- [ ] Compare native model, chart XML, ownership, and rendered output. Preserve both decks and observations.
- [x] Test official JSON / `.ppttc` regeneration independently from datasheet edits.
- [ ] Only design a production conversion adapter once this evidence supports it; do not copy cache formatting or remove guards to bypass the unsupported representation.

The Mac experiment established dynamic default percentage labels. The Windows continuation validated native reopen and JSON regeneration on these fixtures. Absolute-to-percent conversion and Windows datasheet UI edits remain outstanding. See [README.md](README.md) for the observed structures and exact limitations.

## Suggested prompt for your Windows Codex task

> Read research/native-percent-labels/README.md and WINDOWS-HANDOFF.md. Reproduce the read-only fixture checks, compare the saved discovery source hash with current code, then implement direct-rendered percentage-label discovery while preserving existing field-backed guards. Use the committed fixtures for regression coverage. Report what was validated on Windows separately from the saved Mac evidence. Do not implement conversion or claim native UI validation unless you can actually verify it.

## Conversion handoff update

The [experimental same-chart relative-field adapter](conversion/README.md) now passes single-label and all-nine conversion, native save/reopen, numerator/denominator updates, exact data, source binding and geometry gates. Real field-backed regression fixtures are committed. This is a coherent native-model adapter, not a captured toolbar operation; the UI-specific boxes above remain open for that independent route.
