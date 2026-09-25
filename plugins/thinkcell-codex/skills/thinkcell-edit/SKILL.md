---
name: thinkcell-edit
description: Create and edit native think-cell charts from supplied donors, including chart data, formatting, placement and bounded feature investigation. Use for think-cell work, not ordinary PowerPoint charts or general slide design.
---

# Create and edit think-cell charts

Use the bundled scripts to update existing native charts without flattening them. This Windows desktop skill requires PowerPoint, licensed think-cell and Python. It has no account connector or background service. Preserve the user's template, notes, scope and existing authorization.

For ordinary presentation creation, slide writing, and non-think-cell PowerPoint edits, use the separate PowerPoint for Codex companion. Keep think-cell chart operations in this skill.

## Choose the route

For a chart added to a finished slide that contains no think-cell content, read [existing-slide.md](references/existing-slide.md). Follow explicit user guidance and infer only the missing chart/layout choices from the slide and supplied data. The experimental route uses native donor composition, exclusive plot/legend anchors and preservation checks. It creates a new one-slide copy.

For a new chart, use `create` to copy an authorized native donor slide matching the chart type, layout and desired features. Optionally load a user-supplied style file for new-element defaults. Then follow the data-update path below. Read [creation-and-features.md](references/creation-and-features.md) for creation, styling, or features beyond data. Never equate a slide clone with generating an arbitrary chart from scratch.

For clean dynamic percentage labels, prefer an existing verified native percentage chart or a compatible native percentage donor. `update` automatically checks direct percentage-label preservation after native regeneration. Use `create --native-percent --data-json <request.json>` to copy, populate and verify a single-chart donor in one operation, or `update --require-native-percent` when native percentages are mandatory. Follow [native-percentage-labels.md](references/native-percentage-labels.md). Do not apply the field-wrapper workaround to a chart already using verified direct native percentage labels.

For an existing ordinary absolute stacked-column chart matching the verified 3x3 profile, `convert-percent` converts one selected segment or all nine labels to bare native relative fields while retaining the absolute axis and native plot. This is preferred over the invisible-wrapper workaround for eligible charts. The command executes on a new local copy and requires native regeneration/reopen and field-binding checks. Later `update` calls automatically verify these converted fields too. See the conversion profile and limits in [native-percentage-labels.md](references/native-percentage-labels.md).

For formatting, axes, annotations, Gantt taskbars/milestones or existing Excel-link rebinding, use the relevant section of [expanded native controls](references/expanded-features.md) or [reliable features](references/reliable-features.md). These use their own adapters and checks, not necessarily the JSON data-update path below. Inspection-only requests finish with the requested findings.

## Ordinary JSON data updates

1. Resolve the exact slide and chart from the user's request. Use `scripts/thinkcell.py inspect --input <file>` for compact chart identities and the source SHA-256. Add `--data` only when the values are needed. Do not load the implementation into context unless diagnosing a failure.
2. Obtain a whole-slide copy. A one-slide local PPTX can be used directly. For a connected PowerPoint deck, read [plugin-handoff.md](references/plugin-handoff.md). For a saved multi-slide file, use `scripts/extract_slide.ps1` as described in [usage.md](references/usage.md), then inspect the extracted chart again.
3. For the selected chart, generate its canonical data request with `inspect --request-out <new.json>` and the exact selectors. Read [data-contract.md](references/data-contract.md) when editing the request. Update both the matrix and expected model from the user's data.
4. Run `update` with the inspected source hash, exact selectors, request and a distinct output path. Use `--execute` for the authorized copy update; omit it only when a preview is requested or useful. Experimental naming is **enabled by default** for unnamed charts on copies. `--named-only` disables it. Never require manual naming as the routine first step.
5. The command generates through official think-cell JSON, checks exact data and untouched siblings, saves/reopens natively, and produces a preview and compact report. Check `integrity_scope`: specialized charts require visual semantic review and do not have the generic native-cache parity check. Inspect the preview for clipping and label/layout problems, then deliver the verified output. For an authorized live replacement, follow the handoff reference and verify the returned slide.

Run `doctor` once for a new environment or after a dependency failure. See [setup.md](references/setup.md). Reuse successful runtime discovery for the session. Native checks remain part of each update; do not rerun broad family certification for routine work. On a new chart structure or runtime failure, isolate a representative copy and investigate only the failing operation.

## Boundaries

- Automatic naming edits only the chart and owning data-table name fields on a separate local copy. It is an experimental compatibility adapter, not an official naming API. Regenerate through official JSON before native use; never deliver a naming-only candidate.
- The ordinary single-target JSON route covers bounded pie, sequence, scatter and bubble structures with its documented count limits. For count or topology changes, inspect [multi-chart](references/multi-chart.md) and [specialized-chart](references/specialized-charts.md) routes before declaring a limitation. Gantt taskbar dates and milestone dates/markers have separate bounded adapters in [expanded controls](references/expanded-features.md); general Gantt restructuring is not covered by the ordinary JSON updater.
- Persistent or unknown external links, including linked siblings, must not pass through the internal-datasheet JSON updater. Inspect link state and use the documented link-preserving operation, including the existing-workbook rebind adapter. If no matching route exists, investigate on isolated presentation and workbook copies; never silently detach links or replace a dynamic chart with a static one.
- Preserve whole native slides and dependencies. Use the selected documented model-and-shape adapter for its permitted fields, followed by its regeneration/native readback checks. The coordinate adapter is one such route, not the only permitted native edit. A PowerPoint cache-only change does not establish a valid think-cell edit.
- Inputs, reports and outputs must stay distinct. After a timeout or uncertain native write, inspect the recorded state instead of retrying blindly. Close only task-owned presentations; never quit PowerPoint.

## Guarded native controls

For an existing chart, use [expanded native controls](references/expanded-features.md) before selecting a
specialized adapter: secondary-axis ranges and series colors; semantic scalar
and percentage labels; existing or new axis breaks; bounded CAGR annotations;
series connectors; Gantt milestones/taskbars; existing Excel-link rebinding;
and authentic error-bar, legend, or trendline donors. Each route has a narrow
native-structure contract. Preserve the user-supplied donor, use separate
working/output artifacts, and release only after its documented model and
native readback checks pass. Do not represent a prepared candidate, an
unverified donor, or a static overlay as a completed native update.

## When the requested operation has no matching adapter

An authorized chart request includes bounded local investigation needed to complete it; no separate permission is needed to prepare and test isolated copies. Use [native feature research](references/native-feature-research.md): inspect relevant current APIs and adapters, try a supported method or a narrow candidate, and verify the saved result. A capability list or failed guard describes the tested route's limits, not everything think-cell can do. Diagnose the mismatch and adapt the candidate; do not disable checks simply to force it through. Preserve the original presentation and any linked workbook. Stop with a precise blocker only when required access or inputs are missing, a materially different action needs authorization, or evidence shows the available approaches cannot produce a valid result. Report the supported portion and remaining gap without claiming unverified success.

Keep routine replies short: result, relevant limitation, file. Ask only for missing information that changes the update. Follow the user's instructions over skill preferences; do not add approval pauses for already-authorized copy work. Shared-file writes, sends and publication follow the user's existing authorization rules. Use parallel read-only work only when it saves time; serialize Office mutations. This skill does not override the selected model or reasoning effort.
