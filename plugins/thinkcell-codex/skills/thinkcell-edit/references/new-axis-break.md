# Bounded native axis-break insertion adapter

This package contains an experimental copy-only adapter for one demonstrated
think-cell topology: a current-version clustered-column chart with no existing
break, one length-12 ordinary primary-axis interval vector, no native axis
tickmark collection, and one outlying scalar that crosses the generated gap.

Use `scripts/run_native_axis_break_insert.py`; the preparation module is `scripts/experimental_axis_break_insert.py`.
It resolves `chart_geometry`, `runtime`, and the guarded OLE writer from the
same bundled skill. It never resolves an installed skill or a workspace path.
The included OLE writer is placed at its matching bundled-skill-relative path.

## What it checks offline

- Exact SHA-256 hashes for the no-break target and break donor before any write,
  and again after preparation.
- One exact ordinary sequence chart, zero target breaks, the demonstrated
  interval/tickmark topology, and matching native model version.
- The original target axis field order. It replaces interval and break fields
  at their existing slots instead of prepending them.
- Fresh model IDs, opaque tag-to-physical-shape mapping, fresh PowerPoint
  nonvisual shape IDs, live relationships, tag-part content types, no duplicate
  ZIP entries, and unchanged non-model OLE streams.

It rejects wrong hashes, an existing-break target, a mismatched current-model
version, and any unsupported ordinary-axis topology before writing an output.

## Scope and remaining gates

The accepted proof normalized to one `CPPTBreakShape` and three physical tags,
because exactly one scalar crossed the gap. That is expected for this topology;
the two-shape donor graph is only a seed for insertion. The adapter does not
support arbitrary charts or multiple crossing bars.

`run_native_axis_break_insert.py` performs those gates in one serialized
operation: prepare, official JSON generation, native save/reopen, exact
datasource/model and physical-tag readback, and a rendered preview. It reports
the preview path under `preview` and accepts only the proved post-generation
topology: one native break, one `CPPTBreakShape`, and three linked physical
tags.

The seed `assets/feature-fixtures/axis-break-seed-38764.pptx` is bundled. Its SHA-256 is `4E03F81243B29D7BBFFC131077BB64C0FDACA0CE407692DDC3C9E979B7A8289E`. Inspect the actual target and provide its current hash and complete data contract. The bundled ordinary fixture and `new-axis-break-evidence/same-data.json` are reproducible test inputs, not defaults for user data. This route currently requires an already named chart; use the normal naming adapter on a copy first when necessary.

Initial insertion (paths relative to the skill root):

```powershell
python scripts/run_native_axis_break_insert.py --mode insert `
  --input ordinary.pptx --expected-source-sha256 <ordinary SHA256> `
  --break-donor assets/feature-fixtures/axis-break-seed-38764.pptx --expected-donor-sha256 <donor SHA256> `
  --plan same-data.json --ppttc "C:\Program Files (x86)\think-cell\ppttc.exe" `
  --execute --output inserted.pptx --report inserted-report.json
```

Repeat update of a proved one-break output, using its current-data plan for the
precheck and a changed-data plan for generation:

```powershell
python scripts/run_native_axis_break_insert.py --mode repeat `
  --input inserted.pptx --expected-source-sha256 <inserted SHA256> `
  --pre-plan same-data.json --plan changed-data.json `
  --ppttc "C:\Program Files (x86)\think-cell\ppttc.exe" `
  --execute --output updated.pptx --report updated-report.json
```

Both modes require new output/report paths and recheck the source hash after
the native operation. The two completed proof reports and renders are listed
in [new-axis-break-evidence.json](new-axis-break-evidence.json).

The insertion and changed-data certification is already recorded. For routine use, run the requested same-data insertion or requested data update, inspect its preview and report, and stop. Do not change the user's data merely to repeat certification. A new topology needs separate research.
