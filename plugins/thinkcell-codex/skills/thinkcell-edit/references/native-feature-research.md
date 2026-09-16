# Extending native features without clicking

The capability matrix describes verified adapters, not what is possible in
think-cell. When the user asks to extend an operation, investigate it rather
than treating the current matrix as proof of impossibility.

Read [native-feature-findings.md](native-feature-findings.md) before repeating
break, percentage-field or CAGR experiments. It records discriminating
controls and failed approaches from the current native model version.

## Evidence before mutation

1. Read the documented API first for an appropriate supported operation.
2. On an authorized donor copy, inspect the actual native model and its
   PowerPoint shape/field representations. Respect active `reqver` / `endver`
   fields. Do not substitute conceptual graphs for extracted native XML.
3. Trace the selected feature from chart, category and series identity through
   source bindings, owning collections, generated shape tags and text fields.
   Object IDs and visible numeric text alone are not semantic identities.
4. Form one narrow hypothesis. Check unique identifiers, resolved references,
   unchanged non-target streams and relationships before opening Office.

## Test the operation

- Keep the native model, linked text fields and physical shape bindings
  coherent. A cached rendering and its authoring state can restore each other.
- Run official regeneration first. If the requested feature is already absent,
  reject the candidate before native reopen; do not claim retained XML is proof.
- Before evaluating a changed-data feature, verify the requested data actually
  reached both the saved model and embedded datasheet. Require native chart
  consistency; a zero exit code or saved presentation can still contain a chart
  for which think-cell skipped the update. Diagnose unchanged input data before
  concluding that the feature's calculation failed.
- When cloning a label, preserve agreement between its model and PowerPoint
  shape: geometry type, bounds, text styles and native tags. Insert fields after
  paragraph properties, retain run formatting, and align precision with the
  cached display. A field binding alone does not prove this agreement.
- Require the correct selected visible feature, native save/reopen, exact data,
  untouched sources/siblings and a second changed-data update.
- For formatting-only requests, verify the requested appearance with unchanged
  data. Never change user data merely to force a refresh.
- For insertion, require a new physical native feature with correct linked
  semantics, not merely an extra model node or an ordinary PowerPoint overlay.
- Serialize Office work. A timeout or uncertain helper state must resolve
  before another Office mutation starts. Do not activate/select objects, move
  the mouse, close user presentations or quit the application.

Promote only the operation and topology actually demonstrated. Keep falsified
experiments separate from installed adapters, and record the exact difference
between a feature that is unsupported, untested, and tested but rejected.
