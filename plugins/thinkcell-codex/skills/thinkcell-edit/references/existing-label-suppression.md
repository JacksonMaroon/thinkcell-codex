# Existing label suppression (experimental)

`scripts/experimental_label_suppression/prepare_existing_labels.py` prepares a static candidate that removes existing native sequence-chart labels. It is for a selected, already-native chart only. It does not insert annotations, labels, shapes, or data.

Provide an exact existing chart automation name, category text, and one or more existing series names. The helper resolves the live native model from those selectors and refuses ambiguous or missing matches. It can remove the selected scalar labels, the selected chart's existing sum labels, and the existing category-axis line. `--suppress-all-existing-carrier-sum-labels` is deliberately explicit because it also includes any other existing sum labels in that one think-cell carrier.

```powershell
python scripts/experimental_label_suppression/prepare_existing_labels.py `
  --input candidate.pptx --output prepared.pptx --report prepared.json `
  --expected-source-sha256 "EXACT_CURRENT_INPUT_SHA256" `
  --chart-name "Existing chart name" --category "Existing category" `
  --series "Existing series A" --series "Existing series B" `
  --suppress-existing-sum-labels --hide-category-axis-line
```

The expected source SHA-256 is required and must match before any mutation. Input, output, and optional report paths must be distinct new files. The preparation gate also requires unique model ownership, unique physical tag ownership, no remaining dangling IDs, unchanged selected source values, unchanged embedded datasheets, and no package changes outside the selected OLE carrier, slide, slide relationships, and the removed label tag parts.

An unplaced scalar label with direct precision is also supported on the tested sequence profile. It may have a native shape name but must have no live physical shape owner and no bound text variable. That branch removes only the selected model label and its owning scalar reference; slide XML, tag parts, datasource values, and datasheets stay unchanged. The GP 29.7 label case passed full official regeneration and native reopen after this removal.

Some existing labels are direct-precision native model labels: they have a label reference and `m_prec`, but no bound `CTextVariable` and no live physical shape even if their model shape-name field is populated. The adapter supports that form only when the label is explicitly unplaced, its physical shape is proven absent, and the scalar ownership/source gates pass. In that case it changes only the selected OLE carrier.

This prepares an experiment only. Use official generation, native reopen, visual review, and a changed-data update before relying on the result. The validated canary removed two selected scalar labels and existing sum labels, retained source data, then passed native reopen and changed-data verification. It does not establish support for arbitrary label creation or unrelated annotation changes.
