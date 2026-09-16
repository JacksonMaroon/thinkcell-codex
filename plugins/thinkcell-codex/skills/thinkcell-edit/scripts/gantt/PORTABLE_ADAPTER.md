# Portable paired Gantt adapter

`portable_gantt_adapter.py` is a native adapter for a
paired native date plus geometry edit. It discovers the single Gantt owner,
activity vectors, bars, generic-line bindings, weekly scale boxes, visible
shape tags and slide transforms from the package. The caller selects a record
semantically (row plus existing start/end dates, or a selected native id) and
only then mutates that record; no donor-specific bar, line or shape IDs are
hardcoded in the transformation.

The adapter writes typed ISO datetime values, maps task endpoints to the
containing/next weekly scale boundary, updates the bound OLE line rectangle and
updates the visible shape carrying the same `THINKCELLSHAPEDONOTDELETE` tag.
Callers must run native reopen/readback afterward. This is an explicit paired
adapter and does not claim automatic think-cell reflow.

The portable entry point currently edits one taskbar's typed dates and paired
position. Milestones remain covered by the separate native-backed candidate;
the donor contains no explicit dependency/link node to parameterize safely.

Example:

```python
from portable_gantt_adapter import paired_date_geometry_edit
paired_date_geometry_edit(
    "named-gantt.pptx", "portable-output.pptx",
    {"start": "2012-01-02", "end": "2012-01-28"},
    "2012-01-09", "2012-02-04",
)
```

The portable output passed the lane verifier, including model consistency,
owner relationship, bar-to-line graph, visible tag binding and non-target
record preservation.

Release checks exercised relocation and native calendar mapping on copied donors, preserving non-target records.

For an integration-ready command-line call, use the bundled
`run_paired_gantt_edit.py`. Supply --expected-sha256 to bind the source bytes. It refuses stale output and report paths and records
`source_sha256_before`, `source_sha256_after`, and `source_unchanged` in the
requested report. Use the shared native verifier and inspect the saved output; see
[expanded controls](../../references/expanded-features.md).
