# Native total-label precision

`total_label_precision.py` synchronizes the active decimal setting and native format representations of an existing
dynamic total label. Its checked scope is one native sequence chart, one
category, and one scalar group with a numeric total. It accepts 0–3 decimal
places and rejects unsupported currency prefixes, suffixes, magnitudes and
nonstandard locales rather than guessing their meaning.

Use the existing feature pipeline with this feature entry:

```json
{"kind":"total_label_precision","selector":{"slide_number":1,"shape_tag":"CHART_TAG"},"decimal_digits":1}
```

The feature plan also needs the source hash and complete `data_plan` described
in [reliable-features.md](reliable-features.md). Supply exactly the intended
data; use a canonical same-data plan for formatting-only requests. The output
is withheld unless the requested decimal setting, selected total source value,
and its exact linked dynamic field agree after regeneration and native reopen.

Verification follows category → scalar group → absolute total source → text
variable → its unique dynamic field. A matching number elsewhere on the slide
is not sufficient. The adapter updates the selected `m_bstrFormat` key together
with that exact field's type and displayed text. Its reversal proofs allow only
those formatting changes and active precision; all numeric model values,
numeric source caches, datasource streams, other fields, and package parts stay
unchanged. Never clear `m_bstrFormat` or modify it without its bound field:
it connects the text variable to the corresponding PowerPoint field.

Native proof kept the original data total `39.2` and displayed `39.20` after
same-data official regeneration and native reopen. A subsequent genuine change
to total `40.3` recomputed the same bound label to `40.30`. Precision-only edits
and numeric-cache invalidation both retained stale text and are not this route.
Do not change user data to force formatting refresh. The prepared copy still
requires official regeneration, native reopen and exact field verification.
