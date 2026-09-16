# Native series connector adapter

`series_connector_adapter.py` imports one genuine donor connector into a
target native chart. It prepares a candidate without opening Office.

Use `--input <target.pptx> --inspect` to discover chart names and semantic
category/series labels. To prepare, supply `--input`, `--donor`, fresh
`--output` and `--report` paths, `--series-name`, `--from-category` and
`--to-category`. Supply `--chart-name` and `--donor-chart-name` when the input
contains several charts. Indices are available for intentionally unlabeled
source categories/series.

Use `--expected-input-sha256` and `--expected-donor-sha256` to bind the run to
inspected inputs. Both files are checked again before the final report is
written. A preflight mismatch creates no candidate; a later source change
withholds the success report and leaves any candidate unverified.

The donor must contain a connector between the specified semantic endpoints.
The adapter resolves its complete model and property structure, remaps model
IDs and imports the tagged physical shape with fresh relationship identities.

## Native follow-up

Run `thinkcell.py inspect --data --request-out <fresh.json>` on the prepared
candidate to obtain the canonical request. Use the reported slide/chart
identity and generated automation name. Update the complete matrix and
expected model, then use `thinkcell.py update` with distinct output/report
paths and the normal official regeneration and native verification gates.

Verify exact named-datasheet values, connector count and semantic endpoint
ownership, saved model and physical geometry, and the final render. Test
changed values at an interior boundary to confirm both endpoints recalculate
with the relevant stacked proportions. A nonzero movement alone is not proof
of correct binding.

The tested portable adapter retained the pre-existing connector and the new
interior connector. Both new endpoints moved by the mathematically expected
boundary proportions after data changed. Evidence is bundled at
`../../../verification/feature-expansion-20260916/series-connector.json`.
