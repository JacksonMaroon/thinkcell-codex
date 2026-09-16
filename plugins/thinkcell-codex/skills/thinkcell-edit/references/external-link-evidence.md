# Existing Excel link auto-update

An existing auto-updating native link was tested using a disposable local
workbook and presentation. Editing the workbook changed the first native
series name to `TC_AUTO_PROBE_V4` and its first value to `314160`; the categories
remained `FY25` and `FY26`. The saved presentation was reopened and read back.

The link GUID, link ID, advise-sink identity, range, auto-update flag and
workbook/range moniker remained intact. Source identity, workbook numeric
readback and preservation of other open presentations/workbooks passed.
No `Send`, `UpdateBatch` or internal-datasheet replacement was used.

The strict audit retained one pre-existing failure: master/layout embeddings
had multiple owners in both baseline and output. All other audit checks,
including model/visible-chart strict parity, passed. Do not suppress this
check globally; compare the exact ownership records when reviewing a similar
source.

This establishes that an existing persistent link can continue updating
without clicks. Resolve the actual source workbook and native link identity
before selecting this route. Update only explicitly requested cells, retain
numeric types, save and verify both files, and preserve other open work.
The internal-datasheet JSON path must continue rejecting linked charts.
Creating arbitrary new links and accepting unknown link structures still
need separate implementation and proof. The disposable test runner is retained
as research evidence, not shipped as a general workbook-edit command.
