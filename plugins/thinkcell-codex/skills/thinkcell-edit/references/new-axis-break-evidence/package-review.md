# Staged new native break package review

## Verdict

No blocker found for the explicitly bounded profile. The staged package's
claims match the two canaries and the implementation is appropriately
fail-closed for the tested 3-series by 4-category clustered-column chart with
one unique outlier and a 20% user break.

## Guards reviewed

- `experimental_axis_break_insert.py` pins both source and donor SHA-256,
  requires matching model version `38764`, one ordinary sequence chart, one
  no-break axis, one length-12 interval vector, and no target tickmark graph.
- It preserves the ordinary axis child order by replacing the interval vector
  and empty break collection in place. Fresh model IDs, fresh opaque physical
  tags, fresh slide shape IDs, relationship target checks, content types and
  unchanged non-model OLE streams are checked before native use.
- `run_native_axis_break_insert.py` validates the narrow data contract, rejects
  nonfinite/boolean/negative values, requires one unique outlier, checks exact
  model and embedded datasheet cells, and requires the user-owned 20% native
  break. The prepared seed explicitly requires two shapes and length-3
  intervals; generated/reopened output requires the observed one-shape,
  length-12 topology.
- Native integrity correctly excludes only the cache-parity assertions known
  to be invalid for a transformed broken-axis cache, while retaining the
  remaining package audit and native source/sibling preservation checks.
- The repeat path is source-hash guarded and requires the supplied pre-plan
  when the current data differs, as exercised by the 90-to-120 canary.

The three staged offline regression tests passed: wrong source hash,
nonfinite/boolean or inconsistent data, and rejection of the unregenerated
two-shape seed as a final one-shape result.

## Canary consistency

`packaged-break-insert-v2.json` records successful official generation and
native reopen. The prepared candidate has two seed `CPPTBreakShape` nodes and
six physical tags; generation normalizes it to one native break shape and three
tags, with a gap `[45, 90]` crossed only by 90.

`packaged-break-repeat.json` retains the same native break and shape IDs while
changing the gap to `[45, 120]`, crossed only by 120. Both canaries report
exact model/datasheet readback, native reopen, unchanged source, unchanged
other presentations, and unchanged donor where applicable.

## Scope note

The package documentation correctly limits support to this one-outlier,
one-gap clustered-column profile. No additional certification is needed for
the stated claims. Arbitrary break insertion remains excluded.
