# Ordered native break insertion review

## Verdict

The execution is a real new native axis break on the ordinary clustered-column
chart for the tested one-outlier/one-gap topology. I found no blocker in the
saved PPTX evidence.

## Direct saved-model checks

I inspected the original no-break source
`work/plugin-features-v2/break-insertion/ordinary-column-normalized-native.pptx`
and both saved outputs, rather than relying on the reports alone.

- Original chart identity is model `63`, automation name
  `TC_AUTO_4FD4FB579092_516_8`, shape tag `t3NL4BqaIewPMw01VrwngWg`, primary
  axis `62`; its axis has `m_cdaxisbreak length="0"` and no break nodes.
- Same-data saved output retains the same chart and axis identities, but owns
  `m_cdaxisbreak length="1"` pointing to `CDataAxisBreak id="266"`, which owns
  `CPPTBreakShape id="267"`. The model has three linked opaque physical shape
  tags, and all three resolve to physical PPTX tag parts. The original has zero
  matching break tags.
- The break node is therefore attached to the original ordinary chart model,
  not a swapped broken-axis donor. The native generator's reduction from the
  six prepared shapes to one three-tag break shape is expected for the tested
  topology: only the 90 outlier crosses the generated gap.
- Changed-data saved output retains the same chart, axis, break and break-shape
  identities (`266`/`267`) and the same three physical tag bindings. Its model
  and datasheet contain `Column 1 / 3rd Qtr = 120.0`; the other cells remain
  unchanged.

## Data and visual evidence

Direct model/datasheet readback is exact:

- Same-data: the 90 value is present in model and embedded datasheet.
- Changed-data: 90 changes to 120.0 in both model and embedded datasheet; the
  category extent changes from 146.6 to 176.6.

The saved native renders show a visible axis discontinuity through the 90 bar
and, after the changed-data update, through the 120 bar. This confirms the
break is rendered by think-cell's native chart cache rather than an overlay.

Both native reports record successful save/reopen, unchanged source, and
unchanged other presentations. The changed-data pass is a meaningful second
update because it changes the outlier while retaining the native break graph.

## Scope

This supports promotion only for the demonstrated current clustered-column
profile: one outlier crossing one generated gap. It does not establish support
for multiple breaks, multiple outliers/gaps, other chart families, or arbitrary
shared-axis topologies.
