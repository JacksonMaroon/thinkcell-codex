# Existing native broken-axis charts

`scripts/broken_axis_update.py` updates the tested user-supplied one-break sequence
donor through official JSON and native PowerPoint save/reopen. It is a data
update route that preserves the existing break. For position changes, use the
verified `axis_break_position.py` adapter described below. New break insertion
is a separate research operation.

Use a user-supplied native one-break sequence donor. Use normal `create` to copy it
first, then use this dedicated wrapper rather than the generic update path.
Generate the canonical request with `thinkcell.py inspect --input copy.pptx
--slide-number 1 --shape-tag TAG --request-out data.json`. Wrap that
request in a normal complete multi-chart plan with exactly one target:

```json
{"targets":[{"selector":{"slide_number":1,"shape_tag":"TAG"},"name":"MY_BREAK_CHART","data":{"matrix":[],"expected_model":{}}}]}
```

Fill `data` with the complete canonical request and the user's actual data.
Run `python scripts/broken_axis_update.py --input copy.pptx
--expected-sha256 HASH --plan plan.json --output result.pptx
--report result.json --execute` with distinct paths.

The wrapper accepts exactly one existing user-configured break with two native
break shapes and a finite increasing value range. It checks exact model and
datasheet values, native source/other-presentation preservation, the existing
break profile and visible fill sequence. A broken axis transforms the visible
PowerPoint cache, so the ordinary one-to-one cache/data parity check is replaced
by this explicit specialized scope. Only the observed vendor owner-model
normalization is permitted. Unknown grammar changes fail closed.

Two successive data updates passed on the installed native runtime, including
the already-normalized second input. The final file is published only after
the specialized gate passes. Review the rendered discontinuity and labels;
the wrapper does not certify that a scale break is appropriate for the user's
business argument.
# Reposition an existing break

`axis_break_position.py` changes the existing native break's fractional
position on an authorized copy, runs the dedicated broken-axis data updater,
and requires native geometry and fraction readback before releasing output.
The checked profile is one sequence chart, one user break, and two native
break-shape objects. The fraction must be between 0.05 and 0.95.

```powershell
python scripts/axis_break_position.py --input donor.pptx --shape-tag CHART_TAG --fraction 0.2 --plan position.json --data-plan same-data.json --output result.pptx --report result.json --execute
```

Use fresh distinct output, plan and report paths. Obtain the canonical data
plan through chart inspection and keep its values unchanged unless the user
also requests a data change. Same-data tests at 0.2 and 0.8 produced different
native break geometry; the 0.2 setting survived a subsequent changed-data
update and native reopen. A no-op position request does not require movement.

Do not edit `m_vecintvlOrdinal` to position the break: official generation
recomputes those intervals from the data. Break insertion requires a separately
verified operation.
