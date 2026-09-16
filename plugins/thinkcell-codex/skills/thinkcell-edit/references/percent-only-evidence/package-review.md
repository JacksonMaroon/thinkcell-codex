# Final percent package review

## Verdict

The earlier hardcoded ratio/display blocker is closed in the staged runner. The canary evidence reports `30.0/200.0` read from the saved native output and `15%`, with `output_withheld_until_feature_pass: true`.

## Guard evidence

- `ratio()` derives numerator from the explicit plan matrix and denominator from its `100%=` row, applies donor zero-decimal `ROUND_HALF_UP`, and rejects invalid denominator values.
- `saved_ratio()` independently inventories the saved native chart, reads the embedded datasheet, checks owner `CSequenceChartSE`, `m_ect=0`, model version `38764`, percent-axis semantics, selected scalar 191, and compares model numerator and category extent to the saved datasheet.
- Preflight requires the original dual label's absolute and relative text to agree with the saved native data.
- Post-COM checks preserve the original relative field GUID, retain its active percent formatter and zero decimal precision, and verify absolute raw data is unchanged.
- Final checks compare the saved native ratio and derived display to the requested ratio, require one percent field with the original parentheses, and only then copy the staged file to the delivered output.
- `write_new(..., 'xb')`, distinct-path checks, source hash checks, and work-directory containment prevent accidental overwrite or source mutation.

## Canary readback

The canary report is `ALL_GATES_PASS`, with exact datasheet cells, one native chart, unchanged source and non-target streams, native reopen pass, and unchanged other presentations. The evidence records `actual_native_ratio: 30.0/200.0`, `expected_display: 15%`, and `percent_feature_gate: PASS`.

## Residual scope

The route remains intentionally bounded to the tested 4-series by 3-category sequence profile, scalar 191, label tag, active model version, zero-decimal precision, and relative-owned percent field topology. No blocker remains for this scoped package; general chart types and nonzero precision remain outside its certification.
