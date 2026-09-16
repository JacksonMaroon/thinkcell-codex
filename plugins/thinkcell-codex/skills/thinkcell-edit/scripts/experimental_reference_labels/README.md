# Reference scalar-label formatting (experimental)

`prepare_multiline_scalar_format.py` is a narrowly profiled native formatting adapter. It recognizes only the recorded model profile: version 38764, label 458, scalar 431, source 454 with the existing `$29.7B` representation and 11-point fonts. It preserves the existing `$` prefix and sets the suffix to `B`, `Expansion`, and `opportunity` on separate lines.

The command requires an exact input SHA-256 and distinct new output and report paths. It verifies the label/scalar/source bindings and formatting profile, reverses the full model mutation for comparison, preserves every non-model CFB stream, and permits only the selected OLE carrier to differ in the package.

```powershell
python scripts/experimental_reference_labels/prepare_multiline_scalar_format.py `
  --input source.pptx --expected-source-sha256 EXACT_SHA256 `
  --output prepared.pptx --report prepared.json
```

This does not create arbitrary labels, annotations, or other features. The suffix form passed full official JSON regeneration, native reopen, and dynamic changed-data readback from 29.7 to 30.7. Native auto-placement controls the final position; this adapter makes no fixed-position claim. Do not treat static preparation as visual proof.
