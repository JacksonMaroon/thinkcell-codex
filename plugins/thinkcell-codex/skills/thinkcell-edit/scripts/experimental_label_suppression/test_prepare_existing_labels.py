"""Static regression test for prepare_existing_labels.py; does not open Office."""
import argparse, hashlib, importlib.util, json, tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("labels", HERE / "prepare_existing_labels.py")
labels = importlib.util.module_from_spec(spec); spec.loader.exec_module(labels)

p = argparse.ArgumentParser(); p.add_argument("--fixture", required=True)
a = p.parse_args(); fixture = Path(a.fixture)
fixture_hash = hashlib.sha256(fixture.read_bytes()).hexdigest().upper()
with tempfile.TemporaryDirectory() as td:
    td = Path(td); out, report_path = td / "out.pptx", td / "report.json"
    report = labels.prepare(
        fixture, out, fixture_hash, "KTC_STACKBAR_ABS_01", "USPS",
        ["Managed Logistics", "Brokerage"], True, True, True, report_path,
    )
    saved = json.loads(report_path.read_text())
    assert report["source_unchanged"] and saved["output_sha256"] == report["output_sha256"]
    assert len([x for x in report["selected"] if x["kind"] == "scalar_label"]) == 2
    assert len([x for x in report["selected"] if x["kind"] == "sum_label"]) == 8
    assert report["axis"]["from"] == "1" and report["axis"]["to"] == "0"
    assert "no dangling model idrefs" in report["gates_passed"]
    assert report["expected_source_sha256"] == fixture_hash
    no_axis = labels.prepare(
        fixture, td / "no-axis.pptx", fixture_hash, "KTC_STACKBAR_ABS_01", "USPS",
        ["Managed Logistics", "Brokerage"], False, False, False, td / "no-axis.json",
    )
    assert no_axis["axis"] is None
    mismatch_out = td / "mismatch.pptx"
    try:
        labels.prepare(fixture, mismatch_out, "0" * 64, "KTC_STACKBAR_ABS_01", "USPS", ["Brokerage"], False, False, False)
        raise AssertionError("mismatched SHA was accepted")
    except Exception as exc:
        assert "SHA-256 mismatch" in str(exc) and not mismatch_out.exists()
print("PASS static existing-label suppression fixture")
