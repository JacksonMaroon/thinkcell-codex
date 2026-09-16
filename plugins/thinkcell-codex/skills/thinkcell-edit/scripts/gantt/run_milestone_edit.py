"""Portable integration CLI for one semantic milestone edit."""
from __future__ import annotations
import argparse, hashlib, json
from pathlib import Path
from milestone_adapter import Package

def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", type=Path, required=True)
    ap.add_argument("--output", type=Path, required=True)
    ap.add_argument("--report", type=Path, required=True)
    ap.add_argument("--expected-sha256", required=True)
    ap.add_argument("--new-date", required=True)
    ap.add_argument("--milestone-date")
    ap.add_argument("--milestone-id")
    ap.add_argument("--style")
    a = ap.parse_args()
    before = sha(a.source)
    if before.lower() != a.expected_sha256.lower():
        raise SystemExit("source SHA256 mismatch")
    if a.output.exists() or a.report.exists():
        raise SystemExit("output/report already exists")
    edit = Package(a.source).edit(
        a.output,
        selector={"milestone_id": a.milestone_id, "date_value": a.milestone_date, "style": a.style},
        new_date=a.new_date,
    )
    result = {
        "status": "PAIRED_GANTT_MILESTONE_EDIT_PREPARED",
        "source": str(a.source.resolve()),
        "output": str(a.output.resolve()),
        "source_sha256_before": before,
        "source_sha256_after": sha(a.source),
        "source_unchanged": before == sha(a.source),
        "edit": edit,
    }
    a.report.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, indent=2))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
