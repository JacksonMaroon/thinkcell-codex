"""Build a generalized add-step request and official ppttc job offline."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from waterfall_add_step_adapter import add_step, build_ppttc_job


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True, type=Path)
    ap.add_argument("--request", required=True, type=Path)
    ap.add_argument("--category", required=True)
    ap.add_argument("--contributions-json", required=True, help="JSON array with one finite number or null per existing series")
    ap.add_argument("--name", help="Optional automation name; defaults to the input request name")
    ap.add_argument("--output-request", required=True, type=Path)
    ap.add_argument("--job", required=True, type=Path)
    args = ap.parse_args()
    template = args.template.resolve()
    request_path = args.request.resolve()
    output_request = args.output_request.resolve()
    job_path = args.job.resolve()
    if not template.is_file() or template.suffix.lower() != ".pptx":
        raise ValueError(f"Template must be an existing .pptx: {template}")
    if not request_path.is_file() or request_path.suffix.lower() != ".json":
        raise ValueError(f"Request must be an existing .json: {request_path}")
    if output_request == job_path:
        raise ValueError("Output request and ppttc job must be distinct files")
    if output_request in {template, request_path} or job_path in {template, request_path}:
        raise ValueError("Output files must be distinct from template and request inputs")
    if output_request.exists() or job_path.exists():
        raise ValueError("Output request or ppttc job already exists; choose fresh output paths")
    request = json.loads(request_path.read_text(encoding="utf-8-sig"))
    contributions = json.loads(args.contributions_json)
    result = add_step(request, args.category, contributions)
    if args.name:
        result["name"] = args.name
    job = build_ppttc_job(str(template), result)
    output_request.parent.mkdir(parents=True, exist_ok=True)
    job_path.parent.mkdir(parents=True, exist_ok=True)
    with output_request.open("x", encoding="utf-8") as handle:
        json.dump(result, handle, indent=2, ensure_ascii=False, allow_nan=False)
    with job_path.open("x", encoding="utf-8") as handle:
        json.dump(job, handle, indent=2, ensure_ascii=False, allow_nan=False)
    print(json.dumps({"status": "OFFLINE_ADD_STEP_JOB_READY", "request": str(output_request), "job": str(job_path), "category": args.category}, ensure_ascii=False))


if __name__ == "__main__":
    main()
