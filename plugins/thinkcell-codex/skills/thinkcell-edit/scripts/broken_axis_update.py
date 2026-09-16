"""Dedicated data-update wrapper for the validated one-break user-supplied donor.

This is deliberately narrower than the ordinary sequence route. It accepts one
exact chart with one existing `CDataAxisBreak`, preserves that break while
official JSON changes data, and recognizes only the owner-model rewrite
observed on the current think-cell build. It never creates, moves, or removes
a break, and it does not relax ordinary-chart validation.
"""
from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
import math
import sys
import zipfile
from pathlib import Path
from types import SimpleNamespace
from lxml import etree as E

HERE = Path(__file__).resolve().parent


def _support_root():
    """Locate this packaged skill's sibling helpers."""
    if (HERE / "multi_chart_update.py").is_file():
        return HERE
    raise RuntimeError("Cannot locate the packaged think-cell update helpers")


SUPPORT = _support_root()
sys.path[:0] = [str(SUPPORT), str(SUPPORT / "thinkcell_no_click" / "implementation")]

import chart_semantics
import multi_chart_update as multi
from prepare_thinkcell_name import inventory, need

_TRANSIENT_OWNER_TAGS = {"m_edir", "m_nGid", "m_vecgidGroup", "m_rectgrdlnanchorPlot", "m_signCategoryLabelDirection", "m_precOrdinal"}
_RAW_OWNER = multi.owner_semantics


def _break(chart):
    owner, ids = chart["owner"], chart["doc"]["ids"]
    need(owner.tag == "CSequenceChartSE", "Only sequence charts are supported")
    primary = owner.find("m_daxisPrimaryValue")
    need(primary is not None and primary.get("idref") in ids, "Missing primary value axis")
    axis = ids[primary.get("idref")]
    need(axis.tag == "CSequenceChartDataAxis", "Unexpected primary value-axis native type")
    entries = axis.findall("m_cdaxisbreak/elem")
    need(len(entries) == 1 and entries[0].get("idref") in ids, "Require exactly one existing native axis break")
    value = ids[entries[0].get("idref")]
    need(value.tag == "CDataAxisBreak", "Unexpected axis-break native type")
    fraction = value.find("m_fFraction")
    user = value.find("m_bUser")
    shapes = value.findall("m_cpptbreakshp/elem")
    need(fraction is not None and user is not None and len(shapes) == 2,
         "Axis break does not match the validated two-shape donor profile")
    shape_ids = [item.get("idref") for item in shapes]
    need(all(shape_id in ids and ids[shape_id].tag == "CPPTBreakShape" for shape_id in shape_ids)
         and len(set(shape_ids)) == 2,
         "Axis break does not reference two valid native break shapes")
    f = float(fraction.get("val"))
    need(math.isfinite(f) and 0.0 < f < 1.0 and user.get("val") == "1",
         "Axis break is not a positive user-configured native break")
    low, high = axis.find("m_fMinValue"), axis.find("m_fMaxValue")
    need(low is not None and high is not None and math.isfinite(float(low.get("val")))
         and math.isfinite(float(high.get("val"))) and float(high.get("val")) > float(low.get("val")),
         "Broken axis has an invalid value range")
    return {"break_count": 1, "shape_count": len(shapes), "user": user.get("val"), "fraction": f,
            "axis_minimum": float(low.get("val")), "axis_maximum": float(high.get("val"))}


def _owner_items(chart):
    return _RAW_OWNER(chart)


def _validate_known_owner_rewrite(before, after):
    left, right = _owner_items(before), _owner_items(after)
    # A second update begins from the already-normalized official grammar.
    # In that case no owner rewrite is expected; the normalized topology must
    # still be identical.
    if left == right:
        return "already_normalized_no_delta"
    left_only = Counter(left) - Counter(right)
    right_only = Counter(right) - Counter(left)
    expected_removed = Counter({("m_edir", (("val", "-1"),), ""): 1,
                                ("m_evectorsortmode", (("val", "0"),), ""): 1})
    need(left_only == expected_removed, "Unexpected pre-generation owner grammar")
    allowed = {"m_nGid", "m_vecgidGroup", "m_rectgrdlnanchorPlot", "m_signCategoryLabelDirection", "m_precOrdinal"}
    need({item[0] for item in right_only} == allowed and sum(right_only.values()) == len(allowed)
         and all(count == 1 for count in right_only.values()),
         "Official JSON owner rewrite differs from the validated broken-axis profile")
    return "known_current_build_profile"


def _normalized_owner(chart):
    result, seen_vector_sort = [], False
    for item in _owner_items(chart):
        if item[0] in _TRANSIENT_OWNER_TAGS:
            continue
        # The validated vendor rewrite removes one duplicate default vector
        # sort node.  Keep its canonical first occurrence, while preserving
        # every other owner node and its order.
        if item[0] == "m_evectorsortmode":
            if seen_vector_sort:
                continue
            seen_vector_sort = True
        result.append(item)
    return result


def _special_scope(path, audit):
    docs, charts, _ = inventory(Path(path).read_bytes())
    need(len(docs) == len(charts) == 1, "Broken-axis wrapper requires one chart on one slide")
    signature = _break(charts[0])
    exclusions = {"no_model_visible_chart_mismatch", "strict_parity_available", "strict_parity_pass"}
    failures = [key for key, value in audit["assertions"].items() if not value and key not in exclusions]
    need(not failures, "Integrity checks failed: " + ", ".join(failures))
    return {"native_cache_parity": "known_axis_break_cache_transform",
            "specialized_visual_semantics_review_required": True,
            "axis_break_profile": signature,
            "excluded_inapplicable_checks": sorted(exclusions)}


def _axis_break_cache_fills(path, candidate, contract=None):
    """Compare the raw visible cache fills by stable bar/series order.

    This donor's break cache contains one visible bar-series for two model
    series, so the ordinary name-to-series mapping is inapplicable.  We still
    compare every serialized RGB fill before and after official regeneration.
    """
    part = candidate["frames"][0]["native_chart_part"]
    need(part, "Target has no native chart cache part")
    with zipfile.ZipFile(path) as package:
        root = E.fromstring(package.read(part))
    ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
          "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    values = {}
    for index, series in enumerate(root.findall(".//c:barChart/c:ser", ns)):
        fills = tuple(series.xpath(".//a:srgbClr/@val", namespaces=ns))
        if fills:
            values["bar-series-%d" % index] = fills
    return values


def _axis_break_raw_color_fingerprint(path, candidate):
    """Stable one-break color fingerprint from serialized visible RGB fills."""
    return tuple(sorted(_axis_break_cache_fills(path, candidate).items()))


def run(args):
    # The standard runner remains untouched. These process-local substitutions
    # apply only after the one-break profile has been proved above.
    original_owner, original_scope, original_fills, original_color_fingerprint = (
        multi.owner_semantics, multi.audit_scope, multi.cache_fills, multi.raw_color_fingerprint)
    def checked_owner(chart):
        return _normalized_owner(chart)
    multi.owner_semantics = checked_owner
    chart_semantics.audit_scope = _special_scope
    multi.audit_scope = _special_scope
    multi.cache_fills = _axis_break_cache_fills
    multi.raw_color_fingerprint = _axis_break_raw_color_fingerprint
    try:
        # Validate profile before naming/generation, then independently assert
        # the exact vendor rewrite once it has been generated.
        _, before_charts, _ = inventory(args.input.read_bytes())
        need(len(before_charts) == 1, "Broken-axis wrapper requires one native chart")
        before_break = _break(before_charts[0])
        # The shared runner publishes after its own gates. Keep its candidate
        # private until this wrapper has also accepted the exact owner/break
        # profile, then publish atomically with an exclusive create.
        pending = args.output.parent / (args.output.stem + "_axis_break_pending.pptx")
        need(not pending.exists() and not args.output.exists(), "Output already exists")
        inner = SimpleNamespace(**vars(args))
        inner.output = pending
        report = multi.run(inner)
        if args.execute and report.get("status") == "ALL_GATES_PASS":
            # The report does not expose generated path; use its deterministic
            # staging folder and prove the known rewrite on that exact artifact.
            stage = Path(report["staging_directory"])
            _, pre, _ = inventory((stage / "named-1.pptx").read_bytes())
            _, post, _ = inventory((stage / "generated.pptx").read_bytes())
            rewrite = _validate_known_owner_rewrite(pre[0], post[0])
            _, final, _ = inventory(pending.read_bytes())
            after_break = _break(final[0])
            need(after_break["break_count"] == before_break["break_count"] and after_break["shape_count"] == before_break["shape_count"],
                 "Native axis break was lost or changed shape count")
            report["axis_break_contract"] = {"before": before_break, "after": after_break,
                                              "owner_rewrite": rewrite}
            with args.output.open("xb") as published:
                published.write(pending.read_bytes())
            report["output"] = str(args.output)
            report["output_sha256"] = hashlib.sha256(args.output.read_bytes()).hexdigest().upper()
            report["published_after_axis_break_profile_gate"] = True
        return report
    finally:
        multi.owner_semantics, multi.audit_scope, multi.cache_fills, multi.raw_color_fingerprint = (
            original_owner, original_scope, original_fills, original_color_fingerprint)
        chart_semantics.audit_scope = original_scope


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("input", "plan", "output", "report"):
        parser.add_argument("--" + name, type=Path, required=True)
    parser.add_argument("--expected-sha256", required=True)
    parser.add_argument("--ppttc")
    parser.add_argument("--prepare", action="store_true")
    parser.add_argument("--execute", action="store_true")
    args = parser.parse_args()
    result = run(args)
    args.report.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps({"status": result["status"], "report": str(args.report)}))
