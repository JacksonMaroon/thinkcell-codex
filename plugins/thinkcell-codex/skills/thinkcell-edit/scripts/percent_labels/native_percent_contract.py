"""Read-only before/after contract for native percentage labels.

The contract is deliberately independent of Office.  It snapshots the
verified direct percentage rows discovered for one exact native chart and
checks that the same semantic indices and precision survive a later file.
Category and series names are reported for review but indices are the stable
identity used by the preservation check because native data edits may rename
labels without changing their positions.
"""
from __future__ import annotations

import copy
import math
import zipfile
from pathlib import Path
from typing import Any

try:  # Package import for callers using ``percent_labels.native_percent_contract``.
    from .discover_percent_semantics import discover, inventory
except ImportError:  # Direct script/test import from the percent_labels folder.
    from discover_percent_semantics import discover, inventory  # type: ignore


def _identity(target: Any) -> dict[str, Any]:
    """Normalize an inventory chart, discovery owning_chart, or selector dict."""
    if isinstance(target, str):
        return {"shape_tag": target}
    if not isinstance(target, dict):
        raise ValueError("target must be an inventory chart or identity mapping")
    if "owning_chart" in target and isinstance(target["owning_chart"], dict):
        target = target["owning_chart"]

    frames = target.get("frames") or []
    frame = frames[0] if len(frames) == 1 and isinstance(frames[0], dict) else {}
    owner = target.get("owner")
    owner_id = owner.get("id") if hasattr(owner, "get") else None
    owner_type = owner.get("tag") if isinstance(owner, dict) else getattr(owner, "tag", None)
    doc = target.get("doc")
    carrier = doc.get("part") if isinstance(doc, dict) else None
    identity = {
        "carrier": target.get("carrier") or target.get("chart_part") or carrier,
        "owner_id": target.get("owner_id") or target.get("chart_model_id") or owner_id,
        "owner_type": target.get("owner_type") or owner_type,
        "shape_tag": target.get("shape_tag") or frame.get("shape_tag"),
        "shape_id": target.get("shape_id") or frame.get("shape_id"),
        "native_chart_part": target.get("native_chart_part") or frame.get("native_chart_part"),
    }
    return {key: value for key, value in identity.items() if value is not None}


def _inventory_has_target(path: str | Path, identity: dict[str, Any]) -> bool:
    """Return whether an inventory chart matching ``identity`` exists in path."""
    try:
        _, charts, _ = inventory(Path(path).read_bytes())
    except (OSError, ValueError, KeyError, zipfile.BadZipFile):
        return False
    for chart in charts:
        candidate = _identity(chart)
        if identity.get("carrier") is not None and candidate.get("carrier") != identity["carrier"]:
            continue
        if identity.get("owner_type") is not None and candidate.get("owner_type") != identity["owner_type"]:
            continue
        if identity.get("shape_tag") is not None:
            if candidate.get("shape_tag") == identity["shape_tag"]:
                return True
        elif identity.get("owner_id") is not None and candidate.get("owner_id") == identity["owner_id"]:
            return True
    return False


def _target_rows(rows: list[dict], target: Any,
                 path: str | Path | None = None) -> tuple[list[dict], dict[str, Any]]:
    identity = _identity(target)
    carrier = identity.get("carrier")
    shape_tag = identity.get("shape_tag")
    owner_id = identity.get("owner_id")
    if not carrier and not shape_tag and owner_id is None:
        raise ValueError("target must identify a carrier, owner, or exact frame shape tag")

    def matches(row: dict) -> bool:
        chart = row.get("owning_chart") or {}
        if carrier is not None and chart.get("carrier") != carrier:
            return False
        # A preserved exact frame tag is the durable after-file identity.  An
        # owner id is accepted only when no shape tag was supplied.
        if shape_tag is not None:
            return chart.get("shape_tag") == shape_tag
        if owner_id is not None and chart.get("owner_id") != owner_id:
            return False
        return True

    selected = [row for row in rows if matches(row)]
    if not selected:
        # A valid sequence chart can have no scalar labels while a sibling
        # chart contributes discovery rows.  Inventory proves that this is a
        # real target with no applicable contract; otherwise preserve the
        # fail-closed missing-target error (especially for stale frame tags).
        if path is not None and _inventory_has_target(path, identity):
            return [], identity
        raise ValueError("target chart was not found by carrier and exact frame identity")
    charts = {(row["owning_chart"].get("carrier"), row["owning_chart"].get("owner_id"),
               row["owning_chart"].get("shape_tag")) for row in selected}
    if len(charts) != 1:
        raise ValueError("target matched multiple native charts")
    return selected, identity


def _is_percent_candidate(row: dict) -> bool:
    precision = row.get("direct_precision") or {}
    # Candidate detection is intentionally broader than verification: a
    # percent suffix that is present but no longer rendered/verified must make
    # the contract fail closed as partial recognition.
    return bool(precision.get("suffix") == "%"
                and precision.get("decimal_digits") is not None)


def _label(row: dict) -> dict[str, Any]:
    denominator = row.get("denominator")
    numerator = row.get("numerator")
    if denominator in (None, 0):
        raise ValueError("native percentage row has no nonzero denominator")
    ratio = float(numerator) / float(denominator)
    if not math.isfinite(ratio):
        raise ValueError("native percentage row has a non-finite ratio")
    agreement = row.get("cache_agreement") or {}
    if not row.get("verified_direct_percentage") or not agreement.get("verified"):
        raise ValueError("native percentage row is not fully verified")
    expected = agreement.get("expected")
    if expected is None or not math.isclose(float(expected), ratio * 100.0,
                                             rel_tol=1e-6, abs_tol=1e-6):
        raise ValueError("native percentage row cache ratio disagrees with model values")
    return {
        "category_index": row.get("category_index"),
        "series_index": row.get("series_index"),
        "category": row.get("category"),
        "series": row.get("series"),
        "shape_tag": row.get("shape_tag"),
        "label_id": row.get("label_id"),
        "scalar_id": row.get("scalar_id"),
        "numerator": numerator,
        "denominator": denominator,
        "ratio": ratio,
        "direct_precision": copy.deepcopy(row.get("direct_precision")),
        "cache_value": agreement.get("observed"),
        "cache_expected": expected,
        "cache_series_index": (row.get("cache") or {}).get("series_index"),
        "cache_category_index": (row.get("cache") or {}).get("category_index"),
    }


def _not_applicable(path: str | Path, identity: dict[str, Any]) -> dict[str, Any]:
    return {"status": "not_applicable", "applicable": False, "input": str(path),
            "target": identity, "labels": [], "label_count": 0, "count": 0}


def snapshot(path: str | Path, target: Any) -> dict[str, Any]:
    """Capture the verified native percentage contract for ``target``.

    ``target`` may be an inventory chart returned by ``inventory()``, a
    discovery row's ``owning_chart`` object, a selector mapping containing
    ``carrier``/``owner_id``/``shape_tag``, or an exact shape-tag string.
    """
    path = Path(path)
    identity = _identity(target)
    # Pie, scatter, and other native families are outside this sequence-chart
    # contract.  Return early so a mixed deck's sequence sibling cannot make a
    # non-sequence target look like a missing chart.
    if identity.get("owner_type") not in (None, "CSequenceChartSE"):
        return _not_applicable(path, identity)
    discovered = discover(path)
    # Unsupported or label-free charts are outside this contract.  Discovery
    # deliberately returns no rows for them, so callers can keep a generic
    # update/inspect route usable and treat this feature as not applicable.
    if not discovered:
        return _not_applicable(path, identity)
    rows, identity = _target_rows(discovered, target, path)
    if not rows:
        return _not_applicable(path, identity)
    candidates = [row for row in rows if _is_percent_candidate(row)]
    if not candidates:
        return _not_applicable(path, identity)
    if len(candidates) != len(rows) or any(not row.get("verified_direct_percentage") for row in candidates):
        raise ValueError("partial native direct percentage recognition for target chart")

    labels = sorted((_label(row) for row in candidates),
                    key=lambda item: (item["category_index"], item["series_index"]))
    indices = {(item["category_index"], item["series_index"]) for item in labels}
    if len(indices) != len(labels):
        raise ValueError("native percentage labels do not have unique semantic indices")
    chart = rows[0].get("owning_chart") or {}
    return {
        "status": "native_percentage",
        "applicable": True,
        "input": str(path),
        "target": identity,
        "chart": {key: chart.get(key) for key in
                  ("carrier", "owner_id", "owner_type", "shape_tag", "shape_id",
                   "native_chart_part", "exact")},
        "label_count": len(labels),
        "count": len(labels),
        "category_count": len({item["category_index"] for item in labels}),
        "series_count": len({item["series_index"] for item in labels}),
        "labels": labels,
    }


def verify(before: dict | None, after_path: str | Path, target: Any) -> dict[str, Any]:
    """Verify semantic identity and precision after a native save/reopen."""
    if not before or not before.get("applicable") or not before.get("labels"):
        return {"status": "not_applicable", "applicable": False, "before": before,
                "after": None, "checks": {"before_present": False}}
    after = snapshot(after_path, target)
    if after.get("status") != "native_percentage":
        raise ValueError("native percentage contract disappeared after native save")
    before_labels = {(row["category_index"], row["series_index"]): row
                     for row in before["labels"]}
    after_labels = {(row["category_index"], row["series_index"]): row
                    for row in after["labels"]}
    checks = {
        "chart_shape_tag_preserved": bool(before.get("chart", {}).get("shape_tag"))
        and before["chart"]["shape_tag"] == after.get("chart", {}).get("shape_tag"),
        "chart_family_preserved": before.get("chart", {}).get("owner_type")
        == after.get("chart", {}).get("owner_type"),
        "label_indices_preserved": set(before_labels) == set(after_labels),
        "label_count_preserved": before.get("label_count") == after.get("label_count"),
        "category_count_preserved": before.get("category_count") == after.get("category_count"),
        "series_count_preserved": before.get("series_count") == after.get("series_count"),
        "all_after_labels_verified": after.get("label_count", 0) == len(after_labels),
        "direct_precision_preserved": False,
    }
    if checks["label_indices_preserved"]:
        checks["direct_precision_preserved"] = all(
            before_labels[key].get("direct_precision") == after_labels[key].get("direct_precision")
            for key in before_labels
        )
    if not all(checks.values()):
        raise ValueError("native percentage contract verification failed: " + str(checks))
    return {"status": "native_percentage_preserved", "applicable": True,
            "before": before, "after": after, "checks": checks}
