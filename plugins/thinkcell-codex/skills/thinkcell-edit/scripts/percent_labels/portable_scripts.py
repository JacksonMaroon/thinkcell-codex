"""Resolve the installed think-cell script directory for portable helpers."""
from __future__ import annotations

import os
from pathlib import Path


def _candidates(start: Path):
    roots = [start, *start.parents]
    seen = set()
    suffixes = (
        Path("scripts"),
        Path("..") / "scripts",
        Path("staged-plugin") / "skills" / "thinkcell-edit" / "scripts",
        Path("skills") / "thinkcell-edit" / "scripts",
        Path("thinkcell-edit") / "scripts",
    )
    for root in roots:
        for suffix in suffixes:
            candidate = (root / suffix).resolve()
            if candidate not in seen:
                seen.add(candidate)
                yield candidate


def resolve_plugin_scripts(start: str | Path) -> Path:
    override = os.environ.get("THINKCELL_PLUGIN_SCRIPTS")
    if override:
        candidate = Path(override).expanduser().resolve()
        if (candidate / "chart_geometry.py").is_file():
            return candidate
        raise FileNotFoundError(
            "THINKCELL_PLUGIN_SCRIPTS must point to a directory containing chart_geometry.py: "
            + str(candidate)
        )
    start = Path(start).resolve()
    for candidate in _candidates(start):
        if (candidate / "chart_geometry.py").is_file():
            return candidate
    raise FileNotFoundError(
        "Could not auto-discover think-cell scripts via chart_geometry.py from " + str(start)
    )
