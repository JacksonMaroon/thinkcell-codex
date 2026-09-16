"""Portable production entrypoint for tested line series-color styling.

The wrapper discovers staged think-cell scripts from its own directory or the
THINKCELL_NATIVE_SCRIPTS environment override, then delegates to
line_style_extension.py. Point overrides are rejected until a native point
field is proven; they are never silently removed.
"""
from __future__ import annotations

import argparse
import json
import os
import runpy
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
EXTENSION = HERE / "line_style_extension.py"


def discover_script_dir() -> Path:
    candidates = []
    override = os.environ.get("THINKCELL_NATIVE_SCRIPTS")
    if override:
        candidates.append(Path(override).expanduser())
    for ancestor in (HERE, *HERE.parents):
        candidates.extend((ancestor / "scripts", ancestor / "staged-plugin" / "skills" / "thinkcell-edit" / "scripts", ancestor / "skills" / "thinkcell-edit" / "scripts"))
    for candidate in candidates:
        if (candidate / "multi_chart_update.py").is_file() and (candidate / "thinkcell_no_click" / "implementation").is_dir():
            return candidate.resolve()
    raise FileNotFoundError("Cannot locate native think-cell scripts from this adapter's directory")


def reject_point_overrides(plan_path: Path) -> None:
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    targets = plan.get("targets")
    if not isinstance(targets, list) or not targets:
        raise ValueError("Plan must contain a nonempty targets list")
    for target in targets:
        appearance = target.get("appearance") if isinstance(target, dict) else None
        if appearance is None:
            continue
        if not isinstance(appearance, dict):
            raise ValueError("Appearance must be an object")
        points = appearance.get("point_fills", {})
        if points:
            raise ValueError("Point overrides are unsupported until native point-field readback is proven")


def main() -> None:
    if '--help' in sys.argv or '-h' in sys.argv:
        os.environ['THINKCELL_NATIVE_SCRIPTS'] = str(discover_script_dir())
        runpy.run_path(str(EXTENSION), run_name='__main__')
        return
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--plan", type=Path, required=True)
    args, _ = parser.parse_known_args()
    reject_point_overrides(args.plan)
    os.environ["THINKCELL_NATIVE_SCRIPTS"] = str(discover_script_dir())
    if not EXTENSION.is_file():
        raise FileNotFoundError(f"Missing sibling adapter: {EXTENSION}")
    runpy.run_path(str(EXTENSION), run_name="__main__")


if __name__ == "__main__":
    main()
