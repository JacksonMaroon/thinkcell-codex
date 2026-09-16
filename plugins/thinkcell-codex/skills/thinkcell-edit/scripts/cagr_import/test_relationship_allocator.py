"""Offline regression test for imported slide relationship allocation."""
from __future__ import annotations

import importlib.util
from pathlib import Path


def main() -> None:
    path = Path(__file__).with_name("import_cagr_closure.py")
    spec = importlib.util.spec_from_file_location("import_cagr_closure", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("importer is missing")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    used = {"rId1", "rId3"}
    allocated = []
    serial = 1
    for _ in range(3):
        rid, serial = module.allocate_rid(used, serial)
        allocated.append(rid)
    if allocated != ["rId2", "rId4", "rId5"] or len(used) != 5:
        raise RuntimeError(f"relationship allocator regression: {allocated}")
    print({"status": "RELATIONSHIP_ALLOCATOR_PASS", "allocated": allocated, "unique": len(set(allocated)) == len(allocated)})


if __name__ == "__main__":
    main()
