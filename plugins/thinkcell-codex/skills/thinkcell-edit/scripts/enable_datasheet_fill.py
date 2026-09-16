"""Enable a selected chart's native `Use Datasheet Fill on Top` setting.

This is a narrowly guarded private-model preparation step. It changes only the
selected chart's existing ``m_bExcelOnTop`` field; official JSON regeneration
and native evidence are mandatory before an output may be released.
"""
from __future__ import annotations

import argparse, copy, io, json, subprocess, sys, tempfile, zipfile
from pathlib import Path
from lxml import etree as E

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from chart_geometry import _chart_identity, _select_chart, inventory, need, sha, streams, xml  # noqa: E402
from runtime import powershell, powershell_env  # noqa: E402

FIELD = "m_bExcelOnTop"


def make_plan(path, selection):
    raw = Path(path).read_bytes(); _, charts, _ = inventory(raw)
    chart = _select_chart(charts, selection)
    node = chart["table"].find(FIELD)
    need(node is not None and node.get("val") in {"0", "1"}, "Selected chart has no unambiguous datasheet-fill setting")
    target = _chart_identity(chart)
    target["table_id"] = chart["table"].get("id")
    return {"schema_version": 1, "source_sha256": sha(raw), "target": target,
            "field": FIELD, "old_value": node.get("val"), "new_value": "1",
            "required_followup": ["official JSON regeneration with typed fill cells", "native reopen and render RGB readback", "changed-data regeneration with different category totals", "label and annotation retention review"]}


def _assert_only_field_changed(before, after, plan):
    original, candidate = xml(before), xml(after)
    table = candidate.xpath(".//*[@id='%s']" % plan["target"]["table_id"])
    need(len(table) == 1, "Selected data-table identity changed")
    field = table[0].find(FIELD); need(field is not None and field.get("val") == plan["new_value"], "Wrong datasheet-fill setting")
    field.set("val", plan["old_value"])
    need(E.tostring(original, method="c14n") == E.tostring(candidate, method="c14n"), "Non-target native model field changed")


def prepare(input_path, output_path, plan):
    src, out = Path(input_path), Path(output_path)
    need(src.resolve() != out.resolve() and not out.exists(), "Distinct new output required")
    expected = make_plan(src, plan.get("target")); need(plan == expected, "Stale or modified datasheet-fill plan")
    raw = src.read_bytes(); _, charts, _ = inventory(raw); chart = _select_chart(charts, plan["target"])
    doc, before = chart["doc"], chart["doc"]["streams"][("think-cellXML",)]
    root = xml(before); table = root.xpath(".//*[@id='%s']" % plan["target"]["table_id"])
    need(len(table) == 1, "Selected data-table identity changed")
    node = table[0].find(FIELD); need(node is not None and node.get("val") == plan["old_value"], "Datasheet-fill setting changed")
    node.set("val", plan["new_value"]); after = E.tostring(root, encoding="utf-8"); _assert_only_field_changed(before, after, plan)
    with tempfile.TemporaryDirectory(prefix="tc_fill_", dir=out.parent) as temp:
        temp = Path(temp); carrier, payload = temp / "carrier.bin", temp / "model.xml"; carrier.write_bytes(doc["ole"]); payload.write_bytes(after)
        completed = subprocess.run([powershell(), "-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File", str(HERE / "thinkcell_no_click/implementation/replace_ole_stream.ps1"), "-StoragePath", str(carrier), "-StreamBytesPath", str(payload)], capture_output=True, text=True, env=powershell_env(), timeout=60)
        need(completed.returncode == 0, "Datasheet-fill preparation failed: " + completed.stderr[-500:])
        changed = carrier.read_bytes(); updated = streams(changed); need(set(updated) == set(doc["streams"]), "OLE stream inventory changed")
        need(all(v == updated[k] for k, v in doc["streams"].items() if k != ("think-cellXML",)), "Other OLE stream changed")
        _assert_only_field_changed(before, updated[("think-cellXML",)], plan); need(sha(src.read_bytes()) == plan["source_sha256"], "Source changed")
        with zipfile.ZipFile(io.BytesIO(raw)) as zin, zipfile.ZipFile(out, "x") as zout:
            zout.comment = zin.comment
            for item in zin.infolist(): zout.writestr(copy.copy(item), changed if item.filename == doc["part"] else zin.read(item.filename))
    return {"status": "DATASHEET_FILL_ENABLED_REGENERATION_REQUIRED", "source_sha256": plan["source_sha256"], "output_sha256": sha(out.read_bytes()), "only_selected_m_bExcelOnTop_changed": True}


def main():
    p = argparse.ArgumentParser(description=__doc__); s = p.add_subparsers(dest="command", required=True)
    m = s.add_parser("make-plan"); m.add_argument("--input", required=True); m.add_argument("--selection-json", required=True); m.add_argument("--plan-out", required=True)
    r = s.add_parser("prepare"); r.add_argument("--input", required=True); r.add_argument("--plan", required=True); r.add_argument("--output", required=True); a = p.parse_args()
    if a.command == "make-plan":
        plan = make_plan(a.input, json.loads(Path(a.selection_json).read_text())); Path(a.plan_out).write_text(json.dumps(plan, indent=2) + "\n"); print(json.dumps(plan, indent=2))
    else: print(json.dumps(prepare(a.input, a.output, json.loads(Path(a.plan).read_text())), indent=2))

if __name__ == "__main__": main()
