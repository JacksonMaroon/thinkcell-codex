"""Create from locally installed native templates, without requiring a supplied donor.

This is template generation, not an undocumented direct chart-insertion API.
No vendor assets are bundled or downloaded. Discovery is read-only; execution
uses whole-slide native extraction and the existing guarded JSON/native gates.
"""
from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
import os
from pathlib import Path
import sys
import zipfile

from lxml import etree as E

HERE = Path(__file__).resolve().parent
sys.path[:0] = [str(HERE), str(HERE / "thinkcell_no_click/implementation")]
from office_operation_lock import serialized_office
from prepare_thinkcell_name import inventory, link_contract, need
from update_thinkcell_json import model_of, validate_request, canonical_sequence_contract
from chart_semantics import kind, feature_summary


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest().upper()


def installed_roots():
    """Look only in documented installation locations, not private file trees."""
    roots = []
    for variable in ("ProgramFiles(x86)", "ProgramFiles"):
        base = os.environ.get(variable)
        if base:
            root = Path(base) / "think-cell/templates/think-cell Charts"
            if root.is_dir() and root not in roots:
                roots.append(root)
    return roots


def classify(owner_kind, family, cache):
    """Read visible cache subtype alongside native family, never guess from filename."""
    if owner_kind in {"waterfall", "mekko-percent", "mekko-units"}:
        return owner_kind
    ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    if cache is None:
        return None
    tags = [E.QName(x).localname for x in cache.iter()
            if E.QName(x).localname in {"barChart", "lineChart", "areaChart", "pieChart",
                                        "doughnutChart", "scatterChart", "bubbleChart"}]
    if len(tags) != 1:
        return "combination" if len(tags) > 1 else None
    tag = tags[0]
    if tag == "barChart":
        direction = cache.find(".//c:barChart/c:barDir", ns)
        grouping = cache.find(".//c:barChart/c:grouping", ns)
        if direction is None or grouping is None:
            return None
        orientation = {"col": "column", "bar": "bar"}.get(direction.get("val"))
        group = grouping.get("val")
        return f"{group}-{orientation}" if orientation and group in {"stacked", "clustered", "percentStacked"} else None
    if family == "CSequenceChartSE" and tag == "scatterChart":
        return "line"  # Native sequence line charts may use scatter caches
    return {"lineChart": "line", "areaChart": "area", "pieChart": "pie",
            "doughnutChart": "doughnut", "scatterChart": "scatter", "bubbleChart": "bubble"}[tag]


def inspect_template(path):
    source = Path(path).resolve()
    before = source.read_bytes()
    docs, charts, _ = inventory(before)
    counts = Counter(c["doc"]["slide_number"] for c in charts)
    rows = []
    with zipfile.ZipFile(source) as archive:
        for c in charts:
            frames = c["frames"]
            part = frames[0].get("native_chart_part") if len(frames) == 1 else None
            cache = E.fromstring(archive.read(part)) if part else None
            chart_type = classify(kind(c), c["owner"].tag, cache)
            row = {"source": str(source), "source_sha256": hashlib.sha256(before).hexdigest().upper(),
                   "slide_number": c["doc"]["slide_number"], "slide_id": c["doc"]["slide_id"],
                   "type": chart_type, "family": c["owner"].tag,
                   "exact_target": c["exact"], "frames": frames,
                   "charts_on_slide": counts[c["doc"]["slide_number"]],
                   "features_on_slide": feature_summary(c),
                   "schema_versions": {key: sorted({x.get(key) for x in c["doc"]["root"].iter() if x.get(key) is not None})
                                       for key in ("reqver", "endver")}}
            try:
                link_contract(c)
                model = model_of(c)
                row["model"] = model
                if model.get("percent_axis") and chart_type in {"stacked-column", "stacked-bar", "percentStacked-column", "percentStacked-bar"}:
                    row["type"] = "percent-stacked-" + chart_type.rsplit("-", 1)[1]
                row["generation_eligible"] = bool(c["exact"] and len(frames) == 1 and counts[c["doc"]["slide_number"]] == 1 and chart_type)
            except ValueError as error:
                row.update(generation_eligible=False, reason=str(error))
            rows.append(row)
    need(source.read_bytes() == before, "Installed template changed during discovery")
    return rows


def discover(roots=None):
    roots = installed_roots() if roots is None else [Path(root).resolve() for root in roots]
    rows, errors = [], []
    for root in roots:
        need(root.is_dir(), "Template library root does not exist: " + str(root))
        for source in sorted(root.rglob("*.potx")):
            try:
                rows.extend(inspect_template(source))
            except (ValueError, KeyError, zipfile.BadZipFile) as error:
                errors.append({"source": str(source), "error": str(error)})
    return {"read_only": True, "method": "local_installed_native_template_discovery",
            "roots": [str(root) for root in roots], "templates": rows, "errors": errors,
            "direct_chart_constructor": False}


def compatible(row, request):
    """Apply the actual adapter's contract before starting any Office work."""
    source = Path(row["source"])
    need(digest(source) == row["source_sha256"], "Template changed; discover again")
    _, charts, _ = inventory(source.read_bytes())
    c = next(c for c in charts if c["doc"]["slide_number"] == row["slide_number"]
             and c["frames"] == row["frames"])
    link_contract(c)
    if c["owner"].tag == "CSequenceChartSE" and kind(c) == "CSequenceChartSE":
        from multi_chart_update import simple_sequence_contract
        contract = simple_sequence_contract(c, request, True)
        return "flexible_sequence", contract
    validate_request(request, c["owner"].tag)
    contract = canonical_sequence_contract(source, request,
             target={"slide_id": c["doc"]["slide_id"], **c["frames"][0]})
    fill = c["table"].find("m_bExcelOnTop")
    need(fill is None or fill.get("val") != "1", "Non-sequence seed uses datasheet fills requiring another regeneration route")
    return "single_target", contract


def choose_template(catalog, chart_type, request, slide_number=None, source=None, required_features=()):
    matches = [row for row in catalog["templates"] if row["generation_eligible"] and row["type"] == chart_type
               and (slide_number is None or row["slide_number"] == slide_number)
               and (source is None or Path(row["source"]).resolve() == Path(source).resolve())
               and all(row["features_on_slide"].get(feature, 0) > 0 for feature in required_features)]
    passed, rejected = [], []
    for row in matches:
        try:
            route, contract = compatible(row, request)
            passed.append((row, route, contract))
        except (ValueError, KeyError) as error:
            rejected.append({"source": row["source"], "slide_number": row["slide_number"], "reason": str(error)})
    need(passed, "No installed native template passes the requested type/data contract: " + json.dumps(rejected))
    # Prefer an unannotated baseline and deterministic native library identity.
    passed.sort(key=lambda item: (sum(v for k, v in item[0]["features_on_slide"].items()
                                  if "legend" not in k.lower()), item[0]["source"], item[0]["slide_number"]))
    return passed[0]


def sequence_plan(chart, request):
    """Bind both native chart tag and extracted slide number, as choose requires."""
    need(chart["doc"]["slide_number"] == 1, "Construction seed must be the extracted first slide")
    return {"allow_sequence_count_change": True, "targets": [{
        "selector": {"slide_number": 1, "shape_tag": chart["frames"][0]["shape_tag"]},
        "name": chart.get("owner_name") or "ConstructedChart", "data": request}]}


@serialized_office
def construct(args):
    import thinkcell
    request_path = args.data_json.resolve()
    request_bytes = request_path.read_bytes()
    request_hash = hashlib.sha256(request_bytes).hexdigest().upper()
    request = json.loads(request_bytes.decode("utf-8-sig"))
    catalog = discover(args.library_root)
    required_features = getattr(args, "require_feature", []) or []
    need(all(isinstance(feature, str) and feature.startswith("C") and feature.isidentifier() for feature in required_features),
         "Required features must be exact native class names from the library inventory")
    row, route, contract = choose_template(catalog, args.chart_type, request, args.slide_number, args.template, required_features)
    dest = args.output_directory.resolve()
    need(not dest.exists() and not dest.is_relative_to(HERE.parent), "Use a new output directory outside the skill")
    result = {"status": "INSTALLED_TEMPLATE_PREFLIGHT_PASS", "method": "installed_native_template_then_official_json",
              "direct_chart_constructor": False, "template": row, "route": route,
              "data_contract": contract, "writes": False,
              "required_features_on_slide": required_features,
              "data_request_sha256": request_hash,
              "inherited_slide_copy_requires_review": True}
    if not args.execute:
        return result
    need(digest(request_path) == request_hash, "Requested data changed during preflight")
    # create retains the full native slide, its theme and dependent storage.
    extracted = thinkcell.create(argparse.Namespace(input=Path(row["source"]), expected_sha256=row["source_sha256"],
                         slide_number=row["slide_number"], output_directory=dest, style_file=args.style_file,
                         native_percent=False, data_json=None, execute=True))
    seed = Path(extracted["output"])
    frozen_request = dest / "requested-data.json"
    with frozen_request.open("xb") as handle:
        handle.write(request_bytes)
    _, charts, _ = inventory(seed.read_bytes())
    need(len(charts) == 1, "Extracted installed seed no longer contains exactly one chart")
    chart = charts[0]
    output = dest / "native-candidate.pptx"
    if route == "flexible_sequence":
        # Installed templates can use an older native grammar. Let the official
        # generator normalize it with unchanged data through the existing
        # single-target/native gates. Count edits then compare against this
        # normalized grammar; never exempt font/axis fields from that gate.
        baseline = thinkcell.baseline_request(seed, chart)
        baseline_path = dest / "normalization-request.json"
        thinkcell.write_json(baseline_path, baseline)
        normalization = thinkcell.update(normalization_args(seed, chart, baseline_path, dest, args.ppttc))
        seed = Path(normalization["output"])
        normalized_rows = inspect_template(seed)
        need(len(normalized_rows) == 1 and normalized_rows[0]["type"] == row["type"],
             "Native template type changed during unchanged-data normalization")
        result["unchanged_data_normalization"] = normalization
        _, charts, _ = inventory(seed.read_bytes())
        need(len(charts) == 1, "Normalization changed native chart count")
        chart = charts[0]
        from multi_chart_update import run
        plan = sequence_plan(chart, request)
        plan_path = dest / "construction-plan.json"
        thinkcell.write_json(plan_path, plan)
        generated = run(argparse.Namespace(input=seed, expected_sha256=digest(seed), plan=plan_path,
                        output=output, report=dest / "data-generation.json", prepare=False, execute=True, ppttc=args.ppttc))
        thinkcell.write_json(dest / "data-generation.json", generated)
    else:
        generated = thinkcell.update(argparse.Namespace(input=seed, expected_sha256=digest(seed),
                        output=output, report=dest / "data-generation.json", data_json=frozen_request,
                        slide_number=1, slide_id=None, shape_id=None, shape_tag=chart["frames"][0]["shape_tag"],
                        named_only=False, ppttc=args.ppttc, execute=True, require_native_percent=False))
    need(digest(row["source"]) == row["source_sha256"], "Installed source changed during generation")
    need(digest(request_path) == request_hash and digest(frozen_request) == request_hash,
         "Requested data changed during generation; output withheld")
    final_rows = inspect_template(output)
    need(len(final_rows) == 1 and final_rows[0]["type"] == row["type"], "Requested native chart type did not survive generation")
    need(all(final_rows[0]["features_on_slide"].get(feature, 0) > 0 for feature in required_features),
         "Required native feature did not survive generation")
    delivered = dest / "populated.pptx"
    with delivered.open("xb") as handle:
        handle.write(output.read_bytes())
    result.update(status="NATIVE_TEMPLATE_GENERATED_VISUAL_REVIEW_REQUIRED", writes=True,
                  output=str(delivered), output_sha256=digest(delivered), source_unchanged=True,
                  native_generation=generated, render=generated["render"])
    result["requested_type_retained"] = True
    result["requested_native_feature_models_retained"] = True
    result["feature_semantics_review_required"] = bool(required_features)
    thinkcell.write_json(dest / "construction.json", result)
    return result


def normalization_args(seed, chart, request_path, dest, ppttc):
    """Exact unchanged-data update used to let think-cell upgrade its own schema."""
    return argparse.Namespace(input=seed, expected_sha256=digest(seed),
                 output=dest / "normalized-seed.pptx", report=dest / "normalization.json",
                 data_json=request_path, slide_number=1, slide_id=None, shape_id=None,
                 shape_tag=chart["frames"][0]["shape_tag"], named_only=False,
                 ppttc=ppttc, execute=True, require_native_percent=False)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    listing = sub.add_parser("list")
    listing.add_argument("--library-root", type=Path, action="append")
    listing.add_argument("--type", dest="chart_type")
    create = sub.add_parser("create")
    create.add_argument("--type", dest="chart_type", required=True)
    create.add_argument("--library-root", type=Path, action="append")
    create.add_argument("--template", type=Path)
    create.add_argument("--slide-number", type=int)
    create.add_argument("--data-json", type=Path, required=True)
    create.add_argument("--output-directory", type=Path, required=True)
    create.add_argument("--style-file", type=Path)
    create.add_argument("--require-feature", action="append", help="Exact native feature class from list; selected and checked after regeneration")
    create.add_argument("--ppttc")
    create.add_argument("--execute", action="store_true")
    args = parser.parse_args()
    try:
        result = discover(args.library_root) if args.command == "list" else construct(args)
        if args.command == "list" and args.chart_type:
            result["templates"] = [row for row in result["templates"] if row["type"] == args.chart_type]
        print(json.dumps(result, indent=2, allow_nan=False))
    except (ValueError, KeyError) as error:
        print(json.dumps({"status": "REJECTED", "error": str(error)}), file=sys.stderr)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
