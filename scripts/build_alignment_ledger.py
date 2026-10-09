"""Join declared and observed interfaces; presence alone never means compatibility."""

import argparse
import json
from pathlib import Path


def surface(runtime):
    result = {}
    for module in runtime["modules"]:
        for item in module["exports"]:
            result[item["id"]] = item
            if item["kind"] == "type":
                for member in item["members"]:
                    result[item["id"] + "." + member["name"]] = member
    return result


def format_records(formats):
    rows = []
    for direction in ("save", "load"):
        baseline = formats["reports"][direction + "_commercial"]["records"]
        current = {item["name"]: item for item in formats["reports"][direction + "_current"]["records"]}
        checks = {item["format"]: item for item in formats["independent_docx_checks"].get("load_current", [])}
        for sample in baseline:
            name = sample["name"]
            row = {"id": "format." + direction + "." + name, "value": sample["value"],
                   "baseline_behavior": sample.get("outcome", {"status": sample.get("status"), "loaded": sample.get("loaded")}),
                   "implementation": current[name].get("outcome", {"status": current[name].get("status"), "loaded": current[name].get("loaded")}),
                   "gap": "format fidelity and all options/overloads remain unverified",
                   "dependencies": ["aspose.words.Document.save" if direction == "save" else "aspose.words.Document.__init__"],
                   "approach": "implement format model/DOM mapping and error semantics; validate fixed samples and all options",
                   "samples": [{"evidence": "commercial-26.9-format-behavior.json", "direction": direction, "format": name}],
                   "validation": ["bounded actual output/error observation"],
                   "exit_criteria": "all applicable structures/options/overloads/errors, DOM roundtrip and rendering verified",
                   "limits": ["one small input; evaluation watermark; no broad fidelity acceptance"], "delivery": "baseline_observation_only"}
            observed = current[name]
            if direction == "save":
                baseline_result, current_result = sample["outcome"], observed["outcome"]
                if baseline_result["status"] == "returned" and current_result["status"] == "raised":
                    row["gap"] = "baseline save returns an output; current save rejects or fails"
                elif baseline_result["status"] == current_result["status"] == "returned" and baseline_result.get("return_type") != current_result.get("return_type"):
                    row["gap"] = "save return contract differs; format fidelity/options remain unverified"
                elif baseline_result["status"] == current_result["status"] == "raised" and baseline_result.get("exception_type") != current_result.get("exception_type"):
                    row["gap"] = "save exception type differs; other error semantics remain unverified"
            elif observed.get("status") == "raised":
                row["gap"] = "current loading or roundtrip fails for the baseline-produced sample"
                row["implementation"].update({key: observed[key] for key in ("phase", "exception_type") if key in observed})
            if direction == "load" and sample.get("status") == "missing_legal_sample_not_verified":
                row["gap"] = "legal input sample missing; capability remains unverified"
            if direction == "load" and name in checks:
                row["independent_docx_check"] = checks[name]
                if checks[name]["literal_markup_present"]:
                    row["gap"] = "markup accepted as literal text; semantic structure not preserved"
            rows.append(row)
    assert len({row["id"] for row in rows}) == len(rows)
    return rows


def build(declarations, baseline, current, observations=None, formats=None):
    declared = {item["id"]: item for module in declarations["modules"] for item in module["symbols"]}
    observed = surface(baseline)
    implemented = surface(current)
    records = []
    for name in sorted(declared.keys() | observed.keys()):
        runtime = observed.get(name)
        target = current["root_module"] + name.removeprefix(baseline["root_module"])
        owner = None if runtime is None else runtime.get("owner")
        present = target in implemented
        records.append({
            "id": name,
            "declaration": name if name in declared else owner if owner in declared else None,
            "runtime_observed": runtime is not None,
            "current_entrypoint": target if present else None,
            "implementation": "name_present_behavior_unverified" if present else "equivalent_public_name_absent_internal_implementation_unassessed",
            "gap": "behavior_not_compared" if present else "public_surface_and_behavior_not_aligned",
            "dependencies": [owner] if owner and owner != name else [],
            "scope_confirmation": "published_declaration" if name in declared or owner in declared else "runtime_only_requires_public_support_confirmation",
        })
        evidence_key = name if observations and name in observations else owner
        if observations and evidence_key in observations:
            records[-1]["behavior_evidence"] = {"file": "commercial-26.9-dom-structure.json", "capability": evidence_key,
                                                "scope": "partial plain-DOCX cases; not all node kinds or overloads"}
            records[-1]["baseline_behavior"] = "partially_observed; see bounded behavior evidence"
    assert len({item["id"] for item in records}) == len(records)
    return {"schema": 1, "baseline_version": declarations["version"], "current_runtime_version": current["runtime_version"],
            "record_defaults": declarations["capability_defaults"],
            "evidence": {"declarations": "commercial-26.9-api.json", "baseline_runtime": "commercial-26.9-runtime.json", "current_runtime": "commercial-26.9-current-runtime.json"},
            "status": "surface inventory joined; member-specific behavior, dependencies and validation remain incomplete",
            "capability_count": len(records), "records": records, "format_capabilities": [] if formats is None else format_records(formats)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("declarations", type=Path)
    parser.add_argument("baseline", type=Path)
    parser.add_argument("current", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--observations", type=Path)
    parser.add_argument("--formats", type=Path)
    args = parser.parse_args()
    observations = None if args.observations is None else json.loads(args.observations.read_text())["capability_observations"]
    formats = None if args.formats is None else json.loads(args.formats.read_text())
    result = build(*(json.loads(path.read_text()) for path in (args.declarations, args.baseline, args.current)), observations=observations, formats=formats)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"records": result["capability_count"], "matching_names": sum(item["current_entrypoint"] is not None for item in result["records"])}))


if __name__ == "__main__":
    main()
