"""Verify frozen inventory and corpus integrity, not behavioral compatibility."""

import argparse
import hashlib
import json
import runpy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile


def digest(data):
    return hashlib.sha256(data).hexdigest()


def verify_format_outputs(report, archive_path):
    raw = archive_path.read_bytes()
    assert digest(raw) == report["outputs"]["sha256"], "format archive digest mismatch"
    checked = 0
    with ZipFile(BytesIO(raw)) as archive:
        for key, observed in report["reports"].items():
            for record in observed["records"]:
                if key.startswith("save_"):
                    for output in record["files"]:
                        value = archive.read(key + "/" + record["name"] + "/" + output["name"])
                        assert len(value) == output["size"] and digest(value) == output["sha256"]
                        checked += 1
                elif "saved_docx" in record:
                    output = record["saved_docx"]
                    assert digest(archive.read(key + "/" + output["name"])) == output["sha256"]
                    checked += 1
        from docx import Document

        for key, checks in report["independent_docx_checks"].items():
            for check in checks:
                doc = Document(BytesIO(archive.read(key + "/" + check["format"] + ".docx")))
                text = "\n".join([p.text for p in doc.paragraphs] +
                                 [cell.text for table in doc.tables for row in table.rows for cell in row.cells])
                assert len(doc.paragraphs) == check["paragraph_count"] and len(doc.tables) == check["table_count"]
                assert {label: label in text for label in check["labels_present"]} == check["labels_present"]
                assert any(token in text for token in ("<pkg:package", "<html", "<w:wordDocument", "<w:document")) == check["literal_markup_present"]
    return checked


def verify_import_outputs(root):
    inspect = runpy.run_path(str(Path(__file__).parents[1] / "docs" / "probes" / "inspect_import_outputs.py"))["inspect_document"]
    baseline = json.loads((root / "import-node-26.9.json").read_text())
    current = json.loads((root / "import-node-current.json").read_text())
    checked, actual = 0, {}
    for report, phases in ((baseline, baseline["reports"]), (current, {"candidate": current["report"]})):
        raw = (root / report["archive"]).read_bytes()
        assert digest(raw) == report["archive_sha256"], "import archive digest mismatch"
        with ZipFile(BytesIO(raw)) as archive:
            for phase, observed in phases.items():
                checks = report["independent_checks"] if phase == "candidate" else report["independent_checks"][phase]
                recorded = {row["output"]: row for row in checks}
                assert len(recorded) == len(checks) == len(observed["records"])
                actual[phase] = {}
                for row in observed["records"]:
                    name = row["output"] if phase == "candidate" else phase + "/" + row["output"]
                    data = archive.read(name)
                    assert digest(data) == row["sha256"], "import output digest mismatch"
                    result = {"output": row["output"], **inspect(BytesIO(data))}
                    assert result == recorded[row["output"]], "import independent observation mismatch"
                    actual[phase][row["output"]] = result
                    checked += 1
    matched = {row["output"] for row in current["report"]["records"] if row["outcome"]["status"] == "returned"}
    assert matched == set(current["matched_observations"])
    for name in matched:
        assert actual["candidate"][name] == actual["commercial"][name]
    assert current["validation"]["full_import_acceptance"] is False
    return checked


def verify_style_toggles(root):
    report = json.loads((root / "style-toggles-26.9.json").read_text())
    raw = (root / "corpus" / "style-toggles-26.9.zip").read_bytes()
    assert report["version"] == "26.9.0" and report["module"] == "aspose.words"
    assert report["full_format_acceptance"] is False and report["licensed"] is False
    assert digest(raw) == report["corpus_sha256"], "style corpus digest mismatch"
    generated = runpy.run_path(str(Path(__file__).parents[1] / "docs" / "probes" / "style_toggles.py"))["inputs"]
    expected = {name + ".docx" for name, _ in generated()}
    assert len(expected) == len(report["records"]) == 745
    with ZipFile(BytesIO(raw)) as archive:
        assert set(archive.namelist()) == expected == {row["input"] for row in report["records"]}
        for row in report["records"]:
            assert digest(archive.read(row["input"])) == row["sha256"]
            assert isinstance(row["bold"], bool) and isinstance(row["italic"], bool)
    current = json.loads((root / "style-toggles-current.json").read_text())
    native = {row["input"]: row for row in report["records"]}
    mismatches = {}
    for phase in ("before", "candidate"):
        observed = current["reports"][phase]
        assert observed["corpus_sha256"] == report["corpus_sha256"]
        rows = {row["input"]: row for row in observed["records"]}
        assert len(rows) == len(observed["records"]) == 745 and set(rows) == expected
        assert all(row["sha256"] == native[name]["sha256"] for name, row in rows.items())
        assert all(isinstance(row["bold"], bool) and isinstance(row["italic"], bool) for row in rows.values())
        mismatches[phase] = sorted(name for name, row in rows.items() if row != native[name])
    assert mismatches["before"] == sorted(current["before_mismatches"])
    assert mismatches["candidate"] == current["candidate_mismatches"] == []
    assert current["validation"]["full_format_acceptance"] is False
    return len(report["records"])


def verify(root):
    def read(name):
        return json.loads((root / name).read_text())

    declarations = read("commercial-26.9-api.json")
    ledger = read("commercial-26.9-capabilities.json")
    enums = read("commercial-26.9-all-enums.json")
    baseline = read("commercial-26.9-baseline.json")
    registry = baseline["documentation_registry"]
    raw_registry = (root / registry["path"]).read_bytes()
    assert digest(raw_registry.replace(b"\r\n", b"\n")) == registry["sha256"]
    assert [(row["url"], row["sha256"]) for row in json.loads(raw_registry)["documents"]] == [
        (row["url"], row["sha256"]) for row in baseline["documentation"]]
    symbols = [item["id"] for module in declarations["modules"] for item in module["symbols"]]
    assert declarations["version"] == ledger["baseline_version"] == enums["version"] == "26.9.0"
    assert declarations["module_count"] == len(declarations["modules"])
    assert declarations["symbol_count"] == len(symbols) == len(set(symbols))
    assert ledger["capability_count"] == len(ledger["records"]) == len({row["id"] for row in ledger["records"]})
    assert enums["enum_count"] == len(enums["enums"])
    assert enums["value_count"] == sum(len(item["values"]) for item in enums["enums"])
    assert enums["alias_count"] == sum(v["name"] != v["canonical_name"] for item in enums["enums"] for v in item["values"])
    formats = read("commercial-26.9-format-behavior.json")
    assert digest((root / formats["source"]).read_bytes()) == formats["source_sha256"]
    checked = verify_format_outputs(formats, root / formats["outputs"]["archive"])
    assert len(ledger["format_capabilities"]) == sum(len(formats["reports"][key]["records"])
                                                   for key in ("save_commercial", "load_commercial"))
    literal = read("commercial-26.9-literal-text.json")
    raw = (root / literal["fixtures"]["archive"]).read_bytes()
    assert digest(raw) == literal["fixtures"]["sha256"]
    outputs = (root / literal["raw_outputs"]["archive"]).read_bytes()
    assert digest(outputs) == literal["raw_outputs"]["sha256"]
    with ZipFile(BytesIO(raw)) as archive, ZipFile(BytesIO(outputs)) as generated:
        for backend in ("commercial", "docweave"):
            for row in literal[backend]:
                assert digest(archive.read(row["source"])) == row["source_sha256"]
                data = generated.read(row["output_member"])
                assert digest(data) == row["output_sha256"]
                assert data.decode().replace("\r\n", "\n") == row["markdown"]
    imports = verify_import_outputs(root)
    toggles = verify_style_toggles(root)
    return {"declared_symbols": len(symbols), "capability_rows": ledger["capability_count"],
            "checked_import_outputs": imports,
            "checked_style_inputs": toggles,
            "checked_format_outputs": checked, "behavioral_acceptance": False}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("root", nargs="?", type=Path,
                        default=Path(__file__).parents[1] / "docs" / "benchmarks")
    args = parser.parse_args()
    print(json.dumps(verify(args.root)))
