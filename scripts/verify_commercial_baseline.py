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


def verify_font_defaults(root):
    report = json.loads((root / "font-defaults-26.9.json").read_text())
    raw = (root / report["corpus"]).read_bytes()
    assert report["version"] == "26.9.0" and report["licensed"] is False
    assert report["full_format_acceptance"] is False and digest(raw) == report["corpus_sha256"]
    generated = dict(runpy.run_path(str(Path(__file__).parents[1] / "docs" / "probes" / "font_defaults.py"))["inputs"]())
    rows = {row["input"]: row for row in report["records"]}
    assert len(rows) == len(report["records"]) == len(generated) == 5 and set(rows) == set(generated)
    with ZipFile(BytesIO(raw)) as archive:
        assert len(archive.namelist()) == 5 and set(archive.namelist()) == set(rows)
        for name, row in rows.items():
            data = archive.read(name)
            assert digest(data) == row["sha256"]
            with ZipFile(BytesIO(data)) as actual, ZipFile(BytesIO(generated[name])) as expected:
                assert {key: actual.read(key) for key in actual.namelist()} == {key: expected.read(key) for key in expected.namelist()}
            assert row["run_size"] == row["style_size"] == (10.0 if name.startswith("rpr_") else 11.0)
    return len(rows)


def verify_font_default_matrix(root):
    report = json.loads((root / "font-default-matrix-26.9.json").read_text())
    assert report["version"] == "26.9.0" and report["licensed"] is False
    assert report["full_format_acceptance"] is report["rendering_acceptance"] is False
    raw = (root / report["corpus"]).read_bytes()
    assert digest(raw) == report["corpus_sha256"]
    generated = dict(runpy.run_path(str(Path(__file__).parents[1] / "docs/probes/font_default_matrix.py"))["inputs"]())
    rows = {row["input"]: row for row in report["records"]}
    assert len(rows) == len(report["records"]) == len(generated) == 16
    with ZipFile(BytesIO(raw)) as archive:
        assert len(archive.namelist()) == 16 and set(archive.namelist()) == set(rows) == set(generated)
        for name, row in rows.items():
            data = archive.read(name)
            assert digest(data) == row["sha256"]
            with ZipFile(BytesIO(data)) as actual, ZipFile(BytesIO(generated[name])) as expected:
                assert {key: actual.read(key) for key in actual.namelist()} == {key: expected.read(key) for key in expected.namelist()}
            size = 12.0 if name.endswith("-run_size.docx") else 10.0 if "-run_" in name else 11.0
            assert row["run_size"] == row["style_size"] == size
    return len(rows)


def verify_style_imports(root, name="style-import-conflicts-26.9.json"):
    report = json.loads((root / name).read_text())
    assert set(report["reports"]) in ({"commercial", "before", "candidate"}, {"candidate"})
    assert set(report["official_candidate_rereads"]) == set(report["reports"]) - {"commercial"}
    corpus = (root / report["corpus"]["archive"]).read_bytes()
    raw_outputs = (root / report["outputs"]["archive"]).read_bytes()
    assert digest(corpus) == report["corpus"]["sha256"], "style import corpus digest mismatch"
    assert digest(raw_outputs) == report["outputs"]["sha256"], "style import outputs digest mismatch"
    probes = Path(__file__).parents[1] / "docs" / "probes"
    generator = report["corpus"].get("generator", "import_style_conflicts.py")
    counts = {"import_style_conflicts.py": 247, "paragraph_style_defaults.py": 132}
    assert generator in counts
    generate = runpy.run_path(str(probes / generator))["inputs"]
    expected = {key + "/" + phase + ".docx" for key, phase, _ in generate()}
    keys = {name.split("/")[0] for name in expected}
    inspect = runpy.run_path(str(probes / "inspect_style_imports.py"))["inspect_document"]
    assert len(expected) == 2 * len(keys) and len(keys) == report["corpus"]["input_pairs"] == counts[generator]
    checked, phases, formats = 0, {}, {}
    with ZipFile(BytesIO(corpus)) as inputs, ZipFile(BytesIO(raw_outputs)) as outputs:
        assert len(inputs.namelist()) == len(expected) and set(inputs.namelist()) == expected
        assert len(outputs.namelist()) == report["outputs"]["files"] == len(keys) * len(report["reports"])
        for phase in report["reports"]:
            observed = report["reports"][phase]
            assert observed["corpus_sha256"] == report["corpus"]["sha256"]
            rows = {row["case"]: row for row in observed["records"]}
            assert len(rows) == len(observed["records"]) == len(keys) and set(rows) == keys
            phases[phase] = rows
            recorded = {row["case"]: row for row in report["independent_checks"][phase]}
            assert len(recorded) == len(keys) and set(recorded) == keys
            formats[phase] = {}
            for key, row in rows.items():
                for source in ("source", "destination"):
                    assert digest(inputs.read(key + "/" + source + ".docx")) == row["inputs"][source]
                data = outputs.read(phase + "/" + row["output"])
                assert digest(data) == row["output_sha256"]
                actual = {"case": key, "output": row["output"], **inspect(BytesIO(data))}
                assert actual == recorded[key], "style import independent observation mismatch"
                if "story_rereads" in report:
                    formats[phase][key] = saved_story_formats(data)
                if row["outcome"] == "raised":
                    assert phase != "commercial" and row["destination_unchanged"] is True
                else:
                    assert row["outcome"] == "returned" and "reopened_format" in row
                checked += 1
    if "commercial" in phases:
        native = phases["commercial"]
    else:
        baseline = (root / report["baseline"]["file"]).read_bytes()
        assert digest(baseline.replace(b"\r\n", b"\n")) == report["baseline"]["sha256"]
        native = {row["case"]: row for row in json.loads(baseline)["reports"]["commercial"]["records"]}
    for phase in report["official_candidate_rereads"]:
        read = report["official_candidate_rereads"][phase]
        assert read["version"] == "26.9.0" and read["licensed"] is False
        observed = {row["case"]: row for row in read["records"]}
        returned = {key for key, row in phases[phase].items() if row["outcome"] == "returned"}
        assert len(observed) == len(read["records"]) and set(observed) == returned
        assert all(row["output_sha256"] == phases[phase][key]["output_sha256"] for key, row in observed.items())
        mismatches = sorted(key for key, row in observed.items() if row["format"] != native[key]["reopened_format"])
        assert mismatches == sorted(report["reread_mismatches"][phase])
        if phase == "candidate":
            assert not mismatches
    assert sorted(report["native_source_changes"]) == sorted(key for key, row in native.items() if row["source_format"] != row["imported_format"])
    assert sorted(report["native_roundtrip_changes"]) == sorted(key for key, row in native.items() if row["imported_format"] != row["reopened_format"])
    repeated = {row["case"]: row for row in report["repeat_without_source_getter"]}
    assert len(repeated) == len(report["repeat_without_source_getter"]) and set(repeated) == set(report["native_roundtrip_changes"])
    for key, row in repeated.items():
        assert row["no_source_getter_before_import"] is True
        if "imported_format" in row:
            assert row["imported_format"] == native[key]["imported_format"]
            assert row["reopened_format"] == native[key]["reopened_format"]
        else:
            assert row["imported_bold"] == native[key]["imported_format"]["bold"]
            assert row["reopened_bold"] == native[key]["reopened_format"]["bold"]
    if "repeat_outputs" in report:
        raw = (root / report["repeat_outputs"]["archive"]).read_bytes()
        assert digest(raw) == report["repeat_outputs"]["sha256"]
        with ZipFile(BytesIO(raw)) as archive:
            assert len(archive.namelist()) == report["repeat_outputs"]["files"] == len(repeated)
            for row in repeated.values():
                data = archive.read(row["output"])
                assert digest(data) == row["output_sha256"]
                assert saved_story_formats(data)["IMPORT"] == row["reopened_format"]
    if "story_rereads" in report:
        previous = phases.get("before")
        if previous is None:
            path = root / report["before_baseline"]["file"]
            raw = path.read_bytes()
            assert digest(raw.replace(b"\r\n", b"\n")) == report["before_baseline"]["sha256"]
            previous = {row["case"]: row for row in json.loads(raw)["reports"]["candidate"]["records"]}
        source_rows = {"commercial": native, "before": previous, "candidate": phases["candidate"]}
        observations = {}
        assert set(report["story_rereads"]) == set(source_rows)
        for phase, observed in report["story_rereads"].items():
            assert observed["version"] == "26.9.0" and observed["licensed"] is False
            rows = {row["case"]: row for row in observed["records"]}
            assert len(rows) == len(observed["records"]) == len(keys) and set(rows) == keys
            observations[phase] = rows
            for key, row in rows.items():
                original = source_rows[phase][key]
                assert row["output_sha256"] == original["output_sha256"]
                assert set(row["formats"]) == ({"IMPORT", "DESTINATION"} if original["outcome"] == "returned" else {"DESTINATION"})
                if phase in formats:
                    assert row["formats"] == formats[phase][key], "saved story observation mismatch"
        for phase in ("before", "candidate"):
            mismatches = sorted(key for key, row in source_rows[phase].items() if row["outcome"] == "returned"
                                and observations[phase][key]["formats"] != observations["commercial"][key]["formats"])
            assert mismatches == sorted(report["story_mismatches"][phase])
        assert not report["story_mismatches"]["candidate"]
    assert report["validation"]["full_import_acceptance"] is False and report["validation"]["rendering_acceptance"] is False
    if "style_state_baseline" in report:
        raw = (root / report["style_state_baseline"]["file"]).read_bytes()
        assert digest(raw.replace(b"\r\n", b"\n")) == report["style_state_baseline"]["sha256"]
        native_styles = {row["case"]: row["after_reopen"]["styles"] for row in json.loads(raw)["reports"][0]["records"]}
        with ZipFile(BytesIO(raw_outputs)) as outputs:
            for key, row in phases["candidate"].items():
                if row["outcome"] == "returned":
                    assert saved_style_fonts(outputs.read("candidate/" + row["output"])) == native_styles[key], "saved style font mismatch"
    return checked


def saved_story_formats(data):
    from aspose.words_foss import DocxDocument
    from aspose.words_foss.dom.styles import _style_stories

    document = DocxDocument(BytesIO(data))
    paragraphs = []
    for root in _style_stories(document._package):
        part = next(name for name, tree in document._package._trees.items() if tree.documentElement is root)
        paragraphs.extend((document.body if part == "word/document.xml" else document.story(part)).paragraphs)
    result = {}
    for paragraph in paragraphs:
        if paragraph.text not in {"IMPORT", "DESTINATION"}:
            continue
        assert paragraph.text not in result
        font = paragraph.runs[0].effective_font
        alignment = paragraph.effective_paragraph_format.alignment
        result[paragraph.text] = {"bold": font.bold, "italic": font.italic, "size": font.size,
                                  "alignment": {"both": "JUSTIFY"}.get(alignment, (alignment or "left").upper())}
    return result


def verify_style_roundtrips(root):
    report = json.loads((root / "style-save-roundtrips-26.9.json").read_text())
    observed = report["report"]
    assert observed["version"] == "26.9.0" and observed["licensed"] is False
    corpus = (root / report["corpus"]["archive"]).read_bytes()
    raw = (root / report["outputs"]["archive"]).read_bytes()
    assert digest(corpus) == report["corpus"]["sha256"] == observed["corpus_sha256"]
    assert digest(raw) == report["outputs"]["sha256"]
    rows = {(row["case"], row["phase"]): row for row in observed["records"]}
    with ZipFile(BytesIO(corpus)) as inputs, ZipFile(BytesIO(raw)) as outputs:
        assert len(rows) == len(observed["records"]) == len(inputs.namelist()) == len(outputs.namelist()) == 494
        assert {key + "/" + phase + ".docx" for key, phase in rows} == set(inputs.namelist())
        for (key, phase), row in rows.items():
            data = inputs.read(key + "/" + phase + ".docx")
            output = outputs.read(row["output"])
            label = "IMPORT" if phase == "source" else "DESTINATION"
            assert digest(data) == row["input_sha256"] and digest(output) == row["output_sha256"]
            assert saved_story_formats(data) == {label: row["before"]}
            assert saved_story_formats(output) == {label: row["after"]}
    assert report["changes"] == sorted(key + "/" + phase for (key, phase), row in rows.items() if row["before"] != row["after"])
    assert report["full_format_acceptance"] is False and report["rendering_acceptance"] is False
    return len(rows)


def saved_style_fonts(data):
    from docx import Document
    from docx.oxml.ns import qn

    document = Document(BytesIO(data))
    defaults = document.styles.element.find(qn("w:docDefaults"))
    properties = defaults.find(qn("w:rPrDefault")).find(qn("w:rPr"))
    values = {}
    for name, tag in (("bold", "b"), ("italic", "i"), ("size", "sz")):
        element = properties.find(qn("w:" + tag))
        raw = None if element is None else element.get(qn("w:val"))
        values[name] = float(raw) / 2 if name == "size" else element is not None and raw not in {"0", "false", "off"}
    result = {}
    for name in ("Ancestor", "Base", "Derived", "P", "C"):
        if name not in document.styles:
            continue
        style, chain = document.styles[name], []
        while style is not None:
            assert style.style_id not in {item.style_id for item in chain}, "cyclic owned style"
            chain.append(style)
            style = style.base_style
        font = values.copy()
        for style in reversed(chain):
            for attribute in font:
                value = getattr(style.font, attribute)
                if value is not None:
                    font[attribute] = value.pt if attribute == "size" else value
        result[name] = font
    return result


def verify_style_save_state(root, name="style-save-state-26.9.json"):
    counts = {"style-save-state-26.9.json": (247, 132), "character-style-save-state-26.9.json": (108,),
              "style-normalization-contexts-26.9.json": (36,)}
    report = json.loads((root / name).read_text())
    raw = (root / report["outputs"]["archive"]).read_bytes()
    assert digest(raw) == report["outputs"]["sha256"]
    changes, members = [], set()
    with ZipFile(BytesIO(raw)) as outputs:
        for expected, observed in zip(counts[name], report["reports"], strict=True):
            assert observed["version"] == "26.9.0" and observed["licensed"] is False
            corpus = (root / observed["corpus"]).read_bytes()
            assert digest(corpus) == observed["corpus_sha256"]
            rows = {row["case"]: row for row in observed["records"]}
            assert len(rows) == len(observed["records"]) == expected
            with ZipFile(BytesIO(corpus)) as inputs:
                assert set(inputs.namelist()) == {key + "/" + phase + ".docx" for key in rows for phase in ("source", "destination")}
                for key, row in rows.items():
                    for phase in ("source", "destination"):
                        assert digest(inputs.read(key + "/" + phase + ".docx")) == row["inputs"][phase]
                    assert row["before_save"] == row["after_save_live"], "live save observation changed"
                    data = outputs.read(row["output"])
                    assert digest(data) == row["output_sha256"]
                    assert saved_story_formats(data) == row["after_reopen"]["paragraphs"], "reopened paragraph observation"
                    assert saved_style_fonts(data) == row["after_reopen"]["styles"], "reopened style observation"
                    assert row["before_save"]["styles"] == row["after_reopen"]["styles"]
                    if row["before_save"]["paragraphs"] != row["after_reopen"]["paragraphs"]:
                        changes.append(str(expected) + "/" + key)
                    assert row["output"] not in members
                    members.add(row["output"])
        assert members == set(outputs.namelist())
    assert sorted(changes) == report["reopen_paragraph_changes"]
    assert report["live_changes"] == report["reopen_style_changes"] == []
    assert report["full_format_acceptance"] is report["rendering_acceptance"] is False
    if "candidate" in report:
        native = {row["case"]: row for row in report["reports"][0]["records"]}
        candidate = {row["case"]: row for row in report["candidate"]["records"]}
        reread = report["candidate_official_rereads"]
        assert reread["version"] == "26.9.0" and reread["licensed"] is False
        actual = {row["case"]: row for row in reread["records"]}
        assert len(actual) == len(reread["records"]) == len(candidate) == len(report["candidate"]["records"]) == 108
        assert set(native) == set(candidate) == set(actual)
        assert report["candidate"]["corpus_sha256"] == report["reports"][0]["corpus_sha256"]
        for key, row in candidate.items():
            assert row["inputs"] == native[key]["inputs"]
            assert row["output_sha256"] == actual[key]["output_sha256"]
            if row["outcome"] == "returned":
                assert actual[key]["formats"] == native[key]["after_reopen"]["paragraphs"]
            else:
                assert row["outcome"] == "raised" and row["destination_unchanged"] is True
    return len(members)


def saved_style_alignments(data):
    from docx import Document
    from docx.oxml.ns import qn

    document = Document(BytesIO(data))
    default = document.styles.element
    for tag in ("docDefaults", "pPrDefault", "pPr", "jc"):
        default = default.find(qn("w:" + tag)) if default is not None else None
    value = "left" if default is None else default.get(qn("w:val"))
    result = {}
    for name in ("Base", "Derived"):
        style, chain = document.styles[name], []
        while style is not None:
            assert style.style_id not in {item.style_id for item in chain}
            chain.append(style)
            style = style.base_style
        alignment = {"left": "LEFT", "center": "CENTER", "right": "RIGHT", "both": "JUSTIFY"}[value]
        for style in reversed(chain):
            if style.paragraph_format.alignment is not None:
                alignment = style.paragraph_format.alignment.name
        result[name] = alignment
    return result


def verify_style_paragraph_formats(root):
    report = json.loads((root / "style-paragraph-format-26.9.json").read_text())
    assert report["version"] == "26.9.0" and report["licensed"] is False
    assert report["full_format_acceptance"] is report["rendering_acceptance"] is False
    source = (root / report["corpus"]).read_bytes()
    raw = (root / report["outputs"]).read_bytes()
    assert digest(source) == report["corpus_sha256"] and digest(raw) == report["outputs_sha256"]
    generated = dict(runpy.run_path(str(Path(__file__).parents[1] / "docs/probes/style_paragraph_formats.py"))["inputs"]())
    rows = report["records"]
    assert len(rows) == 108 and len(generated) == 27
    assert {(row["input"], row["alignment"]) for row in rows} == {
        (name, alignment) for name in generated for alignment in ("LEFT", "CENTER", "RIGHT", "JUSTIFY")}
    with ZipFile(BytesIO(source)) as inputs, ZipFile(BytesIO(raw)) as outputs:
        assert set(inputs.namelist()) == set(generated)
        assert set(outputs.namelist()) == {row["output"] for row in rows}
        for row in rows:
            data = inputs.read(row["input"])
            assert digest(data) == row["input_sha256"]
            with ZipFile(BytesIO(data)) as actual, ZipFile(BytesIO(generated[row["input"]])) as expected:
                assert {key: actual.read(key) for key in actual.namelist()} == {key: expected.read(key) for key in expected.namelist()}
            assert saved_style_alignments(data) == row["before_edit"]["style_alignment"]
            assert row["after_edit"]["style_alignment"]["Derived"] == row["alignment"]
            assert row["after_edit"] == row["after_save_live"] == row["after_reopen"]
            saved = outputs.read(row["output"])
            assert digest(saved) == row["output_sha256"]
            assert saved_style_alignments(saved) == row["after_reopen"]["style_alignment"]
            assert saved_style_fonts(saved) == row["after_reopen"]["styles"]
            assert saved_story_formats(saved) == row["after_reopen"]["paragraphs"]
    return len(rows)


def verify_style_font_edits(root):
    report = json.loads((root / "style-font-edits-26.9.json").read_text())
    assert report["version"] == "26.9.0" and report["licensed"] is False
    assert report["full_format_acceptance"] is report["rendering_acceptance"] is False
    source = (root / "corpus" / report["corpus"]).read_bytes()
    raw = (root / "corpus/style-font-edits-26.9.zip").read_bytes()
    assert digest(source) == report["corpus_sha256"]
    assert digest(raw) == report["outputs_sha256"]
    rows = report["records"]
    assert len(rows) == 1482
    with ZipFile(BytesIO(source)) as inputs, ZipFile(BytesIO(raw)) as outputs:
        expected = {(name, prop, value) for name in inputs.namelist() if name.endswith("/source.docx")
                    for prop, values in (("bold", (False, True)), ("italic", (False, True)), ("size", (10, 17.5)))
                    for value in values}
        assert {(row["input"], row["property"], row["value"]) for row in rows} == expected
        assert set(outputs.namelist()) == {row["output"] for row in rows}
        for row in rows:
            assert row["style"] == "Derived"
            assert digest(inputs.read(row["input"])) == row["input_sha256"]
            assert row["before_edit"]["styles"] == saved_style_fonts(inputs.read(row["input"]))
            assert row["after_edit"]["styles"]["Derived"][row["property"]] == row["value"]
            assert row["after_edit"] == row["after_save_live"]
            data = outputs.read(row["output"])
            assert digest(data) == row["output_sha256"]
            assert saved_story_formats(data) == row["after_reopen"]["paragraphs"]
            assert saved_style_fonts(data) == row["after_reopen"]["styles"]
            assert row["after_edit"]["styles"] == row["after_reopen"]["styles"]
    assert sum(row["after_edit"] != row["after_reopen"] for row in rows) == 36
    return len(rows)


def verify_style_projections(root):
    report = json.loads((root / "style-import-projections.json").read_text())
    assert report["baseline_version"] == "26.9.0"
    raw = (root / report["outputs"]["archive"]).read_bytes()
    assert digest(raw) == report["outputs"]["sha256"]
    inspect = runpy.run_path(str(Path(__file__).parents[1] / "docs" / "probes" / "inspect_style_imports.py"))["inspect_document"]
    members = set()
    with ZipFile(BytesIO(raw)) as outputs:
        for count, matrix in zip((247, 132, 108, 36), report["reports"], strict=True):
            assert matrix["count"] == count
            baseline = (root / matrix["native"]["file"]).read_bytes()
            assert digest(baseline.replace(b"\r\n", b"\n")) == matrix["native"]["sha256"]
            native = {row["case"]: row for row in json.loads(baseline)["reports"][matrix["native"]["report_index"]]["records"]}
            corpus = (root / matrix["corpus"]["archive"]).read_bytes()
            assert digest(corpus) == matrix["corpus"]["sha256"] == matrix["candidate"]["corpus_sha256"]
            rows = {row["case"]: row for row in matrix["candidate"]["records"]}
            reread = matrix["official_reread"]
            assert reread["version"] == "26.9.0" and reread["licensed"] is False
            observed = {row["case"]: row for row in reread["records"]}
            checks = {row["case"]: row for row in matrix["independent_checks"]}
            assert len(rows) == len(matrix["candidate"]["records"]) == len(native) == len(observed) == len(reread["records"]) == len(checks) == len(matrix["independent_checks"]) == count
            assert set(rows) == set(native) == set(observed) == set(checks)
            with ZipFile(BytesIO(corpus)) as inputs:
                assert set(inputs.namelist()) == {key + "/" + phase + ".docx" for key in rows for phase in ("source", "destination")}
                for key, row in rows.items():
                    assert row["outcome"] == "returned"
                    assert all(digest(inputs.read(key + "/" + phase + ".docx")) == row["inputs"][phase] == native[key]["inputs"][phase]
                               for phase in ("source", "destination"))
                    member = str(count) + "/" + row["output"]
                    data = outputs.read(member)
                    assert digest(data) == row["output_sha256"] == observed[key]["output_sha256"]
                    assert saved_story_formats(data) == observed[key]["formats"] == native[key]["after_reopen"]["paragraphs"]
                    assert saved_style_fonts(data) == native[key]["after_reopen"]["styles"]
                    warm = row["imported_format"]
                    warm = {**warm, "alignment": {"both": "JUSTIFY"}.get(warm["alignment"], (warm["alignment"] or "left").upper())}
                    assert warm == native[key]["before_save"]["paragraphs"]["IMPORT"]
                    assert {"case": key, **inspect(BytesIO(data))} == checks[key]
                    assert member not in members
                    members.add(member)
        assert set(outputs.namelist()) == members and len(members) == report["outputs"]["files"] == 523
    assert report["validation"]["full_import_acceptance"] is report["validation"]["full_format_acceptance"] is report["validation"]["rendering_acceptance"] is False
    return len(members)


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
    style_imports = verify_style_imports(root)
    style_imports += verify_style_imports(root, "style-import-translated.json")
    style_imports += verify_style_imports(root, "style-import-default-on.json")
    style_imports += verify_style_imports(root, "style-import-character-default-on.json")
    style_imports += verify_style_imports(root, "paragraph-style-defaults-26.9.json")
    roundtrips = verify_style_roundtrips(root)
    save_states = verify_style_save_state(root)
    save_states += verify_style_save_state(root, "character-style-save-state-26.9.json")
    save_states += verify_style_save_state(root, "style-normalization-contexts-26.9.json")
    projections = verify_style_projections(root)
    edits = verify_style_font_edits(root)
    paragraph_edits = verify_style_paragraph_formats(root)
    defaults = verify_font_defaults(root)
    default_matrix = verify_font_default_matrix(root)
    return {"declared_symbols": len(symbols), "capability_rows": ledger["capability_count"],
            "checked_import_outputs": imports,
            "checked_style_inputs": toggles,
            "checked_style_import_outputs": style_imports,
            "checked_font_default_inputs": defaults,
            "checked_font_default_matrix_inputs": default_matrix,
            "checked_style_roundtrip_outputs": roundtrips,
            "checked_style_save_states": save_states,
            "checked_style_projection_outputs": projections,
            "checked_style_font_edit_outputs": edits,
            "checked_style_paragraph_edit_outputs": paragraph_edits,
            "checked_format_outputs": checked, "behavioral_acceptance": False}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("root", nargs="?", type=Path,
                        default=Path(__file__).parents[1] / "docs" / "benchmarks")
    args = parser.parse_args()
    print(json.dumps(verify(args.root)))
