"""Replay owned imports against the installed package and frozen official getters."""

import hashlib
import json
import runpy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss.dom.styles import _style_stories

from .test_docx_dom import payloads

ROOT = Path(__file__).parents[1]
BENCHMARKS = ROOT / "docs" / "benchmarks"
REPORT = json.loads((BENCHMARKS / "style-import-projections.json").read_text())
INSPECT = runpy.run_path(str(ROOT / "docs" / "probes" / "inspect_style_imports.py"))["inspect_document"]
VERIFY = runpy.run_path(str(ROOT / "scripts" / "verify_commercial_baseline.py"))


def live_formats(document):
    result = {}
    for root in _style_stories(document._package):
        part = next(name for name, tree in document._package._trees.items() if tree.documentElement is root)
        for paragraph in (document.body if part == "word/document.xml" else document.story(part)).paragraphs:
            if paragraph.text not in {"IMPORT", "DESTINATION"}:
                continue
            font = paragraph.runs[0].effective_font
            alignment = paragraph.effective_paragraph_format.alignment
            assert paragraph.text not in result
            result[paragraph.text] = {"bold": font.bold, "italic": font.italic, "size": font.size,
                                      "alignment": {"both": "JUSTIFY"}.get(alignment, (alignment or "left").upper())}
    return result


@pytest.mark.parametrize("matrix,row", [(matrix, row) for matrix in REPORT["reports"] for row in matrix["candidate"]["records"]],
                         ids=[row["case"] for matrix in REPORT["reports"] for row in matrix["candidate"]["records"]])
def test_style_import_live_and_projected_formats(matrix, row, tmp_path):
    with ZipFile(BENCHMARKS / matrix["corpus"]["archive"]) as archive:
        data = {phase: archive.read(row["case"] + "/" + phase + ".docx") for phase in ("source", "destination")}
    assert all(hashlib.sha256(value).hexdigest() == row["inputs"][phase] for phase, value in data.items())
    native_report = json.loads((BENCHMARKS / matrix["native"]["file"]).read_text())["reports"][matrix["native"]["report_index"]]
    native = next(item for item in native_report["records"] if item["case"] == row["case"])
    source, target = (aw.DocxDocument(BytesIO(data[phase])) for phase in ("source", "destination"))
    node = source.body.paragraphs[0]
    copied = target.import_node(node, True)
    assert copied.owner_document is target and copied.parent_node is None
    target.body.append_child(copied)
    assert live_formats(target) == native["before_save"]["paragraphs"]
    part = target._package._style_projection_part
    live_style_xml = target.part_xml(part)
    saved = target.to_bytes()
    assert live_formats(target) == native["after_save_live"]["paragraphs"]
    assert target.part_xml(part) == live_style_xml
    assert target.to_bytes() == saved
    path = tmp_path / "saved.docx"
    target.save(path)
    assert path.read_bytes() == saved
    assert VERIFY["saved_story_formats"](saved) == native["after_reopen"]["paragraphs"]
    assert VERIFY["saved_style_fonts"](saved) == native["after_reopen"]["styles"]
    expected = next(item for item in matrix["independent_checks"] if item["case"] == row["case"])
    assert INSPECT(BytesIO(saved)) == {key: value for key, value in expected.items() if key != "case"}
    assert node.parent_node is source.body and payloads(source.to_bytes()) == payloads(data["source"])
