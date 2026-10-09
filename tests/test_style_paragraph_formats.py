"""Style alignment edits match owned official live/cold observations."""

import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from aspose.words_foss import DocxDocument

from .test_docx_dom import payloads
from .test_docx_dom_styles import style_doc
from .test_style_font_edits import snapshot

BENCHMARKS = Path(__file__).parents[1] / "docs/benchmarks"
REPORT = json.loads((BENCHMARKS / "style-paragraph-format-26.9.json").read_text())
NAMES = {"LEFT": "left", "CENTER": "center", "RIGHT": "right", "JUSTIFY": "both"}


def formats(document):
    value = snapshot(document)
    value["style_alignment"] = {name: {v: k for k, v in NAMES.items()}[document.styles.get_by_name(name).paragraph_format.alignment]
                                for name in ("Base", "Derived")}
    return value


@pytest.mark.parametrize("row", REPORT["records"], ids=lambda row: row["output"])
def test_style_alignment_live_edit_save_and_reopen(row, tmp_path):
    with ZipFile(BENCHMARKS / REPORT["corpus"]) as archive:
        raw = archive.read(row["input"])
    document = DocxDocument(BytesIO(raw))
    assert formats(document) == row["before_edit"]
    assert payloads(document.to_bytes()) == payloads(raw)
    style = document.styles.get_by_name("Derived")
    style.paragraph_format.alignment = NAMES[row["alignment"]]
    assert formats(document) == row["after_edit"]
    live = document.part_xml("word/styles.xml")
    saved = document.to_bytes()
    assert document.part_xml("word/styles.xml") == live
    assert formats(document) == row["after_save_live"]
    assert formats(DocxDocument(BytesIO(saved))) == row["after_reopen"]
    path = tmp_path / "saved.docx"
    document.save(path)
    assert path.read_bytes() == saved
    for name, data in payloads(raw).items():
        if name != "word/styles.xml":
            assert payloads(saved)[name] == data
    style.direct_paragraph_format.alignment = None
    assert style.direct_paragraph_format.alignment is None
    assert style.paragraph_format.alignment == document.styles.get_by_name("Base").paragraph_format.alignment
    before = document.to_bytes()
    with pytest.raises(TypeError):
        style.paragraph_format.alignment = None
    assert document.to_bytes() == before


@pytest.mark.parametrize("properties", ['<w:pPr/><w:pPr/>', '<w:pPr><w:jc w:val="left"/><w:jc w:val="right"/></w:pPr>'])
def test_duplicate_style_paragraph_properties_fail_atomically(tmp_path, properties):
    document = style_doc(tmp_path, '<w:style w:type="paragraph" w:styleId="S"><w:name w:val="S"/>' + properties + '</w:style>')
    before = document.to_bytes()
    with pytest.raises(ValueError, match="Duplicate"):
        document.styles.get_by_name("S").paragraph_format.alignment = "center"
    assert document.to_bytes() == before


def test_character_style_paragraph_format_is_absent(tmp_path):
    document = style_doc(tmp_path, '<w:style w:type="character" w:styleId="C"><w:name w:val="C"/></w:style>')
    assert document.styles.get_by_name("C").paragraph_format is None
    assert document.styles.get_by_name("C").direct_paragraph_format is None
