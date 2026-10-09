"""Replay real 26.9 style edits across live, saved, and reopened state."""

import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from aspose.words_foss import DocxDocument

from .test_docx_dom import payloads
from .test_docx_dom_styles import style_doc
from .test_style_import_corpus import VERIFY, live_formats

BENCHMARKS = Path(__file__).parents[1] / "docs/benchmarks"
ROWS = json.loads((BENCHMARKS / "style-font-edits-26.9.json").read_text())["records"]


def snapshot(document):
    styles = {}
    for name in ("Ancestor", "Base", "Derived", "P", "C"):
        style = document.styles.get_by_name(name)
        if style is not None:
            styles[name] = {key: getattr(style.font, key) for key in ("bold", "italic", "size")}
    return {"paragraphs": live_formats(document), "styles": styles}


@pytest.mark.parametrize("row", ROWS, ids=lambda row: row["output"])
def test_style_font_edit_matches_official_live_and_cold_formats(row, tmp_path):
    with ZipFile(BENCHMARKS / "corpus/style-import-conflicts-26.9.zip") as archive:
        original = archive.read(row["input"])
    document = DocxDocument(BytesIO(original))
    assert snapshot(document) == row["before_edit"]
    assert payloads(document.to_bytes()) == payloads(original)
    style = document.styles.get_by_name(row["style"])
    setattr(style.font, row["property"], row["value"])
    assert snapshot(document) == row["after_edit"]
    part = document._package._style_projection_part
    live_xml = document.part_xml(part)
    saved = document.to_bytes()
    assert document.part_xml(part) == live_xml
    assert snapshot(document) == row["after_save_live"]
    assert snapshot(DocxDocument(BytesIO(saved))) == row["after_reopen"]
    assert VERIFY["saved_style_fonts"](saved) == row["after_reopen"]["styles"]
    assert VERIFY["saved_story_formats"](saved) == row["after_reopen"]["paragraphs"]
    assert document.to_bytes() == saved
    path = tmp_path / "saved.docx"
    document.save(path)
    assert path.read_bytes() == saved
    for name, data in payloads(original).items():
        if name != part:
            assert payloads(saved)[name] == data


def test_direct_style_values_distinguish_unset_and_false():
    row = ROWS[0]
    with ZipFile(BENCHMARKS / "corpus/style-import-conflicts-26.9.zip") as archive:
        document = DocxDocument(BytesIO(archive.read(row["input"])))
    style = document.styles.get_by_name("Derived")
    style.direct_font.bold = None
    assert style.direct_font.bold is None
    assert style.font.bold == document.styles.get_by_name("Base").font.bold
    style.font.bold = False
    assert style.direct_font.bold is False
    before = document.to_bytes()
    with pytest.raises(TypeError):
        style.font.bold = None
    assert document.to_bytes() == before
    assert document.styles.get_by_id(style.style_id).name == "Derived"
    assert len(document.styles) == len(tuple(document.styles))
    assert document.styles.get_by_name("absent") is None


def test_style_handle_tracks_replaced_related_part_and_preserves_metadata(tmp_path):
    document = style_doc(tmp_path, (
        '<w:style w:type="paragraph" w:styleId="S"><w:name w:val="Custom"/>'
        '<w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b w:future="keep"/>'
        '<x:opaque xmlns:x="urn:opaque"/></w:rPr></w:style>'
    ), target="custom/styles.xml")
    # The relationship is authoritative, including a nonstandard styles location.
    document._package.set_parts({"word/custom/styles.xml": document._package.payload("word/styles.xml")})
    style = document.styles.get_by_name("Custom")
    assert style.document is document and style.owner_document is document
    original = document._package.payload("word/custom/styles.xml")
    document._package.set_parts({"word/custom/styles.xml": original.replace(b'<w:b w:future="keep"/>', b'<w:b w:val="0" w:future="keep"/>')})
    assert style.font.bold is False
    style.font.bold = True
    saved = payloads(document.to_bytes())["word/custom/styles.xml"]
    assert b'future="keep"' in saved and b'urn:opaque' in saved and b'val="center"' in saved
    assert document._package._style_projection_part == "word/custom/styles.xml"
    assert DocxDocument(BytesIO(document.to_bytes())).styles.get_by_name("Custom").font.bold is True


@pytest.mark.parametrize("properties", ['<w:rPr/><w:rPr/>', '<w:rPr><w:b/><w:b/></w:rPr>'])
def test_duplicate_style_properties_fail_without_mutation(tmp_path, properties):
    document = style_doc(tmp_path, '<w:style w:type="character" w:styleId="S"><w:name w:val="S"/>' + properties + '</w:style>')
    before = document.to_bytes()
    with pytest.raises(ValueError, match="Duplicate"):
        document.styles.get_by_name("S").font.bold = True
    assert document.to_bytes() == before
