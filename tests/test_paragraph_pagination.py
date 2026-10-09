"""Replay official pagination property getters and editable style/direct layers."""

import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from aspose.words_foss import DocxDocument
from aspose.words_foss.light_document_model import NodeType

from .test_docx_dom import payloads
from .test_docx_dom_styles import style_doc

BENCHMARKS = Path(__file__).parents[1] / "docs/benchmarks"
REPORT = json.loads((BENCHMARKS / "paragraph-pagination-26.9.json").read_text())


def snapshot(document, prop):
    paragraph = next(node for node in document.get_child_nodes(NodeType.PARAGRAPH, True) if "IMPORT" in node.text)
    return {"base": getattr(document.styles.get_by_name("Base").paragraph_format, prop),
            "derived": getattr(document.styles.get_by_name("Derived").paragraph_format, prop),
            "paragraph": getattr(paragraph.effective_paragraph_format, prop)}


@pytest.mark.parametrize("row", REPORT["records"], ids=lambda row: row["output"])
def test_pagination_edit_roundtrip(row, tmp_path):
    with ZipFile(BENCHMARKS / REPORT["corpus"]) as archive:
        raw = archive.read(row["input"])
    document = DocxDocument(BytesIO(raw))
    prop = row["property"]
    assert snapshot(document, prop) == row["before"]
    assert payloads(document.to_bytes()) == payloads(raw)
    style = document.styles.get_by_name("Derived")
    paragraph = next(node for node in document.get_child_nodes(NodeType.PARAGRAPH, True) if "IMPORT" in node.text)
    node = style if row["target"] == "style" else paragraph
    setattr(node.paragraph_format, prop, row["value"])
    assert snapshot(document, prop) == row["edited"]
    saved = document.to_bytes()
    assert snapshot(document, prop) == row["saved_live"]
    assert snapshot(DocxDocument(BytesIO(saved)), prop) == row["reopened"]
    layout = next(p for p in document.to_light_document().all_paragraphs if "IMPORT" in p.text)
    assert getattr(layout.paragraph_format, prop) == row["edited"]["paragraph"]
    path = tmp_path / "saved.docx"
    document.save(path)
    assert path.read_bytes() == saved
    changed = "word/styles.xml" if row["target"] == "style" else "word/document.xml"
    for name, data in payloads(raw).items():
        if name != changed:
            assert payloads(saved)[name] == data
    direct = style.direct_paragraph_format if row["target"] == "style" else paragraph.paragraph_format
    setattr(direct, prop, None)
    assert getattr(direct, prop) is None
    inherited = getattr(document.styles.get_by_name("Base").paragraph_format, prop) if row["target"] == "style" else getattr(style.paragraph_format, prop)
    assert snapshot(document, prop)["derived" if row["target"] == "style" else "paragraph"] == inherited
    before = document.to_bytes()
    with pytest.raises((TypeError, ValueError)):
        setattr(node.paragraph_format, prop, "invalid")
    assert document.to_bytes() == before


@pytest.mark.parametrize("target", ["style", "paragraph"])
@pytest.mark.parametrize("tag,prop", [("keepNext", "keep_with_next"), ("keepLines", "keep_together"),
                                      ("pageBreakBefore", "page_break_before"), ("widowControl", "widow_control")])
def test_pagination_duplicate_and_invalid_properties_fail_atomically(tmp_path, target, tag, prop):
    document = style_doc(tmp_path, '<w:style w:type="paragraph" w:styleId="S"><w:name w:val="S"/>'
                         f'<w:pPr><w:{tag}/><w:{tag} w:val="0"/></w:pPr></w:style>',
                         f'<w:{tag}/><w:{tag} w:val="0"/>')
    node = document.styles.get_by_name("S") if target == "style" else document.body.paragraphs[0]
    before = document.to_bytes()
    with pytest.raises(ValueError, match="Duplicate"):
        setattr(node.paragraph_format, prop, True)
    assert document.to_bytes() == before
    document = style_doc(tmp_path, '<w:style w:type="paragraph" w:styleId="S"><w:name w:val="S"/>'
                         f'<w:pPr><w:{tag} w:val="invalid"/></w:pPr></w:style>', f'<w:{tag} w:val="invalid"/>')
    node = document.styles.get_by_name("S") if target == "style" else document.body.paragraphs[0]
    before = document.to_bytes()
    with pytest.raises(ValueError, match="on/off"):
        getattr(node.paragraph_format, prop)
    assert document.to_bytes() == before


def test_first_paragraph_page_break_preserves_explicit_input():
    report = json.loads((BENCHMARKS / "first-paragraph-page-break-26.9.json").read_text())
    row = report["records"][0]
    with ZipFile(BENCHMARKS / report["corpus"]) as archive:
        raw = archive.read(row["input"])
    document = DocxDocument(BytesIO(raw))
    assert document.body.paragraphs[0].paragraph_format.page_break_before is True
    assert document.body.paragraphs[0].effective_paragraph_format.page_break_before is True
    assert payloads(document.to_bytes()) == payloads(raw)


@pytest.mark.parametrize("prop", ["keep_with_next", "keep_together", "page_break_before", "widow_control"])
@pytest.mark.parametrize("value", [None, 1, "bad"])
def test_style_pagination_setter_errors_match_official(tmp_path, prop, value):
    record = next(row for row in REPORT["setter_errors"] if row["target"] == "style" and row["property"] == prop and row["value"] == value)
    assert record["error"] == "TypeError"
    document = style_doc(tmp_path, '<w:style w:type="paragraph" w:styleId="S"><w:name w:val="S"/></w:style>')
    before = document.to_bytes()
    with pytest.raises(TypeError):
        setattr(document.styles.get_by_name("S").paragraph_format, prop, value)
    assert document.to_bytes() == before
