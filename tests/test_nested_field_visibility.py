"""Nested field results inside outer instructions are not visible output."""

import warnings
from io import BytesIO

import pytest
from pypdf import PdfReader

from ._word_fixtures import word_document

import dockiln as dk
from dockiln import light_document_model as ldm


def nested_paragraph():
    return ldm.Paragraph(children=[ldm.FieldStart(), ldm.Run(text="OUTER_CODE"),
        ldm.FieldStart(), ldm.Run(text="INNER_CODE"), ldm.FieldSeparator(),
        ldm.Run(text="SECRET_INNER_RESULT"), ldm.FieldEnd(), ldm.Run(text="SECRET_OUTER_CODE"),
        ldm.FieldSeparator(), ldm.Run(text="VISIBLE_RESULT"), ldm.FieldEnd()])


@pytest.mark.parametrize("mode", ["pdf_body", "pdf_table"])
def test_nested_instruction_results_do_not_leak(mode):
    doc = word_document(b"BODY")
    paragraph = nested_paragraph()
    doc.sections[0].body.children = [ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph])])])] if "table" in mode else [paragraph]
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", dk.ContentLossWarning)
        raw = doc.to_bytes("pdf")
    result = "".join(p.extract_text() for p in PdfReader(BytesIO(raw)).pages)
    assert "VISIBLE_RESULT" in result and "SECRET" not in result and "OUTER_CODE" not in result


@pytest.mark.parametrize("table", [False, True])
def test_note_references_inside_instruction_do_not_expose_note_body(table):
    doc = word_document(b"BODY")
    paragraph = ldm.Paragraph(children=[ldm.FieldStart(),
        ldm.NoteReference(kind="footnote", identifier="1"), ldm.FieldSeparator(),
        ldm.Run(text="VISIBLE_RESULT"), ldm.FieldEnd()])
    doc.sections[0].body.children = [ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph])])])] if table else [paragraph]
    doc.light_document_model.source_stories = [ldm.SourceStory(kind="footnote", part_name="word/footnotes.xml", identifier="1", children=[
        ldm.Paragraph(children=[ldm.Run(text="SECRET_NOTE")])])]
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", dk.ContentLossWarning)
        result = "".join(p.extract_text() for p in PdfReader(BytesIO(doc.to_bytes("pdf"))).pages)
    assert "SECRET_NOTE" not in result and "[^note" not in result


def test_nested_visible_results_remain_visible():
    doc = word_document(b"BODY")
    doc.sections[0].body.children = [ldm.Paragraph(children=[ldm.FieldStart(),
        ldm.Run(text="CODE"), ldm.FieldSeparator(), ldm.Run(text="BEFORE"),
        ldm.FieldStart(), ldm.Run(text="SECRET_INNER_CODE"), ldm.FieldSeparator(),
        ldm.Run(text="INNER_VISIBLE"), ldm.FieldEnd(), ldm.Run(text="AFTER"), ldm.FieldEnd()])]
    raw = doc.to_bytes("pdf")
    result = "".join(p.extract_text() for p in PdfReader(BytesIO(raw)).pages)
    assert all(token in result for token in ("BEFORE", "INNER_VISIBLE", "AFTER"))
    assert "CODE" not in result
