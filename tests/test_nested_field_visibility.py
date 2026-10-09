"""Nested field results inside outer instructions are not visible output."""

import warnings
from io import BytesIO

import pytest
from pypdf import PdfReader

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm


def nested_paragraph():
    return ldm.Paragraph(children=[ldm.FieldStart(), ldm.Run(text="OUTER_CODE"),
        ldm.FieldStart(), ldm.Run(text="INNER_CODE"), ldm.FieldSeparator(),
        ldm.Run(text="SECRET_INNER_RESULT"), ldm.FieldEnd(), ldm.Run(text="SECRET_OUTER_CODE"),
        ldm.FieldSeparator(), ldm.Run(text="VISIBLE_RESULT"), ldm.FieldEnd()])


@pytest.mark.parametrize("mode", ["markdown_body", "markdown_table", "html_table", "pdf_body", "pdf_table"])
def test_nested_instruction_results_do_not_leak(mode):
    doc = aw.Document(BytesIO(b"BODY"), aw.MarkdownLoadOptions())
    paragraph = nested_paragraph()
    doc.sections[0].body.children = [ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph])])])] if "table" in mode else [paragraph]
    options = aw.saving.MarkdownSaveOptions()
    if mode == "html_table":
        options.export_as_html = aw.saving.MarkdownExportAsHtml.TABLES
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", aw.ContentLossWarning)
        raw = doc.to_bytes("pdf" if mode.startswith("pdf") else options)
    result = "".join(p.extract_text() for p in PdfReader(BytesIO(raw)).pages) if mode.startswith("pdf") else raw.decode()
    assert "VISIBLE_RESULT" in result and "SECRET" not in result and "OUTER_CODE" not in result


@pytest.mark.parametrize("table", [False, True])
def test_note_references_inside_instruction_do_not_expose_note_body(table):
    doc = aw.Document(BytesIO(b"BODY"), aw.MarkdownLoadOptions())
    paragraph = ldm.Paragraph(children=[ldm.FieldStart(),
        ldm.NoteReference(kind="footnote", identifier="1"), ldm.FieldSeparator(),
        ldm.Run(text="VISIBLE_RESULT"), ldm.FieldEnd()])
    doc.sections[0].body.children = [ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph])])])] if table else [paragraph]
    doc.light_document_model.source_stories = [ldm.SourceStory(kind="footnote", part_name="word/footnotes.xml", identifier="1", children=[
        ldm.Paragraph(children=[ldm.Run(text="SECRET_NOTE")])])]
    options = aw.saving.MarkdownSaveOptions()
    options.export_notes = True
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", aw.ContentLossWarning)
        result = doc.to_bytes(options).decode()
    assert "SECRET_NOTE" not in result and "[^note" not in result


@pytest.mark.parametrize("mode", ["pdf", "md"])
def test_nested_visible_results_remain_visible(mode):
    doc = aw.Document(BytesIO(b"BODY"), aw.MarkdownLoadOptions())
    doc.sections[0].body.children = [ldm.Paragraph(children=[ldm.FieldStart(),
        ldm.Run(text="CODE"), ldm.FieldSeparator(), ldm.Run(text="BEFORE"),
        ldm.FieldStart(), ldm.Run(text="SECRET_INNER_CODE"), ldm.FieldSeparator(),
        ldm.Run(text="INNER_VISIBLE"), ldm.FieldEnd(), ldm.Run(text="AFTER"), ldm.FieldEnd()])]
    raw = doc.to_bytes(mode)
    result = "".join(p.extract_text() for p in PdfReader(BytesIO(raw)).pages) if mode == "pdf" else raw.decode()
    assert all(token in result for token in ("BEFORE", "INNER_VISIBLE", "AFTER"))
    assert "CODE" not in result


def test_nested_result_in_hyperlink_instruction_builds_target_without_becoming_label():
    doc = aw.Document(BytesIO(b"BODY"), aw.MarkdownLoadOptions())
    doc.sections[0].body.children = [ldm.Paragraph(children=[ldm.FieldStart(),
        ldm.Run(text='HYPERLINK "https://'), ldm.FieldStart(), ldm.Run(text="DOCVARIABLE URL"),
        ldm.FieldSeparator(), ldm.Run(text="example.com"), ldm.FieldEnd(),
        ldm.Run(text='/path"'), ldm.FieldSeparator(), ldm.Run(text="LABEL"), ldm.FieldEnd()])]
    result = doc.to_bytes("md").decode()
    assert result.strip() == "[LABEL](https://example.com/path)"
