"""Cell-local empty paragraph and terminal-break rules from owned HTML cases."""

import warnings
from io import BytesIO

import pytest

from aspose.words_foss import Document, MarkdownLoadOptions, SaveFormat
from aspose.words_foss import light_document_model as ldm

NESTED = "<table><tr><td>NESTED</td></tr></table>"


@pytest.mark.parametrize(
    "content,expected",
    [
        ("", [""]),
        ("<p></p><p></p>", [""]),
        ("<p> </p>", [""]),
        ("<p>A</p><p></p><p></p>", ["A"]),
        ("<p></p><p>A</p>", ["A"]),
        ("<p>A</p><p></p><p>B</p>", ["A", "B"]),
        ("<p>A</p><p> </p>", ["A"]),
        ("<p>A</p><p><strong></strong></p>", ["A"]),
        ('<p>A</p><p style="font-size:30pt"></p>', ["A"]),
        ("<p>A</p><p>&nbsp;</p>", ["A", "\u00a0"]),
        ("<p>A<br></p>", ["A"]),
        ("<p><br></p>", [""]),
        ("<p>A</p><p><br></p>", ["A", ""]),
        ("<p>A</p><p><br></p><p>B</p>", ["A", "", "B"]),
        ("<p><br>A</p>", ["\nA"]),
        ("<p>A<br>B</p>", ["A\nB"]),
        ("<p>A<br><br></p>", ["A\n"]),
        ("<p><br><br></p>", ["\n"]),
        ("<p><br>A<br></p>", ["\nA"]),
        ("<p>A</p>TEXT<p></p>", ["A", "TEXT"]),
    ],
)
@pytest.mark.parametrize("save_format", [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_only_contentless_cell_paragraphs_omitted_and_one_final_break_consumed(
    content, expected, save_format
):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        doc = Document(
            BytesIO(("<table><tr><td>" + content + "</td></tr></table>").encode()),
            MarkdownLoadOptions(),
        )
    cold = Document(BytesIO(doc.to_bytes(save_format)))
    for current in (doc, cold):
        table = next(
            node
            for node in current.light_document_model.sections[0].body.children
            if isinstance(node, ldm.Table)
        )
        assert [p.text for p in table.rows[0].cells[0].paragraphs] == expected


@pytest.mark.parametrize(
    "content,expected",
    [
        ("<p>A</p>" + NESTED + "<p></p>", ["A", "TABLE", ""]),
        ("<p></p>" + NESTED, ["TABLE", ""]),
        (NESTED + "<p></p><p>B</p>", ["TABLE", "B"]),
        (NESTED + "<p><br></p>", ["TABLE", ""]),
    ],
)
@pytest.mark.parametrize("save_format", [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_nested_table_requires_trailing_paragraph_but_not_interior_empty(
    content, expected, save_format
):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        doc = Document(
            BytesIO(("<table><tr><td>" + content + "</td></tr></table>").encode()),
            MarkdownLoadOptions(),
        )
    cold = Document(BytesIO(doc.to_bytes(save_format)))
    for current in (doc, cold):
        table = next(
            node
            for node in current.light_document_model.sections[0].body.children
            if isinstance(node, ldm.Table)
        )
        cell = table.rows[0].cells[0]
        assert [
            "TABLE" if isinstance(node, ldm.Table) else node.text
            for node in cell.children
        ] == expected
        assert cell.tables[0].rows[0].cells[0].paragraphs[0].text == "NESTED"


@pytest.mark.parametrize("save_format", [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_vertical_restart_does_not_acquire_exported_empty_continuation_paragraph(
    save_format,
):
    raw = (
        '<table><tr><td rowspan="2"><p>V1</p><p></p></td><td>X</td></tr>'
        "<tr><td>Y</td></tr></table>"
    )
    doc = Document(BytesIO(raw.encode()), MarkdownLoadOptions())
    cold = Document(BytesIO(doc.to_bytes(save_format)))
    for current in (doc, cold):
        table = next(
            node
            for node in current.light_document_model.sections[0].body.children
            if isinstance(node, ldm.Table)
        )
        assert [p.text for p in table.rows[0].cells[0].paragraphs] == ["V1"]
        assert table.rows[0].cells[0].cell_format.vertical_merge == 1
        assert table.rows[1].cells[0].cell_format.vertical_merge == 2
        assert [p.text for p in table.rows[1].cells[0].paragraphs] == [""]
