"""Self-contained inert raw-HTML Markdown table import regressions."""

import warnings
from io import BytesIO

import pytest

from aspose.words_foss import Document, MarkdownLoadOptions, SaveFormat, saving
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.md_import import html_tables
from aspose.words_foss.md_import.document_builder import MarkdownDocumentBuilder


def load(text):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return Document(BytesIO(text.encode()), MarkdownLoadOptions())


def tables(doc):
    return [
        node
        for node in doc.light_document_model.sections[0].body.children
        if isinstance(node, ldm.Table)
    ]


def shape(table):
    return [
        [
            (
                cell.cell_format.grid_span,
                cell.cell_format.vertical_merge,
                [paragraph.text for paragraph in cell.paragraphs],
            )
            for cell in row.cells
        ]
        for row in table.rows
    ]


@pytest.mark.parametrize("save_format", [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
@pytest.mark.parametrize("rtl", [False, True])
def test_headers_omitted_slots_and_spans_save_cold(save_format, rtl):
    source = (
        '<p>BEFORE</p><table dir="' + ("rtl" if rtl else "ltr") + '">'
        "<thead><tr><th>H0</th><th>H1</th><th>H2</th></tr></thead><tbody>"
        '<tr><td></td><td rowspan="2">V</td><td>X</td></tr>'
        '<tr><td>C</td><td>Y</td></tr><tr><td colspan="2">W</td><td>Z</td></tr>'
        "</tbody></table><p>AFTER</p>"
    )
    doc = load(source)
    expected = [
        [(1, 0, ["H0"]), (1, 0, ["H1"]), (1, 0, ["H2"])],
        [(1, 0, [""]), (1, 1, ["V"]), (1, 0, ["X"])],
        [(1, 0, ["C"]), (1, 2, [""]), (1, 0, ["Y"])],
        [(2, 0, ["W"]), (1, 0, ["Z"])],
    ]
    assert shape(tables(doc)[0]) == expected
    assert tables(doc)[0].bidi is rtl
    before = doc.light_document_model.model_dump()
    cold = Document(BytesIO(doc.to_bytes(save_format)))
    assert shape(tables(cold)[0]) == expected
    assert tables(cold)[0].bidi is rtl
    assert doc.light_document_model.model_dump() == before
    assert [
        p.text for p in cold.light_document_model.sections[0].body.paragraphs if p.text
    ] == ["BEFORE", "AFTER"]


@pytest.mark.parametrize("save_format", [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_nested_tables_keep_order_text_and_run_format(save_format):
    doc = load(
        "<table><tr><td><p>A<strong>B</strong><em>C</em><br>D &amp; E</p>"
        "<table><tr><td>NESTED</td></tr></table><p>AFTER</p></td></tr></table>"
    )
    for current in (doc, Document(BytesIO(doc.to_bytes(save_format)))):
        cell = tables(current)[0].rows[0].cells[0]
        assert [type(node).__name__ for node in cell.children] == [
            "Paragraph",
            "Table",
            "Paragraph",
        ]
        assert cell.paragraphs[0].text == "ABC\nD & E"
        assert cell.paragraphs[1].text == "AFTER"
        assert shape(cell.tables[0]) == [[(1, 0, ["NESTED"])]]
        runs = cell.paragraphs[0]._children
        assert next(run for run in runs if run.text == "B").font.bold
        assert next(run for run in runs if run.text == "C").font.italic


@pytest.mark.parametrize(
    "source,expected",
    [
        ('<tr><td colspan="2">A</td><td>B</td></tr>', [[(1, 0, ["A"]), (1, 0, ["B"])]]),
        (
            '<tr><td colspan="3">A</td><td>B</td></tr><tr><td colspan="2">C</td><td>D</td><td>E</td></tr>',
            [
                [(2, 0, ["A"]), (1, 0, ["B"])],
                [(1, 0, ["C"]), (1, 0, ["D"]), (1, 0, ["E"])],
            ],
        ),
        (
            '<tr><td rowspan="2" colspan="2">A</td><td>B</td></tr><tr><td>C</td></tr>',
            [[(1, 1, ["A"]), (1, 0, ["B"])], [(1, 2, [""]), (1, 0, ["C"])]],
        ),
        ("<tr><td>A<td>B", [[(1, 0, ["A"]), (1, 0, ["B"])]]),
        (
            '<tr><td rowspan="0">A</td><td>B</td></tr><tr><td>C</td></tr>',
            [[(1, 0, ["A"]), (1, 0, ["B"])], [(1, 0, ["C"]), (1, 0, [""])]],
        ),
        ('<tr><td colspan="-2" rowspan="bad">A</td></tr>', [[(1, 0, ["A"])]]),
    ],
)
def test_native_column_boundary_projection_and_repair(source, expected):
    doc = load("<table>" + source + "</table>")
    assert shape(tables(doc)[0]) == expected
    assert (
        shape(tables(Document(BytesIO(doc.to_bytes(SaveFormat.DOCX))))[0]) == expected
    )


def test_html_table_in_existing_builder_cell_uses_ordered_typed_lists():
    builder = MarkdownDocumentBuilder()
    cell = builder.insert_cell()
    builder.write("BEFORE")
    builder.insert_html("<table><tr><td>NESTED</td></tr></table><p>AFTER</p>")
    builder.write("TAIL")
    assert all(isinstance(node, ldm.Paragraph) for node in cell.paragraphs)
    assert [type(node).__name__ for node in cell.children] == [
        "Paragraph",
        "Table",
        "Paragraph",
        "Paragraph",
    ]
    assert [p.text for p in cell.paragraphs] == ["BEFORE", "AFTER", "TAIL"]
    assert shape(cell.tables[0]) == [[(1, 0, ["NESTED"])]]


def test_active_content_resources_css_diagnosed_without_secret_leak():
    doc = load(
        '<table style="width:3pt"><tr><td onclick="private-secret">OK'
        '<script>private-secret()</script><iframe src="file:///private-secret"></iframe>'
        '<img src="https://private-secret" alt="ALT"><a href="https://private-secret">LINK</a>'
        "</td></tr></table>"
    )
    assert shape(tables(doc)[0]) == [[(1, 0, ["OKALTLINK"])]]
    codes = {d.code for d in doc.diagnostics}
    assert {
        "markdown.html_active_content",
        "markdown.html_resources",
        "markdown.html_css",
    } <= codes
    assert all("private-secret" not in d.message for d in doc.diagnostics)


@pytest.mark.parametrize(
    "html,match",
    [
        ("<!DOCTYPE html><table></table>", "declarations"),
        ('<table><tr><td colspan="1025">A</td></tr></table>', "safety limit"),
        ('<table><tr><td rowspan="100001">A</td></tr></table>', "safety limit"),
        ("<table>" + "<div>" * 65 + "</div>" * 65 + "</table>", "nesting"),
        (
            '<table><tr><td>A</td><td rowspan="2">B</td></tr><tr><td colspan="2">C</td></tr></table>',
            "overlap",
        ),
    ],
)
def test_adversarial_input_rejected_before_builder_mutation(html, match):
    builder = MarkdownDocumentBuilder()
    builder.write("UNCHANGED")
    before = builder.document.model_dump()
    with warnings.catch_warnings(), pytest.raises(ValueError, match=match):
        warnings.simplefilter("ignore")
        builder.insert_html(html)
    assert builder.document.model_dump() == before


def test_node_and_expanded_cell_budgets(monkeypatch):
    monkeypatch.setattr(html_tables, "MAX_HTML_NODES", 8)
    with pytest.raises(ValueError, match="node safety"):
        html_tables.parse_html_tables(
            "<table><tr>" + "<td>A</td>" * 4 + "</tr></table>", ldm.Font()
        )
    monkeypatch.setattr(html_tables, "MAX_HTML_NODES", 100)
    parser = html_tables.TableHTMLParser()
    parser.feed("<table><tr><td>A</td><td>B</td></tr><tr><td>C</td></tr></table>")
    monkeypatch.setattr(html_tables, "MAX_HTML_NODES", 3)
    with pytest.raises(ValueError, match="cell safety"):
        parser.blocks(parser.root.children, ldm.Font())


@pytest.mark.parametrize("save_format", [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_header_semantics_and_exported_html_reload(save_format):
    doc = load(
        "<table><thead><tr><th>A &lt; B</th><th>H1</th></tr></thead>"
        "<tbody><tr><td></td><td>X</td></tr></tbody></table>"
    )
    saved = Document(BytesIO(doc.to_bytes(save_format)))
    assert tables(saved)[0].rows[0].row_format.heading_format
    assert not tables(saved)[0].rows[1].row_format.heading_format
    options = saving.MarkdownSaveOptions()
    options.export_as_html = saving.MarkdownExportAsHtml.TABLES
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        cold_html = load(doc.to_bytes(options).decode())
    cold = Document(BytesIO(cold_html.to_bytes(save_format)))
    assert shape(tables(cold)[0]) == shape(tables(doc)[0])
    assert tables(cold)[0].rows[0].row_format.heading_format
    header = tables(cold)[0].rows[0].cells[0].paragraphs[0]
    assert header.paragraph_format.alignment == 1
    assert header._children[0].font.bold
    assert not tables(cold)[0].rows[1].row_format.heading_format


def test_entities_preserve_nonbreaking_space_and_literal_markup():
    doc = load("<table><tr><td>A&nbsp;B &lt;script&gt; &amp; C</td></tr></table>")
    assert shape(tables(doc)[0]) == [[(1, 0, ["A\u00a0B <script> & C"])]]
