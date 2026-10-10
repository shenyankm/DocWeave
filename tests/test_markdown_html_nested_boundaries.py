"""Multiline HTML table blocks end at their outer table, not a nested close."""

import warnings
from io import BytesIO

import pytest

from aspose.words_foss import Document, MarkdownLoadOptions, SaveFormat
from aspose.words_foss import light_document_model as ldm


@pytest.mark.parametrize(
    "lookalike",
    [
        "",
        "<!-- </table> -->",
        '<script>"</table>"</script>',
        "\n",
        '<span title="</table>"></span>',
    ],
)
@pytest.mark.parametrize("save_format", [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
@pytest.mark.parametrize("close", ["</table>", "</TABLE>", "</table >"])
def test_nested_close_and_inert_lookalikes_do_not_truncate_later_rows(
    lookalike, save_format, close
):
    raw = (
        "<table>\n<tr><td>H0</td></tr>\n<tr><td>\n"
        + (lookalike + "\n" if lookalike else "")
        + "<p>BEFORE</p>\n<table>\n<tr><td>INNER0</td></tr>\n"
        "<tr><td>INNER1</td></tr>\n</table>\n<p>AFTER</p>\n</td></tr>\n"
        "<tr><td>LAST</td></tr>\n</table>\n\nTAIL"
    )
    raw = raw.replace("</table>", close)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        doc = Document(BytesIO(raw.encode()), MarkdownLoadOptions())
    cold = Document(BytesIO(doc.to_bytes(save_format)))
    for current in (doc, cold):
        body = current.light_document_model.sections[0].body
        table = next(node for node in body.children if isinstance(node, ldm.Table))
        assert len(table.rows) == 3
        assert table.rows[0].cells[0].paragraphs[0].text == "H0"
        assert table.rows[2].cells[0].paragraphs[0].text == "LAST"
        cell = table.rows[1].cells[0]
        assert [type(node).__name__ for node in cell.children] == [
            "Paragraph",
            "Table",
            "Paragraph",
        ]
        assert [p.text for p in cell.paragraphs] == ["BEFORE", "AFTER"]
        assert [r.cells[0].paragraphs[0].text for r in cell.tables[0].rows] == [
            "INNER0",
            "INNER1",
        ]
        assert any(paragraph.text == "TAIL" for paragraph in body.paragraphs)


def test_boundary_preserves_existing_commonmark_blank_line_termination():
    raw = "<div>\nINSIDE\n\nOUTSIDE\n\n</div>"
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        doc = Document(BytesIO(raw.encode()), MarkdownLoadOptions())
    body = doc.light_document_model.sections[0].body
    assert not any(isinstance(node, ldm.Table) for node in body.children)
    assert any(paragraph.text == "INSIDE" for paragraph in body.paragraphs)
    assert any(paragraph.text == "OUTSIDE" for paragraph in body.paragraphs)
