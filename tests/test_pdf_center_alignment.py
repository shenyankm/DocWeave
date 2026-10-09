"""Centered text must be centered in the effective paragraph area."""

import pymupdf
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.saving import PdfSaveOptions


def center_model(kind="plain", columns=1, line_count=1):
    pieces = []
    for index in range(line_count):
        pieces.extend(
            [
                ldm.Run(text=("\n" if index else "") + "CENT", font=ldm.Font(size=12)),
                ldm.Run(text="ER", font=ldm.Font(size=16, bold=True)),
            ]
        )
    pf = ldm.ParagraphFormat(alignment=1, left_indent=24, right_indent=6)
    para = ldm.Paragraph(children=pieces, paragraph_format=pf)
    if kind == "list":
        pf.first_line_indent = -18
        para.list_format = ldm.ListFormat(is_list_item=True, list_id=1)
        para.list_label = ldm.ListLabel(label_string="1.")
    return ldm.Document(
        sections=[
            ldm.Section(
                page_setup=ldm.PageSetup(
                    page_width=300,
                    page_height=180,
                    left_margin=20,
                    right_margin=20,
                    top_margin=20,
                    bottom_margin=20,
                    text_columns=ldm.TextColumns(count=columns, spacing=15),
                ),
                body=ldm.Body(children=[para]),
            )
        ]
    )


@pytest.mark.parametrize("kind", ["plain", "list"])
@pytest.mark.parametrize("shaping", [False, True])
@pytest.mark.parametrize("columns", [1, 2])
@pytest.mark.parametrize("line_count", [1, 3])
def test_centered_mixed_runs_use_effective_area_center(
    kind, shaping, columns, line_count
):
    doc = center_model(kind, columns, line_count)
    snapshot = doc.model_dump()
    options = PdfSaveOptions()
    options.text_shaping = shaping
    raw = LdmPdfWriter(options).write_to_bytes(doc)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        lines = [
            [c for span in line["spans"] for c in span["chars"] if c["c"] in "CENTER"]
            for page in pdf
            for b in page.get_text("rawdict")["blocks"]
            for line in b.get("lines", [])
        ]
        lines = [chars for chars in lines if chars]
        assert len(lines) == line_count
        col_width = (260 - 15 * (columns - 1)) / columns
        expected = (20 + 24 + 20 + col_width - 6) / 2
        for chars in lines:
            assert "".join(c["c"] for c in chars) == "CENTER"
            x0 = min(c["bbox"][0] for c in chars)
            x1 = max(c["bbox"][2] for c in chars)
            assert (x0 + x1) / 2 == pytest.approx(expected, abs=0.05)
    assert doc.model_dump() == snapshot
