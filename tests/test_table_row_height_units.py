"""Table minimum height uses the same millimetre units as measured content."""

import pymupdf
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.pdf_writer.constants import POST_TABLE_SPACING_MM, PT_TO_MM
from aspose.words_foss.saving import PdfSaveOptions


def render_table(size, *, padding=0, orientation=0, text="SMALL", height=0):
    table = ldm.Table(
        top_padding=padding,
        bottom_padding=padding,
        rows=[
            ldm.Row(
                row_format=ldm.RowFormat(
                    height=height,
                    borders=[
                        ldm.Border(line_style=1, line_width=0.5) for _ in range(6)
                    ],
                ),
                cells=[
                    ldm.Cell(
                        cell_format=ldm.CellFormat(orientation=orientation),
                        paragraphs=[
                            ldm.Paragraph(
                                children=[ldm.Run(text=text, font=ldm.Font(size=size))]
                            )
                        ],
                    )
                ],
            )
        ],
    )
    model = ldm.Document(
        sections=[
            ldm.Section(
                page_setup=ldm.PageSetup(
                    page_width=220,
                    page_height=180,
                    left_margin=20,
                    right_margin=20,
                    top_margin=20,
                    bottom_margin=20,
                ),
                body=ldm.Body(children=[table]),
            )
        ]
    )
    return model, table


@pytest.mark.parametrize("size", [6, 8, 12, 20])
@pytest.mark.parametrize("padding", [0, 2])
@pytest.mark.parametrize("shaping", [False, True])
def test_small_font_row_height_follows_content_and_shared_measurement(
    size, padding, shaping
):
    model, table = render_table(size, padding=padding)
    options = PdfSaveOptions()
    options.text_shaping = shaping
    writer = LdmPdfWriter(options)
    with pymupdf.open(stream=writer.write_to_bytes(model), filetype="pdf") as pdf:
        assert len(pdf) == 1
        assert pdf[0].get_text().strip() == "SMALL"
        borders = [d["rect"] for d in pdf[0].get_drawings()]
        row_height = max(r.y1 for r in borders) - min(r.y0 for r in borders)
    assert row_height == pytest.approx(size * 1.4 + padding * 2, abs=0.02)
    assert writer._estimate_table_height(table, 180 * PT_TO_MM) == pytest.approx(
        row_height * PT_TO_MM + POST_TABLE_SPACING_MM, abs=0.02
    )


@pytest.mark.parametrize("height", [4, 24])
def test_explicit_row_height_is_a_floor_above_content(height):
    model, _ = render_table(8, height=height)
    with pymupdf.open(
        stream=LdmPdfWriter().write_to_bytes(model), filetype="pdf"
    ) as pdf:
        borders = [d["rect"] for d in pdf[0].get_drawings()]
        assert max(r.y1 for r in borders) - min(r.y0 for r in borders) == pytest.approx(
            max(11.2, height), abs=0.02
        )


@pytest.mark.parametrize("orientation", [1, 2])
@pytest.mark.parametrize("text", ["I", "LONGER"])
def test_rotated_small_font_keeps_text_inside_row(orientation, text):
    model, table = render_table(8, orientation=orientation, text=text)
    writer = LdmPdfWriter()
    with pymupdf.open(stream=writer.write_to_bytes(model), filetype="pdf") as pdf:
        borders = [d["rect"] for d in pdf[0].get_drawings()]
        top, bottom = min(r.y0 for r in borders), max(r.y1 for r in borders)
        words = pdf[0].get_text("words")
        assert words[0][4] == text
        assert words[0][1] >= top - 0.1
        assert words[0][3] <= bottom + 0.1
        if text == "I":
            assert bottom - top == pytest.approx(5.5, abs=0.02)
    assert writer._estimate_table_height(table, 180 * PT_TO_MM) == pytest.approx(
        (bottom - top) * PT_TO_MM + POST_TABLE_SPACING_MM, abs=0.02
    )
