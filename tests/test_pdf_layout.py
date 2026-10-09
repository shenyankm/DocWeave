"""Layout regressions use an independent PDF renderer, not just text extraction."""

import pymupdf
import pytest
from fpdf import FPDF
from fpdf.enums import MethodReturnValue

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.pdf_writer.font import register_fonts


@pytest.mark.parametrize("alignment,highlight", [(0, True), (1, False), (2, False)])
def test_long_chinese_runs_wrap_inside_margins(tmp_path, alignment, highlight):
    text = "中文长句换行验证" * 35
    paragraph = ldm.Paragraph(
        paragraph_format=ldm.ParagraphFormat(alignment=alignment),
        children=[
            ldm.Run(
                text=text[:110],
                font=ldm.Font(
                    size=11,
                    bold=True,
                    highlight_color="Color [A=255, R=255, G=255, B=0]" if highlight else "",
                ),
            ),
            ldm.Run(
                text=text[110:],
                font=ldm.Font(
                    size=13,
                    italic=True,
                    highlight_color="Color [A=255, R=255, G=255, B=0]" if highlight else "",
                ),
            ),
        ],
    )
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[paragraph]))])
    docx = tmp_path / "long.docx"
    LdmDocxWriter().write(model, docx)
    output = tmp_path / "long.pdf"
    aw.Document(docx).save(output)
    with pymupdf.open(output) as pdf:
        actual = "".join(page.get_text() for page in pdf).replace("\n", "").replace(" ", "")
        assert actual == text
        lines = []
        for page in pdf:
            for block in page.get_text("dict")["blocks"]:
                for line in block.get("lines", []):
                    lines.append(line)
                    x0, y0, x1, y1 = line["bbox"]
                    assert x0 >= 19 * 72 / 25.4
                    assert x1 <= page.rect.width - 19 * 72 / 25.4
                    assert 0 <= y0 < y1 <= page.rect.height
            assert page.get_pixmap().samples
        assert len(lines) > 3
        if highlight:
            fills = [d for page in pdf for d in page.get_drawings() if d.get("fill") == (1, 1, 0)]
            assert len(fills) > 3
            assert all(d["rect"].x1 <= pdf[0].rect.width - 19 * 72 / 25.4 for d in fills)


def test_height_uses_font_metrics_and_explicit_line_breaks():
    text = "中文正文" * 30 + "\n第二行"
    paragraph = ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(size=11))])
    pdf = FPDF()
    pdf.c_margin = 0  # Measure the full paragraph width, without FPDF cell padding.
    register_fonts(pdf)
    pdf.add_page()
    pdf.set_font("DocumentSansSC", size=11)
    writer = LdmPdfWriter()
    line_h = writer._paragraph_renderer.line_height_mm(11, paragraph.paragraph_format)
    expected = pdf.multi_cell(
        w=40, h=line_h, text=text, dry_run=True, output=MethodReturnValue.HEIGHT
    )
    assert writer._estimate_paragraph_height(paragraph, 40) == pytest.approx(expected)


def test_chinese_columns_keep_content_and_header_footer(tmp_path):
    paragraphs = [
        ldm.Paragraph(children=[ldm.Run(text=f"段落{i}：" + "中文内容" * 20)]) for i in range(18)
    ]
    section = ldm.Section(
        page_setup=ldm.PageSetup(text_columns=ldm.TextColumns(count=2)),
        body=ldm.Body(children=paragraphs),
        headers_footers=[
            ldm.HeaderFooter(
                header_footer_type=0, children=[ldm.Paragraph(children=[ldm.Run(text="中文页眉")])]
            ),
            ldm.HeaderFooter(
                header_footer_type=1, children=[ldm.Paragraph(children=[ldm.Run(text="中文页脚")])]
            ),
        ],
    )
    output = tmp_path / "columns.pdf"
    LdmPdfWriter().write(ldm.Document(sections=[section]), output)
    with pymupdf.open(output) as pdf:
        text = "".join(page.get_text() for page in pdf).replace("\n", "")
        for i in range(18):
            assert f"段落{i}：" in text
        assert text.count("中文内容") == 360
        for page in pdf:
            assert "中文页眉" in page.get_text()
            assert "中文页脚" in page.get_text()
            for word in page.get_text("words"):
                assert 0 <= word[0] < word[2] <= page.rect.width


def test_wrapping_keeps_short_latin_words_and_underlines():
    pdf = FPDF()
    register_fonts(pdf)
    pdf.add_page()
    pdf.set_font("DocumentSansSC", style="U", size=11)
    run = ldm.Run(text="alpha bravo charlie delta", font=ldm.Font(size=11, underline=True))
    writer = LdmPdfWriter()
    rows = writer._run_renderer.wrap_segments(
        pdf, [(run, run.text, 0, 11, None)], pdf.get_string_width("charlie") + 1
    )
    lines = ["".join(segment[1] for segment in row) for row in rows]
    assert "".join(lines) == run.text
    assert " ".join(lines).split() == run.text.split()
    assert pdf.underline


def test_mixed_font_sizes_measure_largest_run():
    paragraph = ldm.Paragraph(
        children=[
            ldm.Run(text="小字", font=ldm.Font(size=8)),
            ldm.Run(text="大", font=ldm.Font(size=40)),
        ]
    )
    writer = LdmPdfWriter()
    expected = writer._paragraph_renderer.line_height_mm(40, paragraph.paragraph_format)
    assert writer._estimate_paragraph_height(paragraph, 180) == pytest.approx(expected)
