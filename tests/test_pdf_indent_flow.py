"""Paragraph insets must survive pages/columns without leaking into page bands."""
import pymupdf
import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.pdf_writer.constants import PT_TO_MM


@pytest.mark.parametrize("alignment", [0, 1, 2])
@pytest.mark.parametrize("first_indent", [-18, 18])
def test_measured_height_matches_mixed_indented_runs(tmp_path, monkeypatch, alignment, first_indent):
    paragraph = ldm.Paragraph(
        paragraph_format=ldm.ParagraphFormat(alignment=alignment, left_indent=36,
            right_indent=18, first_line_indent=first_indent, space_before=3, space_after=5),
        children=[ldm.Run(text="中文混合" * 12, font=ldm.Font(size=10, bold=True)),
                  ldm.Run(text="字号高亮" * 12, font=ldm.Font(size=16, italic=True,
                    highlight_color="Color [A=255, R=255, G=255, B=0]"))])
    doc = ldm.Document(sections=[ldm.Section(page_setup=ldm.PageSetup(
        page_width=360, page_height=1200, left_margin=36, right_margin=36),
        body=ldm.Body(children=[paragraph]))])
    writer = LdmPdfWriter()
    original = writer._paragraph_renderer.render_paragraph
    measurements = []
    def record(pdf, para):
        expected = writer._estimate_paragraph_height(para, pdf.epw)
        y, page = pdf.y, pdf.page
        original(pdf, para)
        assert pdf.page == page
        measurements.append((expected, pdf.y - y))
    monkeypatch.setattr(writer._paragraph_renderer, "render_paragraph", record)
    writer.write(doc, tmp_path / "mixed.pdf")
    assert len(measurements) == 1
    assert measurements[0][0] == pytest.approx(measurements[0][1], abs=0.05)


@pytest.mark.parametrize("columns", [1, 2])
@pytest.mark.parametrize("highlight", [False, True])
def test_insets_and_page_bands_survive_page_and_column_breaks(tmp_path, columns, highlight):
    text = "中文正文测试" * 240
    paragraph = ldm.Paragraph(paragraph_format=ldm.ParagraphFormat(
        left_indent=36, right_indent=18, first_line_indent=-18),
        children=[ldm.Run(text=text, font=ldm.Font(size=12,
            highlight_color="Color [A=255, R=255, G=255, B=0]" if highlight else ""))])
    section = ldm.Section(page_setup=ldm.PageSetup(page_width=500, page_height=500,
        left_margin=36, right_margin=36, top_margin=50, bottom_margin=50,
        text_columns=ldm.TextColumns(count=columns, spacing=18)),
        body=ldm.Body(children=[paragraph]), headers_footers=[
            ldm.HeaderFooter(header_footer_type=0,
                children=[ldm.Paragraph(children=[ldm.Run(text="HEADER")])]),
            ldm.HeaderFooter(header_footer_type=1,
                children=[ldm.Paragraph(children=[ldm.Run(text="FOOTER")])]),
        ])
    source = tmp_path / "flow.docx"
    LdmDocxWriter().write(ldm.Document(sections=[section]), source)
    output = tmp_path / "flow.pdf"
    aw.Document(source).save(output)
    with pymupdf.open(output) as pdf:
        assert len(pdf) > 1
        extracted = "".join(page.get_text() for page in pdf).replace("\n", "")
        assert extracted.replace("HEADER", "").replace("FOOTER", "") == text
        width = (428 - (columns - 1) * 18) / columns
        seen = set()
        for page_index, page in enumerate(pdf):
            column_rows = [0] * columns
            for block in page.get_text("dict")["blocks"]:
                for line in block.get("lines", []):
                    for span in line["spans"]:
                        x0, y0, x1, y1 = span["bbox"]
                        if span["text"] in ("HEADER", "FOOTER"):
                            assert x0 < 45, "Body indentation leaked into a page band"
                            continue
                        column = 0 if columns == 1 or x0 < 250 else 1
                        seen.add(column)
                        column_rows[column] += 1
                        base = 36 + column * (width + 18)
                        assert x0 >= base + 18 - 0.1
                        assert x1 <= base + width - 18 + 0.1
                        assert 0 <= y0 < y1 <= page.rect.height
            if page_index < len(pdf) - 1:
                assert all(count > 4 for count in column_rows), "New page reused an exhausted column cursor"
            assert page.get_pixmap().samples
        assert seen == set(range(columns))


def test_keep_together_advances_to_next_column_not_next_page(tmp_path):
    prefix = ldm.Paragraph(paragraph_format=ldm.ParagraphFormat(space_after=65),
        children=[ldm.Run(text="前置占位", font=ldm.Font(size=14))])
    text = "中文测量" * 19
    target = ldm.Paragraph(paragraph_format=ldm.ParagraphFormat(
        first_line_indent=72, keep_together=True),
        children=[ldm.Run(text=text, font=ldm.Font(size=14))])
    document = ldm.Document(sections=[ldm.Section(page_setup=ldm.PageSetup(
        page_width=348, page_height=310, left_margin=15, right_margin=15,
        top_margin=30, bottom_margin=30, text_columns=ldm.TextColumns(count=2, spacing=18)),
        body=ldm.Body(children=[prefix, target]))])
    output = tmp_path / "keep-columns.pdf"
    LdmPdfWriter().write(document, output)
    with pymupdf.open(output) as pdf:
        assert len(pdf) == 1
        assert text in pdf[0].get_text().replace("\n", "")
        assert any(word[0] > 174 for word in pdf[0].get_text("words"))


@pytest.mark.parametrize("extra_space_mm", [0.0, 0.01])
def test_keep_together_exact_fit_and_just_over(tmp_path, extra_space_mm):
    text = "中文测量" * 19
    target = ldm.Paragraph(paragraph_format=ldm.ParagraphFormat(
        first_line_indent=72, keep_together=True),
        children=[ldm.Run(text=text, font=ldm.Font(size=14))])
    prefix = ldm.Paragraph(children=[ldm.Run(text="前置占位", font=ldm.Font(size=14))])
    setup = ldm.PageSetup(page_width=180, page_height=100 / PT_TO_MM,
        left_margin=15, right_margin=15, top_margin=30, bottom_margin=30)
    document = ldm.Document(sections=[ldm.Section(page_setup=setup,
        body=ldm.Body(children=[target]))])
    writer = LdmPdfWriter()
    writer.write(document, tmp_path / "control.pdf")
    width = 150 * PT_TO_MM
    remaining = 100 - 60 * PT_TO_MM
    before = remaining - writer._estimate_paragraph_height(target, width)
    before -= writer._estimate_paragraph_height(prefix, width)
    assert before > 0
    prefix.paragraph_format.space_after = (before + extra_space_mm) / PT_TO_MM
    document.sections[0].body.children.insert(0, prefix)
    output = tmp_path / "fit.pdf"
    writer.write(document, output)
    with pymupdf.open(output) as pdf:
        assert len(pdf) == (1 if extra_space_mm == 0 else 2)
        assert text in pdf[-1].get_text().replace("\n", "")


def test_too_tall_keep_together_splits_without_an_empty_first_page(tmp_path):
    text = "中文测量" * 90
    document = ldm.Document(sections=[ldm.Section(page_setup=ldm.PageSetup(
        page_width=180, page_height=280, left_margin=15, right_margin=15,
        top_margin=30, bottom_margin=30), body=ldm.Body(children=[ldm.Paragraph(
            paragraph_format=ldm.ParagraphFormat(first_line_indent=72, keep_together=True),
            children=[ldm.Run(text=text, font=ldm.Font(size=14))])]))])
    output = tmp_path / "too-tall.pdf"
    LdmPdfWriter().write(document, output)
    with pymupdf.open(output) as pdf:
        assert len(pdf) > 1
        assert all(page.get_text().strip() for page in pdf)
        assert "".join(page.get_text() for page in pdf).replace("\n", "") == text
