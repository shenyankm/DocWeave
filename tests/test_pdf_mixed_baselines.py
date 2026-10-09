"""Mixed glyph sizes share a baseline while native layout keeps cursor ownership."""

import pymupdf
import pytest
from fpdf import FPDF
from fpdf.enums import CharVPos

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.pdf_writer.baseline import baseline_scope
from aspose.words_foss.saving import PdfSaveOptions


def baseline_model(align=0, location="body", long=False, tabs=False):
    runs = [ldm.Run(text="QSMALL\t" if tabs else "QSMALL ", font=ldm.Font(size=12)),
            ldm.Run(text="ZBIG ", font=ldm.Font(size=18, bold=True)),
            ldm.Run(text="KMID", font=ldm.Font(size=14, italic=True))]
    if long:
        runs = runs * 35
    paragraph = ldm.Paragraph(children=runs, paragraph_format=ldm.ParagraphFormat(alignment=align))
    section = ldm.Section(page_setup=ldm.PageSetup(
        page_width=240, page_height=150, left_margin=15, right_margin=15, top_margin=15, bottom_margin=15),
        body=ldm.Body(children=[paragraph]))
    if location in ("header", "footer"):
        section.body.children = [ldm.Paragraph(children=[ldm.Run(text="BODY")])]
        section.headers_footers = [ldm.HeaderFooter(
            header_footer_type=0 if location == "header" else 1, children=[paragraph])]
    elif location == "table":
        section.body.children = [ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(children=[paragraph])])])]
    return ldm.Document(sections=[section])


def glyphs(page):
    return [(c, span["size"]) for b in page.get_text("rawdict")["blocks"]
            for line in b.get("lines", []) for span in line["spans"] for c in span["chars"]]


@pytest.mark.parametrize("align", [0, 1, 2])
@pytest.mark.parametrize("location", ["body", "table", "header", "footer"])
@pytest.mark.parametrize("shaping", [False, True])
def test_mixed_fonts_share_origin_without_changing_glyph_sizes(align, location, shaping):
    model = baseline_model(align, location)
    snapshot = model.model_dump()
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf") as pdf:
        chars = [(c, s) for c, s in glyphs(pdf[0]) if c["c"] in "QZK"]
        assert len(chars) == 3
        origins = [c["origin"][1] for c, _ in chars]
        assert max(origins) - min(origins) <= 0.03
        assert [s for _, s in chars] == pytest.approx([12, 18, 14], abs=0.05)
    assert model.model_dump() == snapshot


@pytest.mark.parametrize("shaping", [False, True])
def test_toc_title_and_numeric_runs_share_last_line_baseline(shaping):
    model = baseline_model()
    paragraph = model.sections[0].body.children[0]
    paragraph.paragraph_format.style_name = "TOC 1"
    paragraph.paragraph_format.tab_stops.add(150, alignment=2, leader=1)
    paragraph._children = [ldm.Run(text="TITLE", font=ldm.Font(size=12)),
                           ldm.Run(text="\t12", font=ldm.Font(size=16, bold=True)),
                           ldm.Run(text=".", font=ldm.Font(size=12)),
                           ldm.Run(text="30", font=ldm.Font(size=18, italic=True))]
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf") as pdf:
        origins = [c["origin"][1] for c, _ in glyphs(pdf[0]) if c["c"].strip()]
        assert max(origins) - min(origins) <= 0.03
        assert "".join(c["c"] for c, _ in glyphs(pdf[0])) == "TITLE12.30"
        assert len(pdf[0].get_drawings()) == 1


def test_plain_default_tab_preserves_mixed_run_baseline_and_grid():
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(baseline_model(tabs=True)), filetype="pdf") as pdf:
        chars = [(c, s) for c, s in glyphs(pdf[0]) if c["c"] in "QZK"]
        assert max(c["origin"][1] for c, _ in chars) - min(c["origin"][1] for c, _ in chars) <= 0.03
        assert next(c["bbox"][0] for c, _ in chars if c["c"] == "Z") == pytest.approx(87, abs=0.05)


@pytest.mark.parametrize("vpos", [CharVPos.LINE, CharVPos.SUP, CharVPos.SUB])
@pytest.mark.parametrize("spacing", [0, 1])
def test_native_vertical_position_and_spacing_survive_baseline_scope(vpos, spacing):
    pdf = FPDF(unit="pt", format=(200, 80))
    pdf.set_margins(10, 10, 10)
    pdf.set_auto_page_break(True, 10)
    pdf.add_page()
    pdf.set_font("Helvetica", size=12)
    pdf.char_vpos = vpos
    pdf.set_char_spacing(spacing)
    with baseline_scope(pdf, 18):
        for size, text in [(12, "A"), (18, "B"), (14, "C")]:
            pdf.set_font("Helvetica", size=size)
            pdf.cell(w=30, h=25, text=text)
    assert "_text_baseline_size" not in pdf.__dict__
    assert "_render_styled_text_line" not in pdf.__dict__
    assert pdf.char_vpos == vpos and pdf.char_spacing == spacing
    lift = pdf.sup_lift if vpos == CharVPos.SUP else pdf.sub_lift if vpos == CharVPos.SUB else 0
    scale = pdf.sup_scale if vpos == CharVPos.SUP else pdf.sub_scale if vpos == CharVPos.SUB else 1
    with pymupdf.open(stream=bytes(pdf.output()), filetype="pdf") as rendered:
        chars = glyphs(rendered[0])
        corrected = [c["origin"][1] + size * lift for (c, _), size in zip(chars, [12, 18, 14])]
        assert max(corrected) - min(corrected) <= 0.03
        assert [s for _, s in chars] == pytest.approx([s * scale for s in [12, 18, 14]], abs=0.05)


def test_scope_restores_overrides_after_render_failure(monkeypatch):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=12)

    def fail(*args, **kwargs):
        raise RuntimeError("paint failed")

    monkeypatch.setattr(pdf, "_render_styled_text_line", fail)
    with pytest.raises(RuntimeError, match="paint failed"), baseline_scope(pdf, 18):
        pdf.cell(w=30, h=10, text="A")
    assert pdf._render_styled_text_line is fail
    assert "_text_baseline_size" not in pdf.__dict__


@pytest.mark.parametrize("location", ["body", "table"])
@pytest.mark.parametrize("shaping", [False, True])
def test_list_marker_and_text_share_mixed_baseline(location, shaping):
    model = baseline_model(location=location)
    paragraph = (model.sections[0].body.children[0].rows[0].cells[0].paragraphs[0]
                 if location == "table" else model.sections[0].body.children[0])
    paragraph.list_format = ldm.ListFormat(is_list_item=True, list_id=1, list_level_number=0)
    paragraph.list_label = ldm.ListLabel(label_string="1.")
    options = PdfSaveOptions()
    options.text_shaping = shaping
    options.export_document_structure = True
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf") as pdf:
        origins = [c["origin"][1] for c, _ in glyphs(pdf[0]) if c["c"] in "1QZK"]
        assert len(origins) == 4 and max(origins) - min(origins) <= 0.03


@pytest.mark.parametrize("location", [0, 1])
@pytest.mark.parametrize("columns", [1, 2])
@pytest.mark.parametrize("shaping", [False, True])
def test_page_band_heading_does_not_inherit_body_baseline(location, columns, shaping):
    model = baseline_model(long=True)
    section = model.sections[0]
    section.page_setup.text_columns = ldm.TextColumns(count=columns, spacing=12)
    section.headers_footers = [ldm.HeaderFooter(header_footer_type=location, children=[
        ldm.Paragraph(children=[ldm.Run(text="HEADER", font=ldm.Font(size=10))],
                      paragraph_format=ldm.ParagraphFormat(is_heading=True, outline_level=1))])]
    options = PdfSaveOptions()
    options.text_shaping = shaping
    options.export_document_structure = True
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf") as pdf:
        assert len(pdf) > 1
        body = [(c, s) for page in pdf for c, s in glyphs(page) if c["c"] in "QZK"]
        assert sum(c["c"] == "Q" for c, _ in body) == 35
        assert sum(c["c"] == "Z" for c, _ in body) == 35
        assert sum(c["c"] == "K" for c, _ in body) == 35
        headers = [next(w for w in page.get_text("words") if w[4] == "HEADER") for page in pdf]
        assert all(h[1:4] == pytest.approx(headers[0][1:4], abs=0.03) for h in headers)
        assert all(s == pytest.approx({"Q": 12, "Z": 18, "K": 14}[c["c"]], abs=0.05) for c, s in body)
