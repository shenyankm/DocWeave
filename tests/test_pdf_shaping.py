"""Optional shaping uses real HarfBuzz glyphs, not merely Unicode font coverage."""

from io import BytesIO
from pathlib import Path
import warnings

from fontTools.ttLib import TTFont
from fpdf import FPDF
import pymupdf
from pypdf import PdfReader
import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import PdfMissingGlyphWarning
from aspose.words_foss.pdf_writer.writer import LdmPdfWriter
from aspose.words_foss.pdf_writer.constants import PT_TO_MM
from aspose.words_foss.pdf_writer.font import close_fonts, register_fonts


def arabic_font():
    candidates = [Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
                  Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
                  Path("C:/Windows/Fonts/arial.ttf")]
    for path in candidates:
        if path.is_file() and all(ord(char) in TTFont(path).getBestCmap() for char in "سلام"):
            return path
    pytest.skip("No suitable trusted Arabic font installed")


def test_shaped_arabic_embeds_real_contextual_glyph_outlines():
    hb = pytest.importorskip("uharfbuzz")
    font_path = arabic_font()
    source = "سلام"
    buffer = hb.Buffer()
    buffer.add_str(source)
    buffer.guess_segment_properties()
    hb.shape(hb.Font(hb.Face(font_path.read_bytes())), buffer)
    original = TTFont(font_path)
    expected = [original["glyf"][original.getGlyphOrder()[info.codepoint]].getCoordinates(original["glyf"])[0]
                for info in buffer.glyph_infos]
    options = aw.saving.PdfSaveOptions()
    options.text_shaping = True
    options.fallback_fonts = [str(font_path)]
    doc = aw.Document(BytesIO(source.encode()))
    with warnings.catch_warnings():
        warnings.simplefilter("error", PdfMissingGlyphWarning)
        raw = doc.to_bytes(options)
    reader = PdfReader(BytesIO(raw))
    assert source in reader.pages[0].extract_text()
    outlines = []
    for ref in reader.pages[0]["/Resources"]["/Font"].values():
        font = ref.get_object()["/DescendantFonts"][0].get_object()["/FontDescriptor"]
        embedded = TTFont(BytesIO(font["/FontFile2"].get_object().get_data()))
        outlines.extend(embedded["glyf"][name].getCoordinates(embedded["glyf"])[0]
                        for name in embedded.getGlyphOrder())
    assert all(any(actual == wanted for actual in outlines) for wanted in expected)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        assert pdf[0].get_pixmap().samples
        assert pdf[0].get_text("dict")["blocks"]


@pytest.mark.parametrize("alignment,highlight", [(0, False), (1, True), (2, False)])
def test_shaping_measurement_matches_real_wrapped_mixed_text(alignment, highlight):
    pytest.importorskip("uharfbuzz")
    options = aw.saving.PdfSaveOptions()
    options.text_shaping = True
    options.fallback_fonts = [str(arabic_font())]
    para = ldm.Paragraph(children=[ldm.Run(text="中文 hello سلام " * 18,
        font=ldm.Font(size=14, highlight_color="Color [A=255, R=255, G=255, B=0]" if highlight else ""))])
    para.paragraph_format.alignment = alignment
    model = ldm.Document(sections=[ldm.Section(page_setup=ldm.PageSetup(
        page_width=60 / PT_TO_MM + 72, page_height=1000,
        left_margin=36, right_margin=36, top_margin=36, bottom_margin=36),
        body=ldm.Body(children=[para]))])
    writer = LdmPdfWriter(options)
    raw = writer.write_to_bytes(model)
    estimate = writer._estimate_paragraph_height(para, 60)
    measure = writer._measurement_pdf
    shadow = writer._measurement_writer
    measure.set_xy(0, 0)
    measure.set_margins(0, 0, measure.w - 60)
    shadow._paragraph_renderer.render_paragraph(measure, para)
    assert measure.get_y() == pytest.approx(estimate, abs=0.001)
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        assert len(pdf) == 1
        spans = [span for block in pdf[0].get_text('dict')['blocks']
                 for line in block.get('lines', []) for span in line['spans']]
        baselines = {round(span['origin'][1], 1) for span in spans}
        line_height = writer._paragraph_renderer.line_height_mm(14, para.paragraph_format)
        assert len(baselines) * line_height == pytest.approx(estimate, abs=0.01)
        assert all(35 <= span['bbox'][0] < span['bbox'][2] <= pdf[0].rect.width - 35 for span in spans)
        assert ''.join(page.get_text() for page in pdf).count('hello') == 18
        if highlight:
            assert any(drawing['fill'] == (1, 1, 0) for drawing in pdf[0].get_drawings())


@pytest.mark.parametrize('alignment', [0, 1, 2])
@pytest.mark.parametrize('parts', [('س', 'لام'), ('A', 'V', 'A', 'T', 'A', 'R')])
def test_identical_style_run_boundaries_preserve_shaped_pixels(parts, alignment):
    pytest.importorskip('uharfbuzz')
    options = aw.saving.PdfSaveOptions()
    options.text_shaping = True
    options.fallback_fonts = [str(arabic_font())]
    pixels = []
    for texts in [(''.join(parts),), parts]:
        para = ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(size=28)) for text in texts],
                             paragraph_format=ldm.ParagraphFormat(alignment=alignment))
        model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[para]))])
        snapshot = model.model_dump()
        raw = LdmPdfWriter(options).write_to_bytes(model)
        assert model.model_dump() == snapshot
        with pymupdf.open(stream=raw, filetype='pdf') as pdf:
            pixels.append(pdf[0].get_pixmap(matrix=pymupdf.Matrix(2, 2)).samples)
    assert pixels[0] == pixels[1]


@pytest.mark.parametrize('shaping', [False, True])
@pytest.mark.parametrize('in_table', [False, True])
@pytest.mark.parametrize('split', [False, True])
def test_latin_after_fallback_font_is_visible_and_extractable(shaping, in_table, split):
    if shaping:
        pytest.importorskip('uharfbuzz')
    options = aw.saving.PdfSaveOptions()
    options.text_shaping = shaping
    options.fallback_fonts = [str(arabic_font())]
    parts = ('سلام ', 'ABC ', 'عالم ', 'DEF ', '123 ', 'نهاية')
    para = ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(size=20))
        for text in (parts if split else (''.join(parts),))])
    child = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[para])])]) if in_table else para
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[child]))])
    raw = LdmPdfWriter(options).write_to_bytes(model)
    parsed = PdfReader(BytesIO(raw))
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        text = ''.join(page.get_text() for page in pdf)
        independent = ''.join(page.extract_text() for page in parsed.pages)
        assert all(text.count(token) == independent.count(token) == 1 for token in ('ABC', 'DEF', '123'))
        assert pdf[0].get_pixmap().samples


def test_shaping_preserves_distinct_styles_and_link_targets():
    pytest.importorskip('uharfbuzz')
    options = aw.saving.PdfSaveOptions()
    options.text_shaping = True
    para = ldm.Paragraph(children=[
        ldm.Run(text='[LEFT](https://left.example)', is_hyperlink=True,
                font=ldm.Font(size=14, bold=True, strike_through=True, highlight_color='Color [A=255, R=255, G=255, B=0]')),
        ldm.Run(text='[RIGHT](https://right.example)', is_hyperlink=True, font=ldm.Font(size=18, italic=True)),
    ])
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[para]))])
    raw = LdmPdfWriter(options).write_to_bytes(model)
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        assert [link['uri'] for link in pdf[0].get_links()] == ['https://left.example', 'https://right.example']
        spans = [span for block in pdf[0].get_text('dict')['blocks']
                 for line in block.get('lines', []) for span in line['spans']]
        left = next(span for span in spans if 'LEFT' in span['text'])
        right = next(span for span in spans if 'RIGHT' in span['text'])
        assert left['size'] == pytest.approx(14) and left['flags'] & 16
        assert right['size'] == pytest.approx(18) and right['flags'] & 2
        drawings = [item for drawing in pdf[0].get_drawings() for item in drawing['items']]
        assert any(item[0] == 're' for item in drawings)
        assert any(item[0] == 'l' for item in drawings)


def test_shaped_rtl_punctuation_matches_native_glyph_positions():
    pytest.importorskip('uharfbuzz')
    text = 'سلام (123) عالم'
    options = aw.saving.PdfSaveOptions()
    options.text_shaping = True
    options.fallback_fonts = [str(arabic_font())]
    model = ldm.Document(sections=[ldm.Section(page_setup=ldm.PageSetup(
        page_width=400, page_height=250, left_margin=20, right_margin=20,
        top_margin=20, bottom_margin=20), body=ldm.Body(children=[
            ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(size=28))])]))])
    raw = LdmPdfWriter(options).write_to_bytes(model)
    native = FPDF(format=(400 * PT_TO_MM, 250 * PT_TO_MM))
    try:
        native.set_margins(20 * PT_TO_MM, 20 * PT_TO_MM, 20 * PT_TO_MM)
        native.set_auto_page_break(True, margin=20 * PT_TO_MM)
        register_fonts(native, model, options.fallback_fonts)
        native.set_text_shaping(True)
        native.add_page()
        native.set_font('DocumentSansSC', size=28)
        native.multi_cell(native.epw, 28 * 1.4 * PT_TO_MM, text)
        reference = bytes(native.output())
    finally:
        close_fonts(native)

    def glyph_positions(data):
        with pymupdf.open(stream=data, filetype='pdf') as pdf:
            return sorted((char['c'], *char['bbox']) for block in pdf[0].get_text('rawdict')['blocks']
                          for line in block.get('lines', []) for span in line['spans']
                          for char in span['chars'] if not char['c'].isspace())

    actual, expected = glyph_positions(raw), glyph_positions(reference)
    assert [item[0] for item in actual] == [item[0] for item in expected]
    # Separate PDF text objects round positions slightly differently.
    for item, target in zip(actual, expected, strict=True):
        assert item[1:] == pytest.approx(target[1:], abs=0.05)


@pytest.mark.parametrize('bold,italic', [(False, False), (True, False), (False, True), (True, True)])
def test_raw_font_tables_preserve_native_shaped_output(monkeypatch, bold, italic):
    pytest.importorskip('uharfbuzz')
    import aspose.words_foss.pdf_writer.writer as writer_module

    options = aw.saving.PdfSaveOptions()
    options.text_shaping = True
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[ldm.Paragraph(
        children=[ldm.Run(text='AVATAR office e\u0301 中文繁體 12345 !?',
                          font=ldm.Font(size=22, bold=bold, italic=italic))])]))])
    optimized = LdmPdfWriter(options).write_to_bytes(model)
    original = writer_module.register_fonts

    def native_fonts(*args, **kwargs):
        families = original(*args, **kwargs)
        for font in args[0].fonts.values():
            font._hbfont = None
        return families

    monkeypatch.setattr(writer_module, 'register_fonts', native_fonts)
    reference = LdmPdfWriter(options).write_to_bytes(model)
    with pymupdf.open(stream=optimized, filetype='pdf') as actual, \
            pymupdf.open(stream=reference, filetype='pdf') as native:
        assert len(actual) == len(native) == 1
        assert actual[0].get_text('words') == native[0].get_text('words')
        assert actual[0].get_pixmap(matrix=pymupdf.Matrix(2, 2)).samples == \
            native[0].get_pixmap(matrix=pymupdf.Matrix(2, 2)).samples
    actual_fonts = PdfReader(BytesIO(optimized)).pages[0]['/Resources']['/Font'].values()
    native_fonts = PdfReader(BytesIO(reference)).pages[0]['/Resources']['/Font'].values()
    for actual_font, native_font in zip(actual_fonts, native_fonts, strict=True):
        actual_font, native_font = actual_font.get_object(), native_font.get_object()
        assert actual_font['/ToUnicode'].get_data() == native_font['/ToUnicode'].get_data()
        assert actual_font['/DescendantFonts'][0].get_object()['/FontDescriptor']['/FontFile2'].get_data() == \
            native_font['/DescendantFonts'][0].get_object()['/FontDescriptor']['/FontFile2'].get_data()


def test_shaped_font_registration_does_not_reserialize_font(monkeypatch):
    pytest.importorskip('uharfbuzz')

    def reject_save(*args, **kwargs):
        raise AssertionError('Full font serialization is unnecessary for shaping')

    monkeypatch.setattr(TTFont, 'save', reject_save)
    pdf = FPDF()
    pdf.set_text_shaping(True)
    try:
        register_fonts(pdf)
        assert len(pdf.fonts) == 4
        for font in pdf.fonts.values():
            assert font._hbfont is not None
            assert font.hbfont.face.glyph_count == len(font.ttfont.getGlyphOrder())
    finally:
        close_fonts(pdf)
    assert all(font._hbfont is None for font in pdf.fonts.values())
