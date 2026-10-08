"""Optional shaping uses real HarfBuzz glyphs, not merely Unicode font coverage."""

from io import BytesIO
from pathlib import Path
import warnings

from fontTools.ttLib import TTFont
import pymupdf
from pypdf import PdfReader
import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import PdfMissingGlyphWarning
from aspose.words_foss.pdf_writer.writer import LdmPdfWriter


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
        font=ldm.Font(size=14, highlight_color="Color [Yellow]" if highlight else ""))])
    para.paragraph_format.alignment = alignment
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[para]))])
    writer = LdmPdfWriter(options)
    raw = writer.write_to_bytes(model)
    estimate = writer._estimate_paragraph_height(para, 60)
    measure = writer._measurement_pdf
    shadow = writer._measurement_writer
    measure.set_xy(0, 0)
    measure.set_margins(0, 0, measure.w - 60)
    shadow._paragraph_renderer.render_paragraph(measure, para)
    assert measure.get_y() == pytest.approx(estimate, abs=0.001)
    assert len(PdfReader(BytesIO(raw)).pages) > 0
