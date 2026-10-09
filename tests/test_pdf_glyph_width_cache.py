"""Repeated scalar measurements may be cached only for the same font state."""

import pytest
from fpdf import FPDF
from fpdf.enums import CharVPos

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.font import register_fonts
from aspose.words_foss.pdf_writer.run_renderer import RunRenderer


def setup_pdf():
    pdf = FPDF()
    register_fonts(pdf)
    pdf.add_page()
    pdf.set_font("DocumentSansSC", size=11)
    return pdf


def segments(text="中文 alpha 中文 alpha", size=11, bold=False):
    run = ldm.Run(text=text, font=ldm.Font(size=size, bold=bold))
    return [(run, text, 0, size, None)]


def test_repeated_wrap_reuses_glyph_widths(monkeypatch):
    pdf = setup_pdf()
    calls = []
    original = pdf.get_string_width

    def measure(text, *args, **kwargs):
        calls.append(text)
        return original(text, *args, **kwargs)

    monkeypatch.setattr(pdf, "get_string_width", measure)
    first = RunRenderer.wrap_segments(pdf, segments(), 45)
    count = len(calls)
    assert count > 0
    second = RunRenderer.wrap_segments(pdf, segments(), 45)
    assert len(calls) == count
    assert [[part[1:4] for part in row] for row in first] == [
        [part[1:4] for part in row] for row in second
    ]


@pytest.mark.parametrize(
    "change", ["size", "bold", "stretch", "spacing", "unit", "superscript", "fallback"]
)
def test_font_state_changes_do_not_reuse_stale_widths(change, monkeypatch):
    pdf = setup_pdf()
    RunRenderer.wrap_segments(pdf, segments(), 45)
    if change == "stretch":
        pdf.set_stretching(150)
    elif change == "spacing":
        pdf.set_char_spacing(0.8)
    elif change == "unit":
        pdf.k *= 1.5
    elif change == "superscript":
        pdf.char_vpos = CharVPos.SUP
    elif change == "fallback":
        pdf.set_fallback_fonts(["DocumentSansSC"], exact_match=False)
    values = segments(size=17 if change == "size" else 11, bold=change == "bold")
    cached = RunRenderer.wrap_segments(pdf, values, 45)
    pdf.__dict__.pop("_plain_glyph_widths", None)
    fresh = RunRenderer.wrap_segments(pdf, values, 45)
    assert [[part[1:4] for part in row] for row in cached] == [
        [part[1:4] for part in row] for row in fresh
    ]
    if change in ("superscript", "fallback"):
        assert "_plain_glyph_widths" not in pdf.__dict__


def test_glyph_cache_is_bounded_and_lives_on_each_pdf():
    pdf = setup_pdf()
    text = "".join(chr(0x4E00 + i) for i in range(2200))
    for size in (9, 10, 11):
        RunRenderer.wrap_segments(pdf, segments(text, size), 1000)
    assert 0 < len(pdf._plain_glyph_widths) <= 4096
    assert "_plain_glyph_widths" not in setup_pdf().__dict__


def test_shaping_uses_native_context_without_scalar_cache():
    pytest.importorskip("uharfbuzz")
    pdf = FPDF()
    pdf.set_text_shaping(True)
    doc = ldm.Document(
        sections=[
            ldm.Section(
                body=ldm.Body(children=[ldm.Paragraph(children=[segments()[0][0]])])
            )
        ]
    )
    register_fonts(pdf, doc)
    pdf.add_page()
    pdf.set_font("DocumentSansSC", size=11)
    rows = RunRenderer.wrap_segments(pdf, segments(), 45)
    assert rows and "_plain_glyph_widths" not in pdf.__dict__
