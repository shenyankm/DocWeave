"""Font registration, application and reset helpers for the PDF writer."""

from pathlib import Path

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.color import set_text_color
from aspose.words_foss.pdf_writer.constants import DEFAULT_FONT_NAME, DEFAULT_FONT_SIZE_PT


def register_fonts(pdf: FPDF) -> None:
    """Register bundled Unicode fonts before any page/header/footer is rendered."""
    font_dir = Path(__file__).with_name("fonts")
    for style, suffix in (("", "Regular"), ("B", "Bold"), ("I", "Oblique"), ("BI", "BoldOblique")):
        pdf.add_font(DEFAULT_FONT_NAME, style, str(font_dir / f"DocumentSansSC-{suffix}.ttf"))


def apply_run_font(pdf: FPDF, font: ldm.Font, default_size: float = DEFAULT_FONT_SIZE_PT) -> str:
    """Apply font properties from an LDM Font and return the fpdf2 style string."""
    style = ""
    if font.bold:
        style += "B"
    if font.italic:
        style += "I"
    if font.underline:
        style += "U"

    size = font.size if font.size > 0 else default_size
    # ponytail: one Unicode family replaces source fonts; add font matching for layout fidelity.
    pdf.set_font(DEFAULT_FONT_NAME, style=style, size=size)
    set_text_color(pdf, font.color)
    return style


def reset_font(pdf: FPDF) -> None:
    """Reset font to defaults."""
    pdf.set_font(DEFAULT_FONT_NAME, size=DEFAULT_FONT_SIZE_PT)
    pdf.set_text_color(0, 0, 0)
