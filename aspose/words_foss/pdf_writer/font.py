"""Font registration, application and reset helpers for the PDF writer."""
from io import BytesIO
from pathlib import Path

from fontTools.ttLib.sfnt import SFNTWriter
from fpdf import FPDF
from fpdf.enums import CharVPos

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.color import set_text_color
from aspose.words_foss.pdf_writer.constants import DEFAULT_FONT_NAME, DEFAULT_FONT_SIZE_PT
from aspose.words_foss.pdf_writer.diagnostics import document_nodes


def close_fonts(pdf: FPDF) -> None:
    """Release font readers even when fpdf2 never reaches its output cleanup."""
    for font in pdf.fonts.values():
        if font.type == 'TTF':
            font.subset.pick.cache_clear()
            font.subset.get_glyph.cache_clear()
            font.close()


def register_fonts(pdf: FPDF, doc=None, fallback_fonts=()) -> list[str]:
    """Load only required styles; WOFF resources retain all original glyphs and metrics."""
    styles = {""} if doc is not None else {"", "B", "I", "BI"}
    if doc is not None:
        for node in document_nodes(doc):
            if isinstance(node, ldm.Run):
                styles.add(("B" if node.font.bold else "") + ("I" if node.font.italic else ""))
            elif isinstance(node, ldm.Paragraph):
                if node.paragraph_format.is_heading:
                    styles.add("B")
                if "Quote" in (node.paragraph_format.style_name or ""):
                    styles.add("I")
    font_dir = Path(__file__).with_name("fonts")
    for style, suffix in (("", "Regular"), ("B", "Bold"), ("I", "Oblique"), ("BI", "BoldOblique")):
        if style in styles:
            pdf.add_font(DEFAULT_FONT_NAME, style, str(font_dir / f"DocumentSansSC-{suffix}.woff"))
            if pdf.text_shaping:
                from fpdf.fonts import HarfBuzzFont
                import uharfbuzz as hb

                font = pdf.fonts[DEFAULT_FONT_NAME.lower() + style]
                # shortcut: fpdf2's private cache avoids WOFF recompilation; recheck on upgrades.
                reader = font.ttfont.reader
                buffer = BytesIO()
                sfnt = SFNTWriter(buffer, len(reader.tables), reader.sfntVersion)
                for tag in reader.keys():
                    sfnt[tag] = reader[tag]
                sfnt.close()
                font._hbfont = HarfBuzzFont(hb.Face(buffer.getvalue()))
    families = []
    for index, path in enumerate(fallback_fonts):
        family = f"Fallback{index}"
        pdf.add_font(family, fname=str(path))
        families.append(family)
    if families:
        pdf.set_fallback_fonts(families, exact_match=False)
    return families


def apply_run_font(pdf: FPDF, font: ldm.Font, default_size: float = DEFAULT_FONT_SIZE_PT) -> str:
    """Apply font properties from an LDM Font and return the fpdf2 style string."""
    style = ""
    if font.bold:
        style += "B"
    if font.italic:
        style += "I"
    if font.underline:
        style += "U"
    if font.strike_through:
        style += "S"

    size = font.size if font.size > 0 else default_size
    # ponytail: one Unicode family replaces source fonts; add font matching for layout fidelity.
    pdf.char_vpos = CharVPos.SUP if font.superscript else CharVPos.SUB if font.subscript else CharVPos.LINE
    pdf.set_font(DEFAULT_FONT_NAME, style=style, size=size)
    set_text_color(pdf, font.color)
    return style


def reset_font(pdf: FPDF) -> None:
    """Reset font to defaults."""
    pdf.char_vpos = CharVPos.LINE
    pdf.set_font(DEFAULT_FONT_NAME, size=DEFAULT_FONT_SIZE_PT)
    pdf.set_text_color(0, 0, 0)
