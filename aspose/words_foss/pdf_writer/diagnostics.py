"""Actionable warnings for PDF conversion instead of silent loss."""

import re
import unicodedata
from aspose.words_foss.diagnostics import ContentLossWarning, ConversionWarning, document_nodes, header_footer_losses, source_story_losses, warn

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.constants import DEFAULT_FONT_NAME
from aspose.words_foss.pdf_writer.text import apply_caps, extract_link_segments, safe_text
from aspose.words_foss.saving import PdfCompliance


class PdfConversionWarning(ConversionWarning):
    """PDF output differs from a requested or source feature."""

    code = "pdf.content_loss"


class PdfContentLossWarning(PdfConversionWarning, ContentLossWarning):
    """Known document content or layout is omitted from PDF."""


class PdfMissingGlyphWarning(PdfConversionWarning):
    """No configured font can display a source character."""

    code = "pdf.missing_glyph"


class PdfFontSubstitutionWarning(PdfConversionWarning):
    """Original font metrics or styling are not retained."""

    code = "pdf.font_substitution"


class PdfUnsupportedOptionWarning(PdfConversionWarning):
    """An explicitly requested PDF option is not implemented."""

    code = "pdf.unsupported_option"


def warn_about_conversion(pdf, doc, options, fallback_families):
    unsupported = {
        "embed_full_fonts",
        "use_core_fonts",
        "font_embedding_mode",
        "color_mode",
        "preserve_form_fields",
        "memory_optimization",
    } & options._explicit_options
    if unsupported:
        warn(
            "PDF options are not applied: " + ", ".join(sorted(unsupported)),
            PdfUnsupportedOptionWarning,
            stacklevel=3,
        )
    if options.compliance not in (PdfCompliance.PDF17, PdfCompliance.PDF20):
        warn(
            "Only the PDF version is set; PDF/A and PDF/UA conformance is not implemented",
            PdfUnsupportedOptionWarning,
            stacklevel=3,
        )

    for code, message in (*header_footer_losses(doc), *source_story_losses(doc)):
        warn(message, PdfContentLossWarning, stacklevel=3, code="pdf." + code)
    source_fonts, missing, unknown = set(), set(), False
    needs_shaping = False
    fallback_style_loss = False
    fallback_coverage = set().union(*(pdf.fonts[name.lower()].cmap for name in fallback_families))
    for node in document_nodes(doc):
        if (
            isinstance(node, ldm.Paragraph)
            and node.list_format
            and node.list_format.is_list_item
            and node.list_label
        ):
            label_codes = {
                ord(c) for c in safe_text(node.list_label.label_string)
                if c.isprintable() or unicodedata.category(c) in ("Cn", "Co")
            }
            coverage = pdf.fonts[DEFAULT_FONT_NAME.lower()].cmap
            missing.update(code for code in label_codes
                           if code not in coverage and code not in fallback_coverage)
        if isinstance(node, ldm.UnknownNode):
            unknown = True
        if not isinstance(node, ldm.Run) or node.font.hidden:
            continue
        font = node.font
        if font.name and font.name != DEFAULT_FONT_NAME:
            source_fonts.add(font.name)
        style = ("B" if font.bold else "") + ("I" if font.italic else "")
        coverage = pdf.fonts[DEFAULT_FONT_NAME.lower() + style].cmap
        text = safe_text(
            apply_caps("".join(chunk for chunk, _ in extract_link_segments(node.text)), font)
        )
        needs_shaping |= bool(re.search(r"[\u0590-\u0FFF]", text))
        # Newer Unicode and private-use characters still need glyphs on older Python.
        codes = {ord(char) for char in text
                 if char.isprintable() or unicodedata.category(char) in ("Cn", "Co")}
        # Check text characters, not the entire font cmap for every run.
        uncovered = {code for code in codes if code not in coverage}
        missing.update(uncovered - fallback_coverage)
        fallback_style_loss |= bool(style and uncovered & fallback_coverage)
    if needs_shaping and not options.text_shaping:
        warn("Complex-script or bidirectional text needs text_shaping=True and suitable fallback fonts",
             PdfContentLossWarning, stacklevel=3, code="pdf.shaping_disabled")
    if source_fonts:
        warn(
            "PDF replaces source fonts with Document Sans SC; layout may differ. Sources: "
            + ", ".join(sorted(source_fonts)[:12]),
            PdfFontSubstitutionWarning,
            stacklevel=3,
        )
    if fallback_style_loss:
        warn(
            "Fallback glyphs use regular faces; their bold/italic styling may change",
            PdfFontSubstitutionWarning,
            stacklevel=3,
        )
    if unknown:
        warn(
            "Unknown document nodes are omitted from PDF", PdfContentLossWarning, stacklevel=3
        )
    if missing:
        codes = ", ".join(f"U+{value:04X}" for value in sorted(missing)[:20])
        warn(
            f"PDF has {len(missing)} unsupported character(s): {codes}. Configure fallback_fonts.",
            PdfMissingGlyphWarning,
            stacklevel=3,
        )
