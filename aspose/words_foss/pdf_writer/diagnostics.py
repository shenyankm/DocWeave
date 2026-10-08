"""Actionable warnings for PDF conversion instead of silent loss."""

import warnings

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.constants import DEFAULT_FONT_NAME
from aspose.words_foss.pdf_writer.text import apply_caps, extract_link_segments, safe_text
from aspose.words_foss.saving import PdfCompliance


class PdfConversionWarning(UserWarning):
    """PDF output differs from a requested or source feature."""


class PdfMissingGlyphWarning(PdfConversionWarning):
    """No configured font can display a source character."""


class PdfFontSubstitutionWarning(PdfConversionWarning):
    """Original font metrics or styling are not retained."""


class PdfUnsupportedOptionWarning(PdfConversionWarning):
    """An explicitly requested PDF option is not implemented."""


def document_nodes(doc):
    """Include text-box paragraphs, which the generic LDM node walker excludes."""
    stack = [doc]
    while stack:
        node = stack.pop()
        yield node
        stack.extend(ldm._walk_children(node))
        if isinstance(node, ldm.Shape) and node.text_box:
            stack.extend(node.text_box.get("paragraphs", []) or [])
        stack.extend(c for c in getattr(node, "children", []) if isinstance(c, ldm.UnknownNode))


def warn_about_conversion(pdf, doc, options, fallback_families):
    unsupported = {
        "text_compression",
        "embed_full_fonts",
        "use_core_fonts",
        "font_embedding_mode",
        "page_mode",
        "color_mode",
        "preserve_form_fields",
        "memory_optimization",
    } & options._explicit_options
    if unsupported:
        warnings.warn(
            "PDF options are not applied: " + ", ".join(sorted(unsupported)),
            PdfUnsupportedOptionWarning,
            stacklevel=3,
        )
    if options.compliance not in (PdfCompliance.PDF17, PdfCompliance.PDF20):
        warnings.warn(
            "Only the PDF version is set; PDF/A and PDF/UA conformance is not implemented",
            PdfUnsupportedOptionWarning,
            stacklevel=3,
        )

    source_fonts, missing, unknown = set(), set(), False
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
                ord(c) for c in safe_text(node.list_label.label_string) if c.isprintable()
            }
            missing.update(
                label_codes - pdf.fonts[DEFAULT_FONT_NAME.lower()].cmap.keys() - fallback_coverage
            )
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
        codes = {ord(char) for char in text if char.isprintable()}
        missing.update(codes - coverage.keys() - fallback_coverage)
        fallback_style_loss |= bool(style and (codes - coverage.keys()) & fallback_coverage)
    if source_fonts:
        warnings.warn(
            "PDF replaces source fonts with Document Sans SC; layout may differ. Sources: "
            + ", ".join(sorted(source_fonts)[:12]),
            PdfFontSubstitutionWarning,
            stacklevel=3,
        )
    if fallback_style_loss:
        warnings.warn(
            "Fallback glyphs use regular faces; their bold/italic styling may change",
            PdfFontSubstitutionWarning,
            stacklevel=3,
        )
    if unknown:
        warnings.warn(
            "Unknown document nodes are omitted from PDF", PdfConversionWarning, stacklevel=3
        )
    if missing:
        codes = ", ".join(f"U+{value:04X}" for value in sorted(missing)[:20])
        warnings.warn(
            f"PDF has {len(missing)} unsupported character(s): {codes}. Configure fallback_fonts.",
            PdfMissingGlyphWarning,
            stacklevel=3,
        )
