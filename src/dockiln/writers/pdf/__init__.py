"""PDF writer module — backward-compatible public API.

Re-exports ``LdmPdfWriter`` and ``_parse_color`` so existing imports
of the form ``from dockiln.writers.pdf import LdmPdfWriter``
continue to work without changes.
"""

from dockiln.writers.pdf.color import parse_color as _parse_color
from dockiln.writers.pdf.writer import LdmPdfWriter
from dockiln.writers.pdf.diagnostics import (
    PdfContentLossWarning, PdfConversionWarning, PdfMissingGlyphWarning, PdfFontSubstitutionWarning, PdfUnsupportedOptionWarning,
)

__all__ = [
    "LdmPdfWriter", "_parse_color", "PdfConversionWarning", "PdfMissingGlyphWarning",
    "PdfFontSubstitutionWarning", "PdfUnsupportedOptionWarning", "PdfContentLossWarning",
]
