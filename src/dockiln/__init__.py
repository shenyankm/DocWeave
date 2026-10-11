"""DocKiln: document conversion, structured extraction, and DOCX editing."""

from dockiln.document import (
    Document,
    SaveFormat,
    LoadFormat,
    LoadOptions,
)
from dockiln import loading, saving, fonts, settings, notes
from dockiln.dom import DocxDocument
from dockiln.dom.importing import ImportFormatMode
from dockiln.diagnostics import ContentLossWarning, ConversionDiagnostic, ConversionWarning
from dockiln.model import wrap_type as wrap_type, enums as enums
from dockiln.light_document_model import NodeType  # noqa: F401
from dockiln.model.enums import (  # noqa: F401 — re-export for dk.ParagraphAlignment etc.
    CellMerge,
    CellVerticalAlignment,
    HeightRule,
    LineSpacingRule,
    LineStyle,
    NumberStyle,
    Orientation,
    ParagraphAlignment,
    SectionStart,
    StyleType,
    Underline,
)

__version__ = "0.1.0"
__all__ = [
    "Document",
    "DocxDocument",
    "ImportFormatMode",
    "SaveFormat",
    "LoadFormat",
    "LoadOptions",
    "NodeType",
    "ConversionDiagnostic",
    "ConversionWarning",
    "ContentLossWarning",
    "loading",
    "saving",
    "fonts",
    "settings",
    "notes",
]
