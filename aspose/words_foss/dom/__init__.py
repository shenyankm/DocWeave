"""Original-package DOCX DOM, separate from the conversion-oriented LDM."""

from aspose.words_foss.dom.document import DocxDocument
from aspose.words_foss.dom.nodes import (
    Body,
    Cell,
    Font,
    HeaderFooter,
    Hyperlink,
    Node,
    Paragraph,
    ParagraphFormat,
    Row,
    Run,
    Table,
    UnknownNode,
)
from aspose.words_foss.dom.ranges import TextRange
from aspose.words_foss.dom.styles import EffectiveFont, EffectiveParagraphFormat

__all__ = [
    "Body", "Cell", "DocxDocument", "EffectiveFont", "EffectiveParagraphFormat", "Font",
    "HeaderFooter", "Hyperlink", "Node", "Paragraph", "ParagraphFormat", "Row", "Run", "Table",
    "TextRange", "UnknownNode",
]
