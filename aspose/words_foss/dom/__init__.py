"""Original-package DOCX DOM, separate from the conversion-oriented LDM."""

from aspose.words_foss.dom.document import DocxDocument
from aspose.words_foss.dom.nodes import (
    Body,
    Cell,
    Comment,
    Footnote,
    StoryPart,
    Font,
    HeaderFooter,
    Hyperlink,
    Node,
    Paragraph,
    ParagraphFormat,
    Row,
    RowFormat,
    Run,
    Table,
    UnknownNode,
)
from aspose.words_foss.dom.ranges import TextRange
from aspose.words_foss.dom.styles import (
    EffectiveFont,
    EffectiveParagraphFormat,
    Style,
    StyleCollection,
    StyleFont,
    StyleParagraphFormat,
)

__all__ = [
    "Body",
    "Cell",
    "Comment",
    "Footnote",
    "StoryPart",
    "DocxDocument",
    "EffectiveFont",
    "EffectiveParagraphFormat",
    "Font",
    "HeaderFooter",
    "Hyperlink",
    "Node",
    "Paragraph",
    "ParagraphFormat",
    "Row",
    "RowFormat",
    "Run",
    "Style",
    "StyleCollection",
    "StyleFont",
    "StyleParagraphFormat",
    "Table",
    "TextRange",
    "UnknownNode",
]
