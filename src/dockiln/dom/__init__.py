"""Original-package DOCX DOM, separate from the conversion-oriented LDM."""

from dockiln.dom.document import DocxDocument
from dockiln.dom.nodes import (
    Body,
    Cell,
    Comment,
    CommentRangeStart,
    CommentRangeEnd,
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
from dockiln.dom.ranges import TextRange
from dockiln.dom.styles import (
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
    "CommentRangeStart",
    "CommentRangeEnd",
    "Footnote",
    "FootnoteType",
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

from dockiln.dom.notes import FootnoteType
