"""Selected CommonMark/GFM examples for the supported Markdown subset, not full compliance."""

from io import BytesIO

import pytest

import aspose.words_foss as aw
from aspose.words_foss.md_import.block_parser import parse_document
from aspose.words_foss.md_import.blocks import (
    AtxHeadingBlock, FencedCodeBlock, InlineCodeBlock, LinkTextBlock, ListBlock,
    QuoteBlock, SetextHeadingBlock, TableBlock, TextBlock,
)
from aspose.words_foss.md_import.inline_parser import parse_inline


def walk(block):
    yield block
    for child in block.children:
        yield from walk(child)


def texts(block):
    return "".join(node.text for node in walk(block) if isinstance(node, TextBlock))


@pytest.mark.parametrize("source,expected", [
    (r"\*literal\*", "*literal*"),
    ("a &amp; b &#35;", "a & b #"),
    ("***bold italic***", "bold italic"),
    ("a_b_c", "a_b_c"),
    ("**strong _nested_**", "strong nested"),
    ("[label](https://example.com/a_(b))", "label"),
    ("[label][ref]\n\n[ref]: https://example.com", "label"),
])
def test_inline_content_examples(source, expected):
    root = parse_document(source)
    assert texts(root) == expected


@pytest.mark.parametrize("source", ["# heading", "   # heading", "### heading ###"])
def test_atx_heading_examples(source):
    root = parse_document(source)
    assert isinstance(root.children[0], AtxHeadingBlock)
    assert texts(root) == "heading"


def test_setext_heading_and_escaped_punctuation():
    root = parse_document("Heading\n=======\n\n\\# not a heading")
    assert isinstance(root.children[0], SetextHeadingBlock)
    assert texts(root) == "Heading# not a heading"


@pytest.mark.parametrize("source,expected", [
    ("`` a ` b ``", "a ` b"),
    ("` foo   bar `", "foo   bar"),
    ("`foo\nbar`", "foo bar"),
])
def test_code_span_examples(source, expected):
    nodes = parse_inline(source)
    code = next(node for node in nodes if isinstance(node, InlineCodeBlock))
    assert code.code == expected


def test_fenced_code_is_literal_and_nested_fence_is_not_a_close():
    root = parse_document("~~~~ python\n**literal**\n```\n~~~~")
    assert isinstance(root.children[0], FencedCodeBlock)
    assert root.children[0].code == "**literal**\n```"


def test_link_destinations_with_escaped_parentheses():
    root = parse_document(r"[label](https://example.com/a\(b\))")
    link = next(node for node in walk(root) if isinstance(node, LinkTextBlock))
    assert link.uri == "https://example.com/a(b)"


def test_quote_and_nested_list_structure():
    root = parse_document("> quoted\n>\n> - first\n>   - nested\n\n1. ordered\n2. next")
    assert isinstance(root.children[0], QuoteBlock)
    lists = [node for node in walk(root) if isinstance(node, ListBlock)]
    assert len(lists) == 3 and not lists[0].ordered and lists[-1].ordered
    assert texts(root) == "quotedfirstnestedorderednext"


def test_gfm_table_with_escaped_pipe_and_alignment():
    root = parse_document("| A \\| B | C |\n| :--- | ---: |\n| 1 | 2 |")
    grid = next(node for node in root.children if isinstance(node, TableBlock))
    assert texts(grid) == "A | BC12"


def test_nested_emphasis_survives_public_docx_round_trip():
    doc = aw.Document(BytesIO(b"**strong _nested_**"), aw.MarkdownLoadOptions())
    loaded = aw.Document(BytesIO(doc.to_bytes("docx")))
    assert loaded.get_text().strip() == "strong nested"
    nested = next(run for run in loaded.get_child_nodes(aw.NodeType.RUN, True) if "nested" in run.text)
    assert nested.font.bold and nested.font.italic
