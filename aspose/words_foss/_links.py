"""Shared encoded-link syntax for LDM readers and conversion writers."""

import re
from string import punctuation

from aspose.words_foss.md_import.document_builder import _link_destination

_ESCAPED_PUNCTUATION = re.compile(r"\\([" + re.escape(punctuation) + r"])")

# Reuse the Markdown writer's destination/title grammar, including escaped brackets.
INLINE_LINK_RE = re.compile(
    r"(?<!!)\[((?:\\.|[^\[\]\\])+)\]"
    r"\((<[^<>\n]*>|(?:\\.|[^\s()\\]|\((?:\\.|[^()\\])*\))*)"
    r"((?:[ \t]+(?:\"(?:\\.|[^\"\\])*\"|'(?:\\.|[^'\\])*'|\((?:\\.|[^()\\])*\)))?)"
    r"[ \t]*\)"
)


def _escape_label(label):
    return "".join("\\" + char if char in punctuation else char for char in label)


def format_link(label, target):
    label = _escape_label(label)
    target = target.replace("\\", "\\\\").replace("&", "\\&")
    return f"[{label}]({_link_destination(target)})"


def decode_link(match):
    target = match.group(2)
    if target.startswith("<") and target.endswith(">"):
        target = target[1:-1]
    return tuple(_ESCAPED_PUNCTUATION.sub(r"\1", value) for value in (match.group(1), target))
