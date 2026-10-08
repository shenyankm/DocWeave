"""Literal, within-run DOCX text edits that retain the original OOXML package.

Not a general document editor or a sanitizer. Unknown parts, macros and external
relationships are retained, not executed. Use only in an appropriately isolated application.
"""

from pathlib import Path
import re
from xml.parsers import expat
from xml.sax.saxutils import escape
from zipfile import ZipFile

from defusedxml.ElementTree import fromstring

from aspose.words_foss._io import atomic_output, check_input_size, validate_docx_archive

_WORD_TEXT = "http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"


def _edit_part(data: bytes, pattern, replacements, found: set[str]) -> tuple[bytes, int]:
    fromstring(data, forbid_dtd=True)
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        raise ValueError("Original-package text edits require UTF-8 XML")
    declaration = re.match(rb"(?:\xef\xbb\xbf)?<\?xml\b[^?]*\?>", data)
    if declaration:
        encoding = re.search(rb'encoding\s*=\s*[\'\"]([^\'\"]+)', declaration.group())
        if encoding and encoding[1].lower().replace(b"-", b"") != b"utf8":
            raise ValueError("Original-package text edits require UTF-8 XML")
    parser = expat.ParserCreate(namespace_separator="}")
    active, text, patches = None, [], []
    count = 0

    def start(name, attrs):
        nonlocal active, text
        if active is not None:
            raise ValueError("Nested elements inside Word text nodes are unsupported")
        if name == _WORD_TEXT:
            begin = parser.CurrentByteIndex
            quote = None
            for i in range(begin, len(data)):
                char = data[i]
                if quote:
                    if char == quote:
                        quote = None
                elif char in (34, 39):
                    quote = char
                elif char == 62:
                    active = (begin, i + 1)
                    break
            text = []

    def characters(value):
        if active is not None:
            text.append(value)

    def end(name):
        nonlocal active, count
        if name != _WORD_TEXT or active is None:
            return
        source = "".join(text)

        def substitute(match):
            found.add(match.group())
            return replacements[match.group()]

        updated, changed = pattern.subn(substitute, source)
        if changed:
            begin, content = active
            opening = data[begin:content]
            if updated[:1].isspace() or updated[-1:].isspace():
                space = rb'\bxml:space\s*=\s*([\'\"])[^\'\"]*\1'
                if re.search(space, opening):
                    opening = re.sub(space, b'xml:space="preserve"', opening)
                else:
                    opening = opening[:-1] + b' xml:space="preserve">'
            patches.append((begin, parser.CurrentByteIndex, opening + escape(updated).encode("utf-8")))
            count += changed
        active = None

    parser.StartElementHandler = start
    parser.CharacterDataHandler = characters
    parser.EndElementHandler = end
    parser.Parse(data, True)
    pieces, offset = [], 0
    for begin, end, replacement in patches:
        pieces.extend((data[offset:begin], replacement))
        offset = end
    pieces.append(data[offset:])
    return b"".join(pieces), count


def replace_text(source: str | Path, destination: str | Path, replacements: dict[str, str]) -> int:
    """Replace literals simultaneously within Word text nodes; return match count.

    All keys must occur. Cross-run matches and signed packages are rejected before
    publication. Unedited XML bytes and all other part payloads are retained verbatim.
    """
    if not isinstance(replacements, dict) or not replacements or any(not isinstance(key, str) or not key or not isinstance(value, str)
                               for key, value in replacements.items()):
        raise ValueError("replacements must be a nonempty mapping of nonempty strings to strings")
    for value in replacements.values():
        if any(ord(char) < 32 or 0xD800 <= ord(char) <= 0xDFFF or ord(char) in (0xFFFE, 0xFFFF)
               for char in value):
            raise ValueError("Replacement text must contain XML characters, without tabs or line breaks")
    source = Path(source)
    check_input_size(source.stat().st_size)
    pattern = re.compile("|".join(re.escape(key) for key in sorted(replacements, key=len, reverse=True)))
    found, parts, count = set(), [], 0
    with ZipFile(source) as archive:
        validate_docx_archive(archive)
        if "word/document.xml" not in archive.namelist():
            raise ValueError("Source is not a DOCX package")
        if any(name.lower().startswith("_xmlsignatures/") for name in archive.namelist()):
            raise ValueError("Signed DOCX packages cannot be edited without invalidating signatures")
        for entry in archive.infolist():
            data = archive.read(entry)
            if entry.filename.endswith(".rels"):
                relationships = fromstring(data, forbid_dtd=True)
                if any(link.get("Type", "").startswith(
                        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/")
                       for link in relationships):
                    raise ValueError("Signed DOCX packages cannot be edited without invalidating signatures")
            if entry.filename.startswith("word/") and entry.filename.endswith(".xml"):
                data, changed = _edit_part(data, pattern, replacements, found)
                count += changed
            parts.append((entry, data))
        comment = archive.comment
    missing = replacements.keys() - found
    if missing:
        raise ValueError("Text not found within a single Word text node (possibly split across runs): "
                         + ", ".join(sorted(missing)))
    with atomic_output(destination) as temporary:
        with ZipFile(temporary, "w") as output:
            output.comment = comment
            for entry, data in parts:
                output.writestr(entry, data)
    return count
