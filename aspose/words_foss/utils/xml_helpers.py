"""XML helper utilities for parsing DOCX files."""

import re
from decimal import Decimal
from typing import Iterator, Mapping, Optional
from xml.etree import ElementTree as ET

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def serialize_xml(node) -> bytes:
    """Serialize DOM attributes with XML-safe characters and explicit whitespace references."""
    from io import StringIO
    from xml.dom import Node
    from aspose.words_foss.docx_writer.xml_utils import escape_attr

    output = StringIO()

    def write(current):
        if current.nodeType == Node.DOCUMENT_NODE:
            output.write('<?xml version="1.0" encoding="utf-8"?>')
            for child in current.childNodes:
                write(child)
        elif current.nodeType == Node.ELEMENT_NODE:
            output.write('<' + current.tagName)
            for index in range(current.attributes.length):
                attr = current.attributes.item(index)
                output.write(' ' + attr.name + '="' + escape_attr(attr.value or '') + '"')
            if current.childNodes:
                output.write('>')
                for child in current.childNodes:
                    write(child)
                output.write('</' + current.tagName + '>')
            else:
                output.write('/>')
        else:
            current.writexml(output)

    write(node)
    return output.getvalue().encode('utf-8')


def combine_style_toggle(default: bool, paragraph: bool | None, character: bool | None) -> bool:
    """Combine nearest style-category values observed in the fixed 26.9 corpus."""
    if paragraph is not None and character is not None:
        return default or (paragraph ^ character)
    return paragraph if paragraph is not None else character if character is not None else default


def parse_font_size(value: str) -> float:
    """Read nonnegative ordinary OOXML sizes, quantized to whole half-points."""
    try:
        match = re.fullmatch(r"([+]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?)(pt|in|cm|mm|pc|pi)?", value)
        if match is None:
            raise ValueError
        number, unit = match.groups()
        if unit is None and number.lstrip("+").isdigit():
            half_points = int(number)
        else:
            if not 0 <= float(number) <= 2**64 - 1:
                raise ValueError
            numerator, denominator = Decimal(number).as_integer_ratio()
            scale, divisor = {None: (1, 1), "pt": (2, 1), "in": (144, 1),
                              "cm": (7200, 127), "mm": (720, 127),
                              "pc": (24, 1), "pi": (24, 1)}[unit]
            half_points = numerator * scale // (denominator * divisor)
        if not 0 <= half_points <= 2**64 - 1:
            raise ValueError
        return half_points / 2
    except (ValueError, OverflowError):
        raise ValueError("Expected a nonnegative representable OOXML font size") from None


def get_element_text(element: ET.Element, default: str = "") -> str:
    """Get text content from an XML element."""
    if element is None:
        return default
    return element.text or default


def find_all_elements(root: ET.Element, tag: str) -> Iterator[ET.Element]:
    """Find all elements with the given tag."""
    for elem in root.iter(tag):
        yield elem


def get_attribute(element: ET.Element, attr: str, default: str = "") -> str:
    """Get an attribute value from an element."""
    if element is None:
        return default
    return element.get(f"{W_NS}{attr}", default)


def has_element(parent: ET.Element, tag: str) -> bool:
    """Check if parent has a child element with the given tag."""
    return parent.find(f"{W_NS}{tag}") is not None


def get_bool_val(element: Optional[ET.Element]) -> bool:
    """Get boolean value from an element with optional val attribute."""
    if element is None:
        return False
    val = element.get(f"{W_NS}val")
    if val is None:
        return True  # Presence without val means true
    return val.lower() not in ("false", "0", "off")


def normalize_font_names(attributes: Mapping[str, str]) -> dict[str, str]:
    """Canonical local rFonts winners observed in the fixed native 26.9 corpus."""
    themes = {'asciiTheme': 'ascii', 'hAnsiTheme': 'hAnsi',
              'cstheme': 'cs', 'eastAsiaTheme': 'eastAsia'}
    known = {prefix + suffix for prefix in ('major', 'minor')
             for suffix in ('Ascii', 'HAnsi', 'Bidi', 'EastAsia')}
    winners = {}
    unknown = set()
    for raw_key, value in attributes.items():
        key = raw_key.removeprefix(W_NS)
        channel = themes.get(key, key)
        if key not in themes and channel not in themes.values():
            winners[key] = value
            continue
        if not value:
            continue
        if key in themes and value not in known:
            unknown.add(channel)
            continue
        winners.pop(channel, None)
        for theme, owner in themes.items():
            if owner == channel:
                winners.pop(theme, None)
        winners[key] = value
    for channel in unknown:
        if channel not in winners and not any(theme in winners and owner == channel for theme, owner in themes.items()):
            # Unknown nonempty tokens mask inheritance, unlike an empty declaration.
            winners[channel] = 'Times New Roman'
    return winners


def resolve_font_names(attributes: Mapping[str, str], theme_fonts: Mapping[str, str]) -> dict[str, str]:
    """Resolve canonical winners without changing the raw declaration mapping."""
    themes = {'asciiTheme': 'ascii', 'hAnsiTheme': 'hAnsi',
              'cstheme': 'cs', 'eastAsiaTheme': 'eastAsia'}
    result = {}
    for key, value in normalize_font_names(attributes).items():
        if key in themes:
            result[themes[key]] = theme_fonts.get(value) or 'Times New Roman'
        elif key in themes.values():
            result[key] = value
    return result
