"""Editable document-owned font registration metadata."""
from enum import IntEnum
from io import BytesIO
from itertools import islice
import re
from operator import index
from xml.dom import Node
from zipfile import ZipFile

from defusedxml.minidom import parseString

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._flat_opc import decode, is_flat_opc
from aspose.words_foss._io import check_input_size
from aspose.words_foss.utils.xml_helpers import serialize_xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

class FontFamily(IntEnum):
    AUTO = 0
    ROMAN = 1
    SWISS = 2
    MODERN = 3
    SCRIPT = 4
    DECORATIVE = 5

class FontPitch(IntEnum):
    DEFAULT = 0
    FIXED = 1
    VARIABLE = 2

def _children(node, name):
    return [
        child
        for child in node.childNodes
        if child.nodeType == Node.ELEMENT_NODE
        and child.namespaceURI == W
        and child.localName == name
    ]


def _charset(value):
    if not value:
        return 0
    text = value.strip()
    if re.fullmatch(r"(?:0[xX])?[0-9a-fA-F]+", text) is None:
        raise RuntimeError("Font table contains an invalid charset")
    digits = (text[2:] if text.lower().startswith("0x") else text).lstrip("0")
    if len(digits) > 8:
        raise RuntimeError("Font charset exceeds signed 32-bit storage")
    number = int(digits or "0", 16)
    return number - (1 << 32) if number >= 1 << 31 else number


def _not_true_type(value):
    if value in (None, "", "1", "true", "on"):
        return True
    if value in ("0", "false", "off"):
        return False
    raise RuntimeError("Font table contains an invalid TrueType declaration")


def _root(data):
    check_input_size(len(data))
    root = parseString(data, forbid_dtd=True).documentElement
    if root.namespaceURI != W or root.localName != "fonts":
        raise ValueError("Expected a WordprocessingML font table")
    for font in _children(root, "font"):
        for child in _children(font, "charset"):
            _charset(child.getAttributeNS(W, "val"))
        for child in _children(font, "notTrueType"):
            _not_true_type(child.getAttributeNS(W, "val"))
    return root


def _key(name):
    # .NET ordinal ignore-case does not expand sharp-s or fold Turkish I.
    return "".join(c.upper() if c != "ı" and len(c.upper()) == 1 else c for c in name)


def _find(root, name):
    return next(
        (
            font
            for font in _children(root, "font")
            if _key(font.getAttributeNS(W, "name")) == _key(name)
        ),
        None,
    )


def _normalize_registrations(root):
    seen = {}
    for font in list(_children(root, "font")):
        key = _key(font.getAttributeNS(W, "name"))
        if key in seen:
            first = seen[key]
            for child in list(font.childNodes):
                if child.nodeType != Node.ELEMENT_NODE:
                    continue
                previous = next((item for item in first.childNodes
                                 if item.nodeType == Node.ELEMENT_NODE
                                 and (item.namespaceURI, item.localName) == (child.namespaceURI, child.localName)), None)
                default = False
                if previous is not None and previous.namespaceURI == W:
                    value = previous.getAttributeNS(W, "val")
                    if previous.localName == "charset":
                        default = _charset(value) == 0
                    elif previous.localName == "notTrueType":
                        default = not _not_true_type(value)
                    elif previous.localName == "altName":
                        default = value == ""
                    elif previous.localName == "family":
                        default = value not in {"roman", "swiss", "modern", "script", "decorative"}
                    elif previous.localName == "pitch":
                        default = value not in {"fixed", "variable"}
                if previous is None or default:
                    if previous is not None:
                        first.removeChild(previous)
                    first.appendChild(child.cloneNode(True))
            root.removeChild(font)
        else:
            seen[key] = font
    return root


class FontInfoCollection:
    """Live registration metadata; saved projections retain original embedded data."""

    def __init__(self, document):
        self._document = document
        self._cache = {}
        self._cached_source = object()
        self._cached_view = None

    @property
    def _source(self):
        return self._document.source_font_table

    @property
    def _view(self):
        source = self._source
        if self._cached_source is not source:
            root = _root((source.current_registrations_data or source.data) if source is not None
                         else f'<w:fonts xmlns:w="{W}"/>'.encode())
            self._cached_view = _normalize_registrations(root)
            self._entries = {_key(font.getAttributeNS(W, "name")): font
                             for font in _children(self._cached_view, "font")}
            self._cached_source = source
        return self._cached_view

    def _lookup(self, name):
        self._view
        return self._entries.get(_key(name))

    @property
    def count(self):
        return len(_children(self._view, "font"))

    def __len__(self):
        return self.count

    def __iter__(self):
        return iter(
            [
                self.get_by_name(font.getAttributeNS(W, "name"))
                for font in _children(self._view, "font")
            ]
        )

    def __getitem__(self, position):
        position = index(position)
        entries = _children(self._view, "font")
        if not 0 <= position < len(entries):
            raise IndexError("list index out of range")
        return self.get_by_name(entries[position].getAttributeNS(W, "name"))

    def get_by_name(self, name):
        if name is None:
            raise RuntimeError("Font name must not be null")
        if not isinstance(name, str):
            raise TypeError("Font name requires str")
        font = self._lookup(name)
        if font is None:
            return None
        key = _key(font.getAttributeNS(W, "name"))
        if key not in self._cache:
            self._cache[key] = FontInfo(self, font, self._source)
        return self._cache[key]

    def contains(self, name):
        return self.get_by_name(name) is not None

    def _flag(self, name):
        flags = self._document.font_embedding
        value = getattr(flags, name) if flags is not None else None
        return bool(value)

    def _set_flag(self, name, value):
        if type(value) is not bool:
            raise TypeError("Font embedding flag requires bool")
        if self._document.font_embedding is None:
            self._document.font_embedding = ldm.FontEmbeddingSettings()
        setattr(
            self._document.font_embedding,
            name,
            value,
        )

    embed_true_type_fonts = property(
        lambda self: self._flag("embed_true_type_fonts"),
        lambda self, value: self._set_flag("embed_true_type_fonts", value),
    )
    embed_system_fonts = property(
        lambda self: self._flag("embed_system_fonts"),
        lambda self, value: self._set_flag("embed_system_fonts", value),
    )
    save_subset_fonts = property(
        lambda self: self._flag("save_subset_fonts"),
        lambda self, value: self._set_flag("save_subset_fonts", value),
    )

    def _update(self, name, tag, value, fallback):
        source = self._source
        original = _normalize_registrations(_root(source.data))
        current = (
            _normalize_registrations(_root(source.current_registrations_data))
            if source.current_registrations_data is not None
            else None
        )
        for root in (original, current):
            if root is None:
                continue
            font = _find(root, name)
            if font is None:
                if root is current:
                    continue  # An obsolete held registration stays detached.
                font = root.ownerDocument.importNode(fallback, True)
                for embed in list(font.childNodes):
                    if (
                        embed.nodeType == Node.ELEMENT_NODE
                        and embed.localName.startswith("embed")
                    ):
                        font.removeChild(embed)
                root.appendChild(font)
            for existing in _children(font, tag):
                font.removeChild(existing)
            if value is not None:
                child = root.ownerDocument.createElementNS(W, "w:" + tag)
                child.setAttributeNS(W, "w:val", value)
                font.appendChild(child)
        # Newly-created prefixes must be bound even when a source uses ns0.
        original.setAttribute("xmlns:w", W)
        data = serialize_xml(original)
        parts = tuple(
            part.model_copy(update={"data": data})
            if part.name == source.part_name
            else part
            for part in source.parts
        )
        if current is not None:
            current.setAttribute("xmlns:w", W)
        self._document.source_font_table = source.model_copy(
            update={
                "data": data,
                "parts": parts,
                "current_registrations_data": serialize_xml(current)
                if current is not None
                else None,
            }
        )


class FontInfo:
    def __init__(self, collection, font, source):
        self._collection = collection
        self._name = font.getAttributeNS(W, "name")
        self._last = font.cloneNode(True)
        self._overrides = {}

    name = property(lambda self: self._name)

    @property
    def _entry(self):
        font = self._collection._lookup(self._name)
        if font is not None:
            self._last = font.cloneNode(True)
        return self._last

    def _value(self, tag, default=None):
        children = _children(self._entry, tag)
        return children[0].getAttributeNS(W, "val") if children else default

    def _set(self, tag, value):
        entry = self._entry
        self._collection._update(self._name, tag, value, entry)
        for child in _children(self._last, tag):
            self._last.removeChild(child)
        if value is not None:
            child = self._last.ownerDocument.createElementNS(W, "w:" + tag)
            child.setAttributeNS(W, "w:val", value)
            self._last.appendChild(child)

    @property
    def alt_name(self):
        return self._overrides.get("alt_name", self._value("altName", ""))

    @alt_name.setter
    def alt_name(self, value):
        if value is None:
            raise RuntimeError("Alternate name must not be null")
        if not isinstance(value, str):
            raise TypeError("Alternate name requires str")
        self._set("altName", value.split("\0", 1)[0] or None)
        self._overrides["alt_name"] = value

    @property
    def charset(self):
        return _charset(self._value("charset", "00"))

    @charset.setter
    def charset(self, value):
        if not isinstance(value, int) or isinstance(value, bool):
            raise TypeError("Charset requires int")
        if not -(1 << 31) <= value < 1 << 31:
            raise OverflowError("Charset exceeds signed 32-bit range")
        value = 0 if value == -1 else value
        self._set("charset", format(value & 0xFFFFFFFF, "02X"))

    @property
    def family(self):
        names = ["auto", "roman", "swiss", "modern", "script", "decorative"]
        value = self._value("family", "auto")
        return FontFamily(names.index(value) if value in names else 0)

    @family.setter
    def family(self, value):
        if not isinstance(value, FontFamily):
            raise TypeError("Family requires FontFamily")
        self._set(
            "family",
            ["auto", "roman", "swiss", "modern", "script", "decorative"][value],
        )

    @property
    def pitch(self):
        names = ["default", "fixed", "variable"]
        value = self._value("pitch", "default")
        return FontPitch(names.index(value) if value in names else 0)

    @pitch.setter
    def pitch(self, value):
        if not isinstance(value, FontPitch):
            raise TypeError("Pitch requires FontPitch")
        self._set("pitch", ["default", "fixed", "variable"][value])

    @property
    def is_true_type(self):
        return not _not_true_type(self._value("notTrueType", "0"))

    @is_true_type.setter
    def is_true_type(self, value):
        if type(value) is not bool:
            raise TypeError("is_true_type requires bool")
        self._set("notTrueType", None if value else "1")

    @property
    def panose(self):
        value = self._value("panose1")
        if not value:
            return None
        digits = "".join(islice((char for char in value if char in "0123456789abcdefABCDEF"), 20))
        decoded = bytearray.fromhex(digits[:len(digits) // 2 * 2])
        return decoded + bytearray(10 - len(decoded))

    @panose.setter
    def panose(self, value):
        if value is not None and not isinstance(value, (bytes, bytearray)):
            raise TypeError("Panose requires bytes or bytearray")
        if value is not None and len(value) != 10:
            raise RuntimeError("Incorrect PANOSE array length")
        self._set("panose1", None if value is None else bytes(value).hex().upper())


def prepare_font_table_commit(document, saved_bytes):
    """Prepare the public metadata view without modifying its original resource graph."""
    source = document.source_font_table
    data = decode(saved_bytes) if is_flat_opc(saved_bytes) else saved_bytes
    with ZipFile(BytesIO(data)) as package:
        registrations = package.read("word/fontTable.xml")
    _root(registrations)
    projection = (source.model_copy(update={"current_registrations_data": registrations})
                  if source is not None else ldm.SourceFontTable(data=registrations))

    def commit():
        document.source_font_table = projection

    return commit
