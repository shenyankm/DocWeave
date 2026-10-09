"""Bounded Flat OPC XML-package decoding for the existing OPC readers."""

import base64
from io import BytesIO
from urllib.parse import unquote, urlsplit
from xml.dom import Node
from zipfile import ZIP_DEFLATED, ZipFile

from defusedxml.ElementTree import ParseError, iterparse
from defusedxml.minidom import parseString

from aspose.words_foss import _io
from aspose.words_foss._opc import bind_namespace_context

PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def is_flat_opc(data):
    """Recognize the actual root namespace, including UTF-16 and inherited prefixes."""
    _io.check_input_size(len(data))
    if data.startswith(b"PK"):
        return False
    try:
        return next(iterparse(BytesIO(data), events=("start",), forbid_dtd=True))[1].tag == f"{{{PKG}}}package"
    except (ParseError, StopIteration):
        return False


def decode(data):
    _io.check_input_size(len(data))
    root = parseString(data, forbid_dtd=True).documentElement
    if root.namespaceURI != PKG or root.localName != "package":
        raise ValueError("Expected a Flat OPC package")
    parts = {}
    names = set()
    types = parseString(f'<Types xmlns="{CT}"/>').documentElement
    total = 0
    for part in root.childNodes:
        if part.nodeType != Node.ELEMENT_NODE:
            continue
        if part.namespaceURI != PKG or part.localName != "part":
            raise ValueError("Unexpected Flat OPC package child")
        uri = part.getAttributeNS(PKG, "name")
        name = uri[1:] if uri.startswith("/") else ""
        parsed = urlsplit(uri)
        decoded_name = unquote(name)
        if (not name or name.startswith("/") or ".." in name.split("/") or "\\" in name
                or parsed.netloc or parsed.query or parsed.fragment or parsed.scheme
                or any(part in {"", ".", ".."} for part in decoded_name.split("/"))
                or any("/" in unquote(segment) for segment in name.split("/"))
                or "\\" in decoded_name or any(ord(char) < 32 for char in decoded_name)
                or decoded_name.casefold() in names or decoded_name.casefold() == "[content_types].xml"):
            raise ValueError("Unsafe or duplicate Flat OPC part name")
        names.add(decoded_name.casefold())
        if len(parts) >= _io.MAX_ZIP_ENTRIES - 1:
            raise ValueError("Too many Flat OPC parts")
        content_type = part.getAttributeNS(PKG, "contentType")
        if not content_type or any(ord(char) < 32 for char in content_type):
            raise ValueError("Flat OPC part requires a content type")
        children = [child for child in part.childNodes if child.nodeType == Node.ELEMENT_NODE]
        if len(children) != 1 or children[0].namespaceURI != PKG:
            raise ValueError("Flat OPC part requires one payload")
        payload = children[0]
        if payload.localName == "binaryData":
            if any(child.nodeType == Node.ELEMENT_NODE for child in payload.childNodes):
                raise ValueError("Nested binary payload")
            text = "".join(child.data for child in payload.childNodes if child.nodeType in (Node.TEXT_NODE, Node.CDATA_SECTION_NODE))
            decoded = base64.b64decode("".join(text.split()), validate=True)
        elif payload.localName == "xmlData":
            elements = [child for child in payload.childNodes if child.nodeType == Node.ELEMENT_NODE]
            if len(elements) != 1:
                raise ValueError("XML payload requires one root element")
            element = elements[0].cloneNode(deep=True)
            if any(child.nodeType in (Node.TEXT_NODE, Node.CDATA_SECTION_NODE) and child.data.strip()
                   for child in payload.childNodes):
                raise ValueError("XML payload has text outside its root")
            bind_namespace_context(elements[0], element)
            decoded = b"".join(element.toxml(encoding="utf-8") if child is elements[0]
                               else child.toxml(encoding="utf-8") for child in payload.childNodes)
        else:
            raise ValueError("Unknown Flat OPC payload")
        total += len(decoded)
        if len(decoded) > _io.MAX_PART_BYTES or total > _io.MAX_EXPANDED_BYTES:
            raise ValueError("Flat OPC payload exceeds size limits")
        parts[decoded_name] = decoded
        override = types.ownerDocument.createElementNS(CT, "Override")
        override.setAttribute("PartName", uri)
        override.setAttribute("ContentType", content_type)
        types.appendChild(override)
    stream = BytesIO()
    manifest = types.toxml(encoding="utf-8")
    if len(manifest) > _io.MAX_PART_BYTES or total + len(manifest) > _io.MAX_EXPANDED_BYTES:
        raise ValueError("Flat OPC content types exceed size limits")
    with ZipFile(stream, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", manifest)
        for name, payload in parts.items():
            archive.writestr(name, payload)
    return stream.getvalue()
