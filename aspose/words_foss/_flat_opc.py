"""Bounded Flat OPC XML-package decoding for the existing OPC readers."""

import base64
from io import BytesIO
from urllib.parse import quote, unquote, urlsplit
from xml.dom import Node
from zipfile import ZIP_DEFLATED, ZipFile

from defusedxml.ElementTree import ParseError, iterparse
from defusedxml.minidom import parseString

from aspose.words_foss import _io
from aspose.words_foss._opc import bind_namespace_context, resolve_target
from aspose.words_foss.diagnostics import ContentLossWarning, warn
from aspose.words_foss.utils.xml_helpers import serialize_xml

PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
MAIN_TYPES = {
    24: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    25: "application/vnd.ms-word.document.macroEnabled.main+xml",
    26: "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    27: "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
}


def is_flat_opc(data: bytes) -> bool:
    """Recognize the actual root namespace, including UTF-16 and inherited prefixes."""
    _io.check_input_size(len(data))
    if data.startswith(b"PK"):
        return False
    try:
        return next(iterparse(BytesIO(data), events=("start",), forbid_dtd=True))[1].tag == f"{{{PKG}}}package"
    except (ParseError, StopIteration):
        return False


def decode(data: bytes) -> bytes:
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
            decoded = b"".join(serialize_xml(element if child is elements[0] else child) for child in payload.childNodes)
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
    manifest = serialize_xml(types)
    if len(manifest) > _io.MAX_PART_BYTES or total + len(manifest) > _io.MAX_EXPANDED_BYTES:
        raise ValueError("Flat OPC content types exceed size limits")
    with ZipFile(stream, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", manifest)
        for name, payload in parts.items():
            archive.writestr(name, payload)
    return stream.getvalue()


def encode(data: bytes, main_content_type: str | None = None) -> bytes:
    """Flatten a bounded OPC package; XML serialization is normalized."""
    _io.check_input_size(len(data))
    document = parseString(f'<pkg:package xmlns:pkg="{PKG}"/>')
    root = document.documentElement
    with ZipFile(BytesIO(data)) as archive:
        _io.validate_docx_archive(archive)
        types = parseString(archive.read("[Content_Types].xml"), forbid_dtd=True).documentElement
        if types.namespaceURI != CT or types.localName != "Types":
            raise ValueError("Expected OPC content types")
        defaults, overrides = {}, {}
        for item in types.childNodes:
            if item.nodeType != Node.ELEMENT_NODE:
                continue
            if item.namespaceURI != CT or item.localName not in {"Default", "Override"}:
                raise ValueError("Unexpected OPC content type declaration")
            mapping = defaults if item.localName == "Default" else overrides
            key = (item.getAttribute("Extension").lower() if mapping is defaults
                   else resolve_target("", item.getAttribute("PartName")))
            value = item.getAttribute("ContentType")
            if not key or not value or key in mapping or any(ord(char) < 32 for char in value):
                raise ValueError("Missing or duplicate OPC content type")
            mapping[key] = value
        seen = set()
        for name in archive.namelist():
            if name == "[Content_Types].xml":
                continue
            if name.endswith("/"):
                raise ValueError("OPC directory entries cannot be flattened as parts")
            uri = "/" + quote(name, safe="/")
            if resolve_target("", uri) != name or name.casefold() in seen:
                raise ValueError("Unsafe or duplicate OPC part name")
            seen.add(name.casefold())
            content_type = overrides.get(name, defaults.get(name.rsplit(".", 1)[-1].lower()))
            if name == "word/document.xml" and main_content_type is not None:
                if main_content_type not in MAIN_TYPES.values():
                    raise ValueError("Unsupported Flat OPC main content type")
                content_type = main_content_type
            if not content_type:
                raise ValueError("OPC part has no content type")
            part = document.createElementNS(PKG, "pkg:part")
            part.setAttributeNS(PKG, "pkg:name", uri)
            part.setAttributeNS(PKG, "pkg:contentType", content_type)
            value = archive.read(name)
            if content_type in {"application/xml", "text/xml"} or content_type.endswith("+xml"):
                payload = document.createElementNS(PKG, "pkg:xmlData")
                xml = parseString(value, forbid_dtd=True)
                omitted = False
                # Native Flat OPC rejects processing instructions inside XML elements.
                for element in xml.getElementsByTagName("*"):
                    for child in list(element.childNodes):
                        if child.nodeType == Node.PROCESSING_INSTRUCTION_NODE:
                            element.removeChild(child)
                            omitted = True
                if omitted:
                    warn("Internal XML processing instructions are omitted from Flat OPC output",
                         ContentLossWarning, code="flat_opc.processing_instruction_omitted")
                for child in xml.childNodes:
                    payload.appendChild(document.importNode(child, deep=True))
            else:
                part.setAttributeNS(PKG, "pkg:compression", "store")
                payload = document.createElementNS(PKG, "pkg:binaryData")
                payload.appendChild(document.createTextNode(base64.b64encode(value).decode("ascii")))
            part.appendChild(payload)
            root.appendChild(part)
    output = serialize_xml(document)
    _io.check_input_size(len(output))
    return output
