"""OPC relationship paths shared by original-package editing and conversion."""

import posixpath
from urllib.parse import unquote, urlsplit
from xml.dom import Node as XmlNode

XML = "http://www.w3.org/XML/1998/namespace"
XMLNS = "http://www.w3.org/2000/xmlns/"


def _namespace_bindings(node):
    bindings = {}
    ancestor = node
    while ancestor is not None and ancestor.nodeType == XmlNode.ELEMENT_NODE:
        for i in range(ancestor.attributes.length):
            attribute = ancestor.attributes.item(i)
            if attribute.namespaceURI == XMLNS:
                prefix = "" if attribute.name == "xmlns" else attribute.localName
                bindings.setdefault(prefix, attribute.value)
        ancestor = ancestor.parentNode
    bindings.setdefault("", "")
    bindings[node.prefix or ""] = node.namespaceURI or ""
    for i in range(node.attributes.length):
        attribute = node.attributes.item(i)
        if attribute.prefix and attribute.namespaceURI not in {XML, XMLNS}:
            bindings[attribute.prefix] = attribute.namespaceURI
    return bindings


def bind_namespace_context(source, target, parent=None):
    """Retain in-scope bindings, including prefixes referenced by XML attribute values."""
    destination = _namespace_bindings(parent) if parent is not None else {}
    for prefix, uri in _namespace_bindings(source).items():
        if prefix != "xml" and destination.get(prefix) != uri:
            target.setAttributeNS(XMLNS, f"xmlns:{prefix}" if prefix else "xmlns", uri)


def relationships_path(part_name: str) -> str:
    directory, filename = posixpath.split(part_name)
    return posixpath.join(directory, "_rels", filename + ".rels")


def resolve_target(part_name: str, target: str) -> str:
    """Resolve an internal URI inside the package, never on the filesystem."""
    uri = urlsplit(target)
    if not target or uri.scheme or uri.netloc or uri.query or uri.fragment:
        raise ValueError("Invalid internal relationship target")
    path = unquote(uri.path)
    if "\\" in path or any(ord(char) < 32 for char in path):
        raise ValueError("Unsafe internal relationship target")
    name = posixpath.normpath(path.lstrip("/") if path.startswith("/") else
                             posixpath.join(posixpath.dirname(part_name), path))
    if name in {".", ".."} or name.startswith("../"):
        raise ValueError("Relationship target escapes the package")
    return name


def related_part_snapshot(archive, part_name: str) -> tuple[tuple[str, str, bytes], ...]:
    """Capture reachable OPC parts once, retaining external links without fetching."""
    from xml.etree.ElementTree import ParseError

    from defusedxml import DefusedXmlException
    from defusedxml.ElementTree import fromstring

    def parse_xml(data):
        try:
            return fromstring(data, forbid_dtd=True)
        except (ParseError, DefusedXmlException):
            raise ValueError("Invalid related OPC XML") from None

    types_ns = "{http://schemas.openxmlformats.org/package/2006/content-types}"
    rels_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    types = parse_xml(archive.read("[Content_Types].xml"))
    if types.tag != types_ns + "Types":
        raise ValueError("Expected an OPC content types part")
    overrides, defaults = {}, {}
    for item in types:
        if item.tag == types_ns + "Override":
            key = resolve_target("", item.get("PartName", ""))
            mapping = overrides
        elif item.tag == types_ns + "Default":
            key = item.get("Extension", "").lower()
            mapping = defaults
        else:
            raise ValueError("Invalid OPC content type declaration")
        value = item.get("ContentType", "")
        if not key or not value or key in mapping:
            raise ValueError("Empty or duplicate OPC content type declaration")
        mapping[key] = value
    names = set(archive.namelist())
    pending, visited, result = [part_name], set(), []
    while pending:
        name = pending.pop()
        if name in visited:
            continue
        visited.add(name)
        if name not in names:
            raise ValueError("Missing related OPC part")
        content_type = overrides.get(name, defaults.get(name.rsplit(".", 1)[-1].lower()))
        if content_type is None:
            raise ValueError("Missing related OPC content type")
        result.append((name, content_type, archive.read(name)))
        rels_name = relationships_path(name)
        if rels_name not in names:
            continue
        data = archive.read(rels_name)
        root = parse_xml(data)
        if root.tag != rels_ns + "Relationships":
            raise ValueError("Expected an OPC relationships part")
        result.append((rels_name, "application/vnd.openxmlformats-package.relationships+xml", data))
        ids = set()
        for item in root:
            rid, target = item.get("Id", ""), item.get("Target", "")
            mode = item.get("TargetMode", "")
            if (item.tag != rels_ns + "Relationship" or not rid or rid in ids
                    or not target or not item.get("Type") or mode not in {"", "Internal", "External"}):
                raise ValueError("Invalid related OPC relationship")
            ids.add(rid)
            if mode != "External":
                dependency = resolve_target(name, target)
                if dependency.endswith(".rels") or dependency == "[Content_Types].xml":
                    raise ValueError("Relationship targets an OPC infrastructure part")
                pending.append(dependency)
    return tuple(result)
