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
