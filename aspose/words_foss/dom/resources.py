"""Insert inline resources without rebuilding unrelated OOXML parts."""

from io import BytesIO
from math import isfinite
from pathlib import Path
from urllib.parse import urlsplit

from defusedxml.minidom import parseString
from PIL import Image

from aspose.words_foss._io import read_bounded, validate_image
from aspose.words_foss._opc import relationships_path, resolve_target
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer.constants import CT_URI, PKG_RELS_URI, REL_HYPERLINK, R_URI, W_URI
from aspose.words_foss.docx_writer.drawing import (
    A_URI, PIC_URI, REL_IMAGE, WP_URI, _render_drawing_run,
)
from aspose.words_foss.dom.nodes import XMLNS, _bind_namespace_context, _elements, _new, _validate_text


def relationship_root(package, part_name):
    name = relationships_path(part_name)
    root = (package.tree(name).documentElement if name in package.part_names else
            parseString(f'<Relationships xmlns="{PKG_RELS_URI}"/>').documentElement)
    if root.namespaceURI != PKG_RELS_URI or root.localName != "Relationships":
        raise ValueError("Expected an OPC relationships part")
    ids = set()
    for node in _elements(root):
        if node.namespaceURI == PKG_RELS_URI and node.localName == "Relationship":
            rid = node.getAttribute("Id")
            if not rid or rid in ids:
                raise ValueError("Empty or duplicate relationship ID")
            ids.add(rid)
    return root


def add_relationship(root, rel_type, target, mode=""):
    relationships = [node for node in _elements(root)
                     if node.namespaceURI == PKG_RELS_URI and node.localName == "Relationship"]
    for node in relationships:
        if (node.getAttribute("Type"), node.getAttribute("Target"), node.getAttribute("TargetMode")) == (rel_type, target, mode):
            return node.getAttribute("Id")
    ids = {node.getAttribute("Id") for node in relationships}
    number = 1
    while f"rId{number}" in ids:
        number += 1
    rid = f"rId{number}"
    tag = f"{root.prefix}:Relationship" if root.prefix else "Relationship"
    node = root.ownerDocument.createElementNS(PKG_RELS_URI, tag)
    for key, value in {"Id": rid, "Type": rel_type, "Target": target}.items():
        node.setAttribute(key, value)
    if mode:
        node.setAttribute("TargetMode", mode)
    root.appendChild(node)
    return rid


def validate_link(target):
    _validate_text(target)
    uri = urlsplit(target)
    if (not target or any(char.isspace() for char in target) or
            not (target.startswith("#") and len(target) > 1 or
                 uri.scheme in {"http", "https"} and uri.netloc or
                 uri.scheme == "mailto" and uri.path)):
        raise ValueError("Hyperlink must be an http(s), mailto or nonempty bookmark URL")


def link_parts(node, target):
    validate_link(target)
    package = node.owner_document._package
    if target.startswith("#"):
        return {}, None
    root = relationship_root(package, node.part_name).cloneNode(deep=True)
    rid = add_relationship(root, REL_HYPERLINK, target, "External")
    return {relationships_path(node.part_name): root.toxml(encoding="utf-8")}, rid


def set_link_target(node, target):
    node._editable()
    parts, rid = link_parts(node, target)
    node.owner_document._package.set_parts(parts)
    element = node._element
    if element.hasAttributeNS(R_URI, "id"):
        element.removeAttributeNS(R_URI, "id")
    if element.hasAttributeNS(W_URI, "anchor"):
        element.removeAttributeNS(W_URI, "anchor")
    if rid is not None:
        element.setAttributeNS(XMLNS, "xmlns:r", R_URI)
        element.setAttributeNS(R_URI, "r:id", rid)
    else:
        element.setAttributeNS(XMLNS, "xmlns:w", W_URI)
        element.setAttributeNS(W_URI, "w:anchor", target[1:])
    node._changed()


def add_hyperlink(paragraph, text, target):
    paragraph._structural_editable()
    _validate_text(text)
    parts, rid = link_parts(paragraph, target)
    element = _new(paragraph._element, "hyperlink")
    if rid is not None:
        element.setAttributeNS(XMLNS, "xmlns:r", R_URI)
        element.setAttributeNS(R_URI, "r:id", rid)
    else:
        element.setAttributeNS(XMLNS, "xmlns:w", W_URI)
        element.setAttributeNS(W_URI, "w:anchor", target[1:])
    node = paragraph.owner_document._wrap(paragraph.part_name, element)
    element.appendChild(paragraph.owner_document.create_run(text, part_name=paragraph.part_name)._element)
    paragraph.owner_document._package.set_parts(parts)
    paragraph._element.appendChild(element)
    paragraph._changed()
    return node


def add_picture(paragraph, source, width, height, alternative_text):
    paragraph._structural_editable()
    _validate_text(alternative_text)
    if isinstance(source, bytes):
        data = read_bounded(BytesIO(source))
    elif hasattr(source, "read"):
        data = read_bounded(source)
    else:
        with Path(source).open("rb") as stream:
            data = read_bounded(stream)
    validate_image(data)
    with Image.open(BytesIO(data)) as image:
        if image.format not in {"PNG", "JPEG"}:
            raise ValueError("DOM picture insertion supports PNG and JPEG only")
        image.verify()
        ext, content_type = ("png", "image/png") if image.format == "PNG" else ("jpeg", "image/jpeg")
        native_width, native_height = image.width * 0.75, image.height * 0.75
    if width is None and height is None:
        width, height = native_width, native_height
    elif width is None:
        width = height * native_width / native_height if isinstance(height, (int, float)) else height
    elif height is None:
        height = width * native_height / native_width if isinstance(width, (int, float)) else width
    if any(isinstance(value, bool) or not isinstance(value, (int, float)) or
           not isfinite(value) or not 0 < value <= (2**63 - 1) / 12700 for value in (width, height)):
        raise ValueError("Picture dimensions must be positive finite point values")
    package = paragraph.owner_document._package
    root = relationship_root(package, paragraph.part_name).cloneNode(deep=True)
    existing = next((name for name in package.part_names
                     if name.startswith("word/media/") and name.endswith("." + ext)
                     and package.payload(name) == data), None)
    number = 1
    names = {name.casefold() for name in package.part_names}
    while f"word/media/image{number}.{ext}".casefold() in names:
        number += 1
    name = existing or f"word/media/image{number}.{ext}"
    from posixpath import dirname, relpath

    rid = add_relationship(root, REL_IMAGE, relpath(name, dirname(paragraph.part_name)))
    types = package.tree("[Content_Types].xml").documentElement.cloneNode(deep=True)
    if types.namespaceURI != CT_URI or types.localName != "Types":
        raise ValueError("Expected an OPC content types part")
    overrides = [item for item in _elements(types) if item.namespaceURI == CT_URI
                 and item.localName == "Override" and item.getAttribute("PartName") == "/" + name]
    if len(overrides) > 1 or overrides and overrides[0].getAttribute("ContentType") != content_type:
        raise ValueError("Conflicting image content type")
    if not overrides:
        tag = f"{types.prefix}:Override" if types.prefix else "Override"
        override = types.ownerDocument.createElementNS(CT_URI, tag)
        override.setAttribute("PartName", "/" + name)
        override.setAttribute("ContentType", content_type)
        types.appendChild(override)
    part_types = {resolve_target("", item.getAttribute("PartName")): item.getAttribute("ContentType")
                  for item in _elements(types) if item.namespaceURI == CT_URI and item.localName == "Override"}
    default_types = {item.getAttribute("Extension"): item.getAttribute("ContentType")
                     for item in _elements(types) if item.namespaceURI == CT_URI and item.localName == "Default"}
    ids = set()
    for part in package.part_names:
        kind = part_types.get(part, default_types.get(part.rsplit(".", 1)[-1], ""))
        if kind.endswith("+xml") or kind in {"application/xml", "text/xml"}:
            for doc_pr in package.tree(part).getElementsByTagNameNS(WP_URI, "docPr"):
                raw = doc_pr.getAttribute("id")
                if not raw.isascii() or not raw.isdecimal() or len(raw) > 10 or not 0 <= int(raw) <= 2**32 - 1:
                    raise ValueError("Invalid drawing ID")
                ids.add(int(raw))
    doc_pr_id = 1
    while doc_pr_id in ids:
        doc_pr_id += 1
    shape = ldm.Shape(has_image=True, is_inline=True, width=width, height=height, name="Picture")
    fragment = _render_drawing_run(shape, rid, doc_pr_id)
    namespaces = {"w": W_URI, "r": R_URI, "wp": WP_URI, "a": A_URI, "pic": PIC_URI}
    declarations = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in namespaces.items())
    tree = parseString(f"<root {declarations}>{fragment}</root>")
    tree.getElementsByTagNameNS(WP_URI, "docPr")[0].setAttribute("descr", alternative_text)
    original = tree.documentElement.firstChild
    element = paragraph._element.ownerDocument.importNode(original, deep=True)
    _bind_namespace_context(original, element, paragraph._element)
    package.set_parts({name: data, relationships_path(paragraph.part_name): root.toxml(encoding="utf-8"),
                       "[Content_Types].xml": types.toxml(encoding="utf-8")})
    paragraph._element.appendChild(element)
    paragraph._changed()
    return paragraph.owner_document._wrap(paragraph.part_name, element)
