"""Cross-document ownership and destination-style mapping for the package DOM."""

from enum import IntEnum

from aspose.words_foss._opc import resolve_target
from aspose.words_foss.dom.nodes import (
    XMLNS,
    Node,
    W,
    _bind_namespace_context,
    _elements,
    _find,
    _is,
    _safe_structure,
)
from aspose.words_foss.dom.styles import StyleResolver


class ImportFormatMode(IntEnum):
    USE_DESTINATION_STYLES = 0
    KEEP_SOURCE_FORMATTING = 1
    KEEP_DIFFERENT_STYLES = 2


def _set_word_attribute(element, name, value):
    element.setAttributeNS(XMLNS, "xmlns:w", W)
    element.setAttributeNS(W, "w:" + name, value)


def _theme_payload(document):
    from aspose.words_foss.dom.resources import relationship_root

    links = [node for node in _elements(relationship_root(document._package, "word/document.xml"))
             if node.getAttribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"]
    if len(links) > 1 or links and links[0].getAttribute("TargetMode") not in {"", "Internal"}:
        raise ValueError("Expected at most one internal theme relationship")
    return document._package.payload(resolve_target("word/document.xml", links[0].getAttribute("Target"))) if links else None


def import_node(destination, node, deep, mode):
    if not isinstance(node, Node):
        raise TypeError("src_node must be a DOM node")
    if not isinstance(deep, bool):
        raise TypeError("is_import_children must be a boolean")
    if not isinstance(mode, ImportFormatMode):
        raise TypeError("import_format_mode must be an ImportFormatMode")
    source = node.owner_document
    if source is not destination and mode != ImportFormatMode.USE_DESTINATION_STYLES:
        raise NotImplementedError("Source-format preservation and different-style merging are not implemented")
    copied = node.clone(deep)
    if source is destination and node.part_name == "word/document.xml":
        return copied
    target_tree = destination._package.tree("word/document.xml")
    element = target_tree.importNode(copied._element, deep=True)
    _bind_namespace_context(copied._element, element)
    if source is destination:
        return destination._wrap("word/document.xml", element)
    source_styles, target_styles = StyleResolver(source), StyleResolver(destination)
    root = target_styles.root.cloneNode(deep=True) if target_styles.root is not None else None
    planned = {}

    def map_style(style_id, kind):
        if not source_styles.has_style(style_id, kind):
            raise ValueError("Missing or incompatible source style")
        if style_id in planned:
            return planned[style_id]
        chain = source_styles._chain(style_id, kind)
        if any(style.getElementsByTagNameNS(W, "numPr") for style in chain):
            raise NotImplementedError("Importing numbered styles requires numbering remapping")
        original = source_styles.styles[style_id]
        name = _find(original, "name")
        if name is None or not name.getAttributeNS(W, "val"):
            raise ValueError("Imported style requires a name")
        matches = [style for style in target_styles.styles.values()
                   if target_styles._kind(style) == kind and _find(style, "name") is not None
                   and _find(style, "name").getAttributeNS(W, "val") == name.getAttributeNS(W, "val")]
        if len(matches) > 1:
            raise ValueError("Ambiguous destination style name")
        if matches:
            result = matches[0].getAttributeNS(W, "styleId")
            planned[style_id] = result
            return result
        if root is None:
            raise NotImplementedError("Creating a missing styles part is not implemented")
        if not _safe_structure(original, properties=True):
            raise NotImplementedError("Importing complex style dependencies requires resource translation")
        source_defaults = _find(source_styles.root, "docDefaults")
        target_defaults = _find(target_styles.root, "docDefaults")
        if (source_defaults.toxml() if source_defaults else None) != (target_defaults.toxml() if target_defaults else None):
            raise NotImplementedError("Importing new styles across different document defaults requires format resolution")
        if _theme_payload(source) != _theme_payload(destination):
            raise NotImplementedError("Importing new styles across different themes requires font/color translation")
        taken = {style.getAttributeNS(W, "styleId") for style in _elements(root) if _is(style, "style")} | set(planned.values())
        result, index = style_id, 0
        while result in taken:
            result = f"{style_id}_{index}"
            index += 1
        planned[style_id] = result
        imported = root.ownerDocument.importNode(original, deep=True)
        _bind_namespace_context(original, imported, root)
        _set_word_attribute(imported, "styleId", result)
        if imported.hasAttributeNS(W, "default"):
            imported.removeAttributeNS(W, "default")
        for relation in ("basedOn", "next", "link"):
            reference = _find(imported, relation)
            if reference is None:
                continue
            dependency = reference.getAttributeNS(W, "val")
            dependency_kind = "character" if relation == "link" and kind == "paragraph" else "paragraph" if relation == "link" else kind
            if relation == "next" and dependency == style_id:
                mapped = result
            else:
                mapped = map_style(dependency, dependency_kind)
            if relation == "basedOn" and mapped in target_styles.styles:
                source_bases = source_styles._chain(dependency, kind)
                target_bases = target_styles._chain(mapped, kind)
                if any(not _safe_structure(base, properties=True) for base in source_bases + target_bases):
                    raise NotImplementedError("Importing complex base-style dependencies requires resource translation")
                # shortcut: compare full ancestry XML until all effective style properties can be translated.
                for properties in ("pPr", "rPr", "tblPr", "tcPr", "trPr", "tblStylePr"):
                    a = [item.toxml() for base in source_bases for item in _elements(base) if _is(item, properties)]
                    b = [item.toxml() for base in target_bases for item in _elements(base) if _is(item, properties)]
                    if a != b:
                        raise NotImplementedError("Importing a new style over a conflicting base requires effective-format translation")
            _set_word_attribute(reference, "val", mapped)
        root.appendChild(imported)
        return result

    for item in [element, *element.getElementsByTagNameNS(W, "*")]:
        if _is(item, "numPr"):
            raise NotImplementedError("Importing list references requires numbering remapping")
        kind = {"pStyle": "paragraph", "rStyle": "character", "tblStyle": "table"}.get(item.localName)
        if item.namespaceURI == W and kind is not None:
            _set_word_attribute(item, "val", map_style(item.getAttributeNS(W, "val"), kind))
    if root is not None and len(_elements(root)) != len(_elements(target_styles.root)):
        part_name = next(name for name, tree in destination._package._trees.items()
                         if tree is target_styles.root.ownerDocument)
        destination._package.set_parts({part_name: root.toxml(encoding="utf-8")})
    return destination._wrap("word/document.xml", element)
