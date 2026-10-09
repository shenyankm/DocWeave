"""Cross-document ownership and destination-style mapping for the package DOM."""

from enum import IntEnum

from aspose.words_foss._opc import resolve_target
from aspose.words_foss.dom.nodes import (
    XMLNS,
    Font,
    Node,
    ParagraphFormat,
    W,
    _bind_namespace_context,
    _elements,
    _find,
    _is,
    _new,
    _onoff,
    _read_size,
    _safe_structure,
)
from aspose.words_foss.dom.styles import (
    StyleResolver,
    _child,
    _style_toggle,
    serialized_style_payload,
)


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


def _format_layers(chain):
    return [[item.toxml() for item in _elements(style)
             if item.localName in {"pPr", "rPr", "tblPr", "tcPr", "trPr", "tblStylePr"}]
            for style in chain]


def _other_format_layers(chain):
    result = []
    for style in chain:
        layers = []
        for original in _elements(style):
            if original.localName not in {"pPr", "rPr", "tblPr", "tcPr", "trPr", "tblStylePr"}:
                continue
            copied = original.cloneNode(deep=True)
            for child in list(_elements(copied)):
                if child.namespaceURI == W and child.localName in ({"b", "i", "sz"} if _is(copied, "rPr") else {"jc"} if _is(copied, "pPr") else set()):
                    copied.removeChild(child)
            attributes = [copied.attributes.item(i) for i in range(copied.attributes.length)
                          if copied.attributes.item(i).namespaceURI != XMLNS]
            content = [child for child in copied.childNodes
                       if child.nodeType != child.TEXT_NODE or child.data.strip()]
            if attributes or content:
                layers.append(copied.toxml())
        result.append(layers)
    return result


def _set_style_property(style, group, name, value):
    properties = _find(style, group)
    if properties is None:
        if value is None:
            return
        properties = _new(style, group)
        anchor = _find(style, "rPr") if group == "pPr" else None
        style.insertBefore(properties, anchor)
    for child in list(_elements(properties)):
        if _is(child, name):
            properties.removeChild(child)
    if value is not None:
        element = _new(properties, name)
        _set_word_attribute(element, "val", value)
        order = Font.order if group == "rPr" else ParagraphFormat.order
        later = set(order[order.index(name) + 1:])
        anchor = next((item for item in _elements(properties) if item.namespaceURI == W and item.localName in later), None)
        properties.insertBefore(element, anchor)


def _nearest_style_property(resolver, chain, kind, group, name):
    element = _child(resolver._defaults(group), name) if kind == "paragraph" else None
    for style in chain:
        candidate = _child(_child(style, group), name)
        if candidate is not None:
            element = candidate
    if name == "sz":
        if element is not None:
            _read_size(element)
            return str(int(element.getAttributeNS(W, "val")))
        return None
    return element.getAttributeNS(W, "val") if element is not None else "left" if kind == "paragraph" else None


def _validate_simple_style_properties(chain):
    for style in chain:
        for group, names in (("rPr", {"b", "i", "sz"}), ("pPr", {"jc"})):
            groups = [node for node in _elements(style) if _is(node, group)]
            if len(groups) > 1:
                raise ValueError("Duplicate style property groups")
            for properties in groups:
                seen = set()
                for element in _elements(properties):
                    if element.namespaceURI != W or element.localName not in names:
                        continue
                    if element.localName in seen:
                        raise ValueError("Duplicate style property")
                    seen.add(element.localName)
                    attributes = [element.attributes.item(i) for i in range(element.attributes.length)]
                    if any(attribute.namespaceURI != XMLNS and (attribute.namespaceURI != W or attribute.localName != "val")
                           for attribute in attributes) or element.childNodes:
                        raise NotImplementedError("Migrating decorated style properties requires metadata preservation")


def _translate_style(imported, source_chain, source_bases, target_bases, source, target, kind):
    if any(not _safe_structure(base, properties=True) for base in source_bases + target_bases):
        raise NotImplementedError("Importing complex base-style dependencies requires resource translation")
    default = source._defaults("rPr")
    if kind not in {"paragraph", "character"}:
        if _format_layers(source_bases) != _format_layers(target_bases):
            raise NotImplementedError("Importing a new style over a conflicting base requires effective-format translation")
        return
    a, b = _other_format_layers(source_bases), _other_format_layers(target_bases)
    if (any(a) or any(b)) and a != b:
        raise NotImplementedError("Importing other conflicting style properties requires effective-format translation")
    if any(style.getElementsByTagNameNS(W, "rStyle") for style in source_chain + target_bases):
        raise NotImplementedError("Nested run-style dependencies require format translation")
    _validate_simple_style_properties(source_chain + target_bases)
    for group, names in (("rPr", ("b", "i", "sz")), ("pPr", ("jc",))):
        if kind == "character" and group == "pPr":
            continue
        for name in names:
            if name in {"b", "i"}:
                value = _style_toggle(source_chain, name)
                base = _style_toggle(target_bases, name)
                if kind == "paragraph":
                    value = _onoff(_child(default, name)) if value is None else value
                    base = _onoff(_child(default, name)) if base is None else base
                elif value is not None:
                    value ^= _onoff(_child(default, name))
                value = None if value is None else "1" if value else "0"
                base = None if base is None else "1" if base else "0"
            else:
                value = _nearest_style_property(source, source_chain, kind, group, name)
                base = _nearest_style_property(target, target_bases, kind, group, name)
                if name == "sz" and kind == "paragraph" and value is None and base is not None:
                    raise NotImplementedError("Migrating an implicit application font size requires default-format calibration")
                if name == "jc" and value != base and any(item not in {"left", "right", "center", "both"} for item in (value, base)):
                    raise NotImplementedError("Migrating complex paragraph alignment requires layout context")
            _set_style_property(imported, group, name, None if kind == "paragraph" and value == base else value)


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
        target_bases = []
        source_bases = source_styles._chain(style_id, kind)[:-1]
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
            if relation == "basedOn":
                target_bases = target_styles._chain(mapped, kind)
            _set_word_attribute(reference, "val", mapped)
        _translate_style(imported, chain, source_bases, target_bases, source_styles, target_styles, kind)
        root.appendChild(imported)
        target_styles.styles[result] = imported
        return result

    for item in [element, *element.getElementsByTagNameNS(W, "*")]:
        if _is(item, "numPr"):
            raise NotImplementedError("Importing list references requires numbering remapping")
        kind = {"pStyle": "paragraph", "rStyle": "character", "tblStyle": "table"}.get(item.localName)
        if item.namespaceURI == W and kind is not None:
            _set_word_attribute(item, "val", map_style(item.getAttributeNS(W, "val"), kind))
    if root is not None:
        part_name = next(name for name, tree in destination._package._trees.items()
                         if tree is target_styles.root.ownerDocument)
        serialized_style_payload(destination._package, part_name, root=root, extra=element)
        if len(_elements(root)) != len(_elements(target_styles.root)):
            tree = destination._package.tree(part_name).cloneNode(deep=True)
            tree.replaceChild(tree.importNode(root, deep=True), tree.documentElement)
            destination._package.set_parts({part_name: tree.toxml(encoding="utf-8")})
        destination._package._style_projection_part = part_name
    return destination._wrap("word/document.xml", element)
