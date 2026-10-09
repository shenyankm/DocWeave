"""Editable style fonts and resolution for supported XML properties."""

from dataclasses import dataclass

from aspose.words_foss import _io
from aspose.words_foss._opc import resolve_target
from aspose.words_foss.dom.nodes import (
    XMLNS,
    Font,
    W,
    _elements,
    _find,
    _is,
    _onoff,
    _read_size,
)


class StyleCollection:
    """Styles in XML order; lookup validates unique IDs in the related part."""

    def __init__(self, document):
        self.document = document

    def __iter__(self):
        return iter(Style(self.document, key) for key in StyleResolver(self.document).styles)

    def __len__(self):
        return len(StyleResolver(self.document).styles)

    def get_by_id(self, identifier):
        return Style(self.document, identifier) if identifier in StyleResolver(self.document).styles else None

    def get_by_name(self, name):
        matches = [style for style in self if style.name == name]
        if len(matches) > 1:
            raise ValueError("Ambiguous style name")
        return matches[0] if matches else None


class Style:
    """Live handle to a style ID, retaining unrelated style XML when editing."""

    def __init__(self, document, identifier):
        self._document = document
        self._identifier = identifier

    @property
    def document(self):
        return self._document

    @property
    def owner_document(self):
        return self._document

    @property
    def style_id(self):
        return self._identifier

    @property
    def _element(self):
        element = StyleResolver(self.owner_document).styles.get(self.style_id)
        if element is None:
            raise ValueError("Style no longer exists")
        return element

    @property
    def name(self):
        element = _child(self._element, "name")
        return element.getAttributeNS(W, "val") if element is not None else self.style_id

    @property
    def type(self):
        return StyleResolver._kind(self._element)

    @property
    def font(self):
        if self.type == "numbering":
            return None
        self._editable()
        return StyleFont(self)

    @property
    def direct_font(self):
        """Direct values; None means unset and assigning None removes a property."""
        self._editable()
        return StyleFont(self, resolved=False)

    def _editable(self):
        if self.type not in {"paragraph", "character"}:
            raise NotImplementedError("Editing fonts of table/list styles requires calibration")
        resolver = StyleResolver(self.owner_document)
        resolver._chain(self.style_id, self.type)

    def _changed(self):
        element = self._element
        properties = _child(element, "rPr")
        if properties is not None:
            anchor = next((node for node in _elements(element) if node.localName in
                           {"tblPr", "trPr", "tcPr", "tblStylePr"} and node.namespaceURI == W), None)
            element.insertBefore(properties, anchor)
        package = self.owner_document._package
        part = next(name for name, tree in package._trees.items() if tree is element.ownerDocument)
        package._dirty.add(part)
        package._style_projection_part = part


class StyleFont(Font):
    """Nearest inherited b/i/sz getters; setters write this style's own layer."""

    def __init__(self, style, resolved=True):
        super().__init__(style)
        self._resolved = resolved

    def _get(self, name):
        if not self._resolved or name not in {"b", "i", "sz"}:
            return super()._get(name)
        resolver = StyleResolver(self._node.owner_document)
        layers = [resolver._defaults("rPr")] + [_child(style, "rPr") for style in
                  resolver._chain(self._node.style_id, self._node.type)]
        return next((element for layer in reversed(layers) if (element := _child(layer, name)) is not None), None)

    def _toggle(self, name):
        value = super()._toggle(name)
        return bool(value) if self._resolved else value

    def _set(self, name, value):
        if name == "rStyle":
            raise NotImplementedError("Nested character references in styles require calibration")
        if value is None and self._resolved:
            raise TypeError("Use direct_font to clear inherited style properties")
        groups = [node for node in _elements(self._node._element) if _is(node, "rPr")]
        if len(groups) > 1 or groups and len([node for node in _elements(groups[0]) if _is(node, name)]) > 1:
            raise ValueError("Duplicate style font properties")
        super()._set(name, value)
        self._node.owner_document._package._style_font_overrides.add((self._node.style_id, name))


def _style_stories(package):
    names = {"word/document.xml"}
    rels = "word/_rels/document.xml.rels"
    if rels in package.part_names:
        for link in package.tree(rels).getElementsByTagNameNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship"):
            if link.getAttribute("Type") in {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
            } and link.getAttribute("TargetMode") != "External":
                name = resolve_target("word/document.xml", link.getAttribute("Target"))
                if name not in package.part_names:
                    raise ValueError("Header/footer relationship points to a missing part")
                root = package.tree(name).documentElement
                if not (_is(root, "hdr") or _is(root, "ftr")):
                    raise ValueError("Expected a Word header/footer part")
                names.add(name)
    return [package.tree(name).documentElement for name in sorted(names)]


def serialized_style_payload(package, part_name, *, root=None, extra=None):
    """Project supported root toggles for referenced character contexts, retaining live XML."""
    tree = package.tree(part_name).cloneNode(deep=True)
    if root is not None:
        tree.replaceChild(tree.importNode(root, deep=True), tree.documentElement)
    root = tree.documentElement
    styles = {node.getAttributeNS(W, "styleId"): node for node in _elements(root) if _is(node, "style")}
    selected = set()
    character_ancestors = set()

    def select(identifier, kind):
        seen = set()
        while identifier:
            if identifier in seen:
                raise ValueError("Cycle in style basedOn chain")
            seen.add(identifier)
            style = styles.get(identifier)
            if style is None or (style.getAttributeNS(W, "type") or "paragraph") != kind:
                raise ValueError("Missing or incompatible style in saved character context")
            base = _child(style, "basedOn")
            if kind == "character" and package._style_font_overrides:
                selected.add(identifier)
                if len(seen) > 1:
                    character_ancestors.add(identifier)
            if base is None:
                selected.add(identifier)
                return
            identifier = base.getAttributeNS(W, "val")

    defaults = [key for key, style in styles.items() if (style.getAttributeNS(W, "type") or "paragraph") == "paragraph"
                and style.getAttributeNS(W, "default") in {"1", "true", "on"}]
    for story in _style_stories(package) + ([extra] if extra is not None else []):
        paragraphs = ([story] if _is(story, "p") else []) + list(story.getElementsByTagNameNS(W, "p"))
        for paragraph in paragraphs:
            references = []
            for reference in paragraph.getElementsByTagNameNS(W, "rStyle"):
                properties = reference.parentNode
                run = properties.parentNode
                if not (_is(properties, "rPr") and _is(run, "r")):
                    raise NotImplementedError("Saving paragraph-mark character styles requires format calibration")
                ancestor = run.parentNode
                while ancestor is not None and not _is(ancestor, "p"):
                    ancestor = ancestor.parentNode
                if ancestor is paragraph:
                    references.append(reference)
            if not references:
                continue
            reference = _child(_child(paragraph, "pPr"), "pStyle")
            if reference is not None:
                select(reference.getAttributeNS(W, "val"), "paragraph")
            elif defaults:
                if len(defaults) != 1:
                    raise ValueError("Multiple default paragraph styles")
                select(defaults[0], "paragraph")
            for reference in references:
                select(reference.getAttributeNS(W, "val"), "character")
    default = _child(_child(_child(root, "docDefaults"), "rPrDefault"), "rPr")
    changed = False
    for key in selected:
        if (package._style_font_overrides and styles[key].getAttributeNS(W, "type") == "character"
                and _child(styles[key], "basedOn") is None and key not in character_ancestors):
            continue
        groups = [node for node in _elements(styles[key]) if _is(node, "rPr")]
        if len(groups) > 1:
            raise ValueError("Duplicate root style property groups")
        properties = groups[0] if groups else None
        for name in ("b", "i"):
            if (key, name) in package._style_font_overrides:
                continue
            elements = [node for node in _elements(properties) if _is(node, name)] if properties is not None else []
            if len(elements) > 1:
                raise ValueError("Duplicate root style toggle")
            element = elements[0] if elements else None
            default_element = _child(default, name)
            base = _child(styles[key], "basedOn")
            while base is not None:
                ancestor = styles[base.getAttributeNS(W, "val")]
                inherited = _child(_child(ancestor, "rPr"), name)
                if inherited is not None:
                    default_element = inherited
                    break
                base = _child(ancestor, "basedOn")
            if element is None or _onoff(element) != _onoff(default_element):
                continue
            for candidate in (element, default_element):
                if candidate is None:
                    continue
                attributes = [candidate.attributes.item(index) for index in range(candidate.attributes.length)]
                if candidate.childNodes or any(attribute.namespaceURI != XMLNS and
                        (attribute.namespaceURI != W or attribute.localName != "val") for attribute in attributes):
                    raise NotImplementedError("Normalizing decorated root style toggles requires metadata preservation")
            properties.removeChild(element)
            changed = True
    if not changed:
        return None
    data = tree.toxml(encoding="utf-8")
    if len(data) > _io.MAX_PART_BYTES:
        raise ValueError("Projected styles part exceeds the safety limit")
    return data


def _child(node, name):
    return _find(node, name) if node is not None else None


def _style_toggle(styles, name):
    value = None
    for style in styles:
        element = _child(_child(style, "rPr"), name)
        if element is not None:
            value = _onoff(element)
    return value


def _effective_toggle(default, paragraph_styles, character_styles, direct, name):
    # 26.9 getters combine style categories, rather than toggling every basedOn ancestor.
    paragraph = _style_toggle(paragraph_styles, name)
    character = _style_toggle(character_styles, name)
    value = _onoff(_child(default, name))
    if paragraph is not None and character is not None:
        value = value or (paragraph ^ character)
    elif paragraph is not None or character is not None:
        value = paragraph if paragraph is not None else character
    element = _child(direct, name)
    return _onoff(element) if element is not None else value


@dataclass(frozen=True)
class EffectiveFont:
    """Resolved w:b/w:i/w:sz (not complex-script selection or rendered glyph metrics)."""

    bold: bool
    italic: bool
    size: float | None


@dataclass(frozen=True)
class EffectiveParagraphFormat:
    alignment: str | None


class StyleResolver:
    def __init__(self, document):
        self.root = document._styles_root()
        self.styles = {}
        if self.root is not None:
            for element in _elements(self.root):
                if not _is(element, "style"):
                    continue
                style_id = element.getAttributeNS(W, "styleId")
                if not style_id or style_id in self.styles:
                    raise ValueError("Styles require unique, nonempty IDs")
                self.styles[style_id] = element

    def has_style(self, style_id, style_type):
        element = self.styles.get(style_id)
        return element is not None and self._kind(element) == style_type

    @staticmethod
    def _kind(element):
        return element.getAttributeNS(W, "type") or "paragraph"

    def _default(self, style_type):
        defaults = []
        for element in self.styles.values():
            if self._kind(element) != style_type:
                continue
            value = element.getAttributeNS(W, "default")
            if value not in {"", "0", "1", "true", "false", "on", "off"}:
                raise ValueError("Invalid default-style flag")
            if value in {"1", "true", "on"}:
                defaults.append(element.getAttributeNS(W, "styleId"))
        if len(defaults) > 1:
            raise ValueError("Multiple default styles of the same type")
        return defaults[0] if defaults else None

    def _chain(self, style_id, style_type):
        chain, visited = [], set()
        while style_id is not None:
            if style_id in visited:
                raise ValueError("Cycle in style basedOn chain")
            if not self.has_style(style_id, style_type):
                raise ValueError("Missing style or incompatible basedOn style type: " + style_id)
            visited.add(style_id)
            element = self.styles[style_id]
            chain.append(element)
            base = _child(element, "basedOn")
            style_id = base.getAttributeNS(W, "val") if base is not None else None
        return list(reversed(chain))

    def _defaults(self, property_name):
        defaults = _child(self.root, "docDefaults")
        return _child(_child(defaults, property_name + "Default"), property_name)

    def _paragraph_layers(self, paragraph):
        paragraph._editable()
        properties = _child(paragraph._element, "pPr")
        reference = _child(properties, "pStyle")
        style_id = reference.getAttributeNS(W, "val") if reference is not None else self._default("paragraph")
        styles = self._chain(style_id, "paragraph")
        layers = [self._defaults("pPr")] + [_child(style, "pPr") for style in styles] + [properties]
        # ponytail: numbering and conditional table formatting need their own resolvers before claiming effective values.
        if any(_child(layer, "numPr") is not None for layer in layers):
            raise NotImplementedError("Effective formatting for numbered paragraphs is unsupported")
        ancestor = paragraph.parent_node
        while ancestor is not None:
            if _is(ancestor._element, "tbl"):
                table_style = _child(_child(ancestor._element, "tblPr"), "tblStyle")
                if table_style is not None or self._default("table") is not None:
                    raise NotImplementedError("Effective formatting for styled tables is unsupported")
            ancestor = ancestor.parent_node
        return styles, layers

    def paragraph_format(self, paragraph):
        _, layers = self._paragraph_layers(paragraph)
        alignment = None
        for layer in layers:
            element = _child(layer, "jc")
            if element is not None:
                alignment = element.getAttributeNS(W, "val")
                if alignment not in {"left", "right", "center", "both", "distribute", "start", "end",
                                     "numTab", "highKashida", "mediumKashida", "lowKashida", "thaiDistribute"}:
                    raise ValueError("Invalid OOXML paragraph alignment")
        return EffectiveParagraphFormat(alignment)

    def font(self, run):
        run._editable()
        paragraph = run.parent_node
        if paragraph is not None and _is(paragraph._element, "hyperlink"):
            paragraph = paragraph.parent_node
        if paragraph is None or not _is(paragraph._element, "p"):
            raise ValueError("Effective run formatting requires a direct paragraph parent")
        paragraph_styles, _ = self._paragraph_layers(paragraph)
        direct = _child(run._element, "rPr")
        reference = _child(direct, "rStyle")
        style_id = reference.getAttributeNS(W, "val") if reference is not None else None
        character_styles = self._chain(style_id, "character")
        default = self._defaults("rPr")
        size = _read_size(_child(default, "sz"))
        for style in paragraph_styles + character_styles:
            properties = _child(style, "rPr")
            if _child(properties, "rStyle") is not None:
                raise NotImplementedError("Nested run-style references in style definitions are unsupported")
            element = _child(properties, "sz")
            if element is not None:
                size = _read_size(element)
        bold = _effective_toggle(default, paragraph_styles, character_styles, direct, "b")
        italic = _effective_toggle(default, paragraph_styles, character_styles, direct, "i")
        element = _child(direct, "sz")
        if element is not None:
            size = _read_size(element)
        return EffectiveFont(bold, italic, size)
