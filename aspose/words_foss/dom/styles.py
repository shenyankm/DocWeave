"""Read-only style resolution for supported XML properties, with explicit context limits."""

from dataclasses import dataclass

from aspose.words_foss.dom.nodes import W, _elements, _find, _is, _onoff, _read_size


def _child(node, name):
    return _find(node, name) if node is not None else None


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
        if paragraph is None or not _is(paragraph._element, "p"):
            raise ValueError("Effective run formatting requires a direct paragraph parent")
        paragraph_styles, _ = self._paragraph_layers(paragraph)
        direct = _child(run._element, "rPr")
        reference = _child(direct, "rStyle")
        style_id = reference.getAttributeNS(W, "val") if reference is not None else self._default("character")
        character_styles = self._chain(style_id, "character")
        bold, italic, size = False, False, None
        default = self._defaults("rPr")
        bold = _onoff(_child(default, "b"))
        italic = _onoff(_child(default, "i"))
        size = _read_size(_child(default, "sz"))
        for style in paragraph_styles + character_styles:
            properties = _child(style, "rPr")
            if _child(properties, "rStyle") is not None:
                raise NotImplementedError("Nested run-style references in style definitions are unsupported")
            # Style-level toggles invert the inherited value; false leaves it unchanged.
            bold ^= _onoff(_child(properties, "b"))
            italic ^= _onoff(_child(properties, "i"))
            element = _child(properties, "sz")
            if element is not None:
                size = _read_size(element)
        element = _child(direct, "b")
        if element is not None:
            bold = _onoff(element)
        element = _child(direct, "i")
        if element is not None:
            italic = _onoff(element)
        element = _child(direct, "sz")
        if element is not None:
            size = _read_size(element)
        return EffectiveFont(bold, italic, size)
