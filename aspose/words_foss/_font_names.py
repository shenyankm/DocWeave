"""Runtime effective-name views over sparse conversion-model declarations."""

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.utils.xml_helpers import resolve_font_names

FIELDS = {'name_ascii': 'ascii', 'name_other': 'hAnsi', 'name_bi': 'cs', 'name_far_east': 'eastAsia'}


class FontNameBinding:
    def __init__(self, document, theme_fonts):
        self.document = document
        self.theme_fonts = dict(theme_fonts)

    def local(self, font):
        if font is None:
            return {}
        if font.source_font_names is not None:
            return resolve_font_names(font.source_font_names.attributes, self.theme_fonts)
        if font.font_names_explicit is False:
            return {}
        raw = font.__dict__
        return {channel: raw[field] or raw['name'] for field, channel in FIELDS.items()
                if raw[field] or raw['name'] and 'name' in font.model_fields_set}

    def chain(self, name):
        document = self.document
        styles = {style.name: style for style in document.styles}
        chain, seen = [], set()
        while name in styles and name not in seen:
            seen.add(name)
            style = styles[name]
            chain.append(style)
            name = style.base_style_name
        return reversed(chain)

    def defaults(self):
        document = self.document
        names = dict.fromkeys(FIELDS.values(), 'Times New Roman')
        names.update(self.local(document.doc_defaults_font))
        return names

    def style(self, name, font):
        names = self.defaults()
        for style in self.chain(name):
            names.update(self.local(style.__dict__['font']))
        names.update(self.local(font))
        return names

    def run(self, paragraph, context, font):
        names = self.defaults()
        table, _, _, conditions = context
        for style in self.chain(table):
            names.update(self.local(style.__dict__['font']))
        for kind in conditions:
            for style in self.chain(table):
                for conditional in style.table_style_properties:
                    if conditional.type == kind:
                        names.update(self.local(conditional.font))
        for name in (paragraph.paragraph_format.style_name or context[1], font.style_name):
            for style in self.chain(name):
                names.update(self.local(style.__dict__['font']))
        names.update(self.local(font))
        return names

    def bind(self):
        document = self.document
        for style in document.styles:
            style._font_name_resolver = lambda font, name=style.name: self.style(name, font)
            if style.__dict__["font"] is None:
                style.__dict__["font"] = ldm.Font(font_names_explicit=False)
            style.__dict__["font"]._name_resolver = style._font_name_resolver
        for paragraph in document.get_child_nodes(ldm.NodeType.PARAGRAPH, True):
            for run in paragraph.runs:
                context = run.font._name_context
                run._font_name_resolver = lambda font, paragraph=paragraph, context=context: self.run(paragraph, context, font)
                run.font._name_resolver = run._font_name_resolver


def bind_font_names(document, theme_fonts):
    binding = FontNameBinding(document, theme_fonts)
    document._font_name_theme_values = dict(theme_fonts)
    binding.bind()
