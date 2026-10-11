"""Live ordinary size getters over sparse conversion-model declarations."""
from defusedxml.ElementTree import fromstring
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.utils.table_conditions import (
    table_font_size, table_size_conditions, table_run_size, table_size_override,
    table_size_paint_conditions, table_font_paint_size,
)


def local_size(font):
    if font is None:
        return None
    raw = font.__dict__
    if raw['size_explicit'] is True or (raw['size_explicit'] is None and
            'size' in font.model_fields_set and raw['size'] != 0):
        return raw['size']
    return None


class FontSizeBinding:
    def __init__(self, document):
        self.document = document

    def chain(self, name):
        styles = {style.name: style for style in self.document.styles}
        chain, seen = [], set()
        while name in styles and name not in seen:
            seen.add(name)
            style = styles[name]
            chain.append(style)
            name = style.base_style_name
        return list(reversed(chain))

    def default(self):
        value = local_size(self.document.doc_defaults_font)
        if value is not None:
            return value
        return 10.0 if self.document.doc_defaults_rpr_present else 11.0

    def paragraph_style(self):
        enabled = [s.name for s in self.document.styles if s.type == 1 and s.is_default]
        return enabled[-1] if enabled else next((s.name for s in reversed(self.document.styles)
            if s.type == 1 and s.built_in and s.style_identifier == 0), '')

    def table(self, paragraph):
        for table in self.document.get_child_nodes(ldm.NodeType.TABLE, True):
            target = next((c for row in table.rows for c in row.cells
                           if any(p is paragraph for p in c.paragraphs)), None)
            if target is None:
                continue
            name = table.style_name or next((s.name for s in reversed(self.document.styles)
                                             if s.type == 3 and s.is_default), '')
            styles = self.chain(name)
            look = ldm.TableStyleLook()
            bands = {'tblStyleRowBandSize': 1, 'tblStyleColBandSize': 1}
            declared_bands = set()
            for style in styles:
                props = style.table_style_format
                if props is not None:
                    look = props.style_look or look
                    for field, key in (('row_band_size', 'tblStyleRowBandSize'),
                                       ('column_band_size', 'tblStyleColBandSize')):
                        value = getattr(props, field)
                        if value is not None:
                            bands[key] = value
                            declared_bands.add(key)
            look = table.style_look or look
            positions = []
            for row in table.rows:
                column, cells = 0, []
                for cell in row.cells:
                    cells.append((cell, column, cell.cell_format.grid_span))
                    column += cell.cell_format.grid_span
                positions.append(cells)
            conditions = table_size_conditions(positions, look, bands,
                {region.type for style in styles for region in style.table_style_properties})
            layers = [(local_size(style.font), [(region.type, local_size(region.font))
                       for region in style.table_style_properties]) for style in styles]
            columns = max((sum(c.cell_format.grid_span for c in row.cells) for row in table.rows), default=0)
            paint = table_size_paint_conditions(positions, columns, look, bands, declared_bands)
            return layers, conditions.get(id(target), ()), paint.get(id(target), ())
        return [], (), ()

    def run(self, font, paragraph, *, mark=False, paint=False):
        if not mark and not any(run.font is font for run in paragraph.runs):
            paragraph = next((p for p in self.document.get_child_nodes(ldm.NodeType.PARAGRAPH, True)
                              if any(run.font is font for run in p.runs)), paragraph)
        layers, conditions, paint_conditions = ([], (), ()) if mark else self.table(paragraph)
        paragraph_styles = self.chain(paragraph.paragraph_format.style_name or self.paragraph_style())
        paragraph_sizes = [(s.built_in and s.style_identifier == 0, local_size(s.font))
                           for s in paragraph_styles]
        character = None
        for style in ([] if mark else self.chain(font.style_name)):
            value = local_size(style.font)
            if value is not None:
                character = value
        if (paint and not layers and local_size(font) is None and character is None
                and not any(value is not None for _, value in paragraph_sizes)
                and local_size(self.document.doc_defaults_font) is None
                and not self.document.doc_defaults_rpr_present):
            # Legacy unmarked programmatic fonts retain the caller's unset fallback.
            return font.__dict__["size"]
        settings = self.document.current_compatibility_settings or self.document.source_settings
        root = fromstring(settings.data) if settings is not None and settings.data is not None else None
        override = table_size_override(self.document.compatibility_mode, root, paint=paint)
        state = self.document.compatibility_options_state
        if state is not None and state.font_size_override is not None:
            override = (state.optimized_for != 3 if paint else state.font_size_override)
        size = table_run_size(self.default(), table_font_size(layers, ()),
            table_font_paint_size(layers, paint_conditions) if paint else table_font_size(layers, conditions),
            paragraph_sizes, character, local_size(font), table_style=bool(layers), paint=paint,
            override_normal=override)
        return 0.5 if paint and size == 0 else size

    def bind(self):
        for paragraph in self.document.get_child_nodes(ldm.NodeType.PARAGRAPH, True):
            for run in paragraph.runs:
                run.font._size_resolver = lambda font, p=paragraph, paint=False: self.run(font, p, paint=paint)
            mark = paragraph.paragraph_format.paragraph_break_font
            if mark is not None:
                mark._size_resolver = lambda font, p=paragraph, paint=False: self.run(font, p, mark=True, paint=paint)


def bind_font_sizes(document):
    FontSizeBinding(document).bind()


def prepare_saved_compatibility_commit(document, saved_bytes):
    """Bind the actual saved compatibility view without replacing its load snapshot."""
    from io import BytesIO
    from zipfile import ZipFile
    from aspose.words_foss._flat_opc import decode, is_flat_opc
    from xml.etree import ElementTree as ET
    data = decode(saved_bytes) if is_flat_opc(saved_bytes) else saved_bytes
    with ZipFile(BytesIO(data)) as archive:
        settings = fromstring(archive.read("word/settings.xml"), forbid_dtd=True)
    from aspose.words_foss.docx_writer.settings_part import W_URI
    compat = settings.find("{" + W_URI + "}compat")
    projection = ldm.SourceSettings(
        data=ET.tostring(compat) if compat is not None else None,
        compatibility_mode=document.compatibility_mode,
        do_not_expand_shift_return=document.do_not_expand_shift_return)

    def commit():
        document.current_compatibility_settings = projection
        if document.compatibility_options_state is not None:
            document.compatibility_options_state.font_size_override = None

    return commit
