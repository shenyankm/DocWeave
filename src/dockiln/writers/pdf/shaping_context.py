"""Retain source positions while fpdf2 clones and wraps shaped fragments."""
from types import SimpleNamespace

from fpdf.fonts import TTFFont
from fpdf.line_break import Fragment


class SourceCharacter(str):
    __slots__ = ('source_index',)

    def __new__(cls, character, index):
        value = super().__new__(cls, character)
        value.source_index = index
        return value

    def __getnewargs__(self):
        return str(self), self.source_index


class ContextualFont(TTFFont):
    # shortcut: fpdf2 font/fragment internals retain context; recheck on upgrades.
    __slots__ = ()

    def __deepcopy__(self, memo):
        font = super().__deepcopy__(memo)
        font.__class__ = type(self)
        return font

    def shaped_text_width(self, text, font_size_pt, text_shaping_params):
        if not (text_shaping_params or {}).get('_source_context'):
            return super().shaped_text_width(text, font_size_pt, text_shaping_params)
        _, positions = self.perform_harfbuzz_shaping(text, font_size_pt, text_shaping_params)
        if not positions:
            return 0, 0
        # Native rendering uses exact advances; integer rounding accumulates across styled runs.
        return len(positions), sum(pos.x_advance for pos in positions) * self.scale * font_size_pt / 1000

    def perform_harfbuzz_shaping(self, text, font_size_pt, text_shaping_params):
        context = (text_shaping_params or {}).get('_source_context')
        if not context:
            return super().perform_harfbuzz_shaping(text, font_size_pt, text_shaping_params)
        import uharfbuzz as hb

        source, start, end, lower, upper = context
        source = self._map_symbol_text(source[lower:upper])
        start -= lower
        self.hbfont.ptem = font_size_pt
        buffer = hb.Buffer()
        buffer.cluster_level = 1
        buffer.add_str(source, start, end - lower - start)
        buffer.guess_segment_properties()
        for key in ('fragment_direction', 'script', 'language'):
            value = text_shaping_params.get(key)
            if value is not None:
                setattr(buffer, 'direction' if key == 'fragment_direction' else key,
                        value.value if key == 'fragment_direction' else value)
        hb.shape(self.hbfont, buffer, text_shaping_params['features'])
        infos = [SimpleNamespace(codepoint=info.codepoint, cluster=info.cluster - start)
                 for info in buffer.glyph_infos]
        return infos, buffer.glyph_positions


class ContextualFragment(Fragment):
    def __init__(self, characters, graphics_state, k, link=None):
        super().__init__(characters, graphics_state.copy(), k, link)

    @property
    def text_shaping_parameters(self):
        return self.graphics_state.text_shaping | {
            '_source_context': getattr(self, '_source_context', None)}

    def _set_context(self, characters):
        params = self.text_shaping_parameters
        self._source_context = None
        if not characters or not all(isinstance(char, SourceCharacter) for char in characters):
            return
        start, end = characters[0].source_index, characters[-1].source_index + 1
        source = params['_source_text']
        if source[start:end] != ''.join(characters):
            return
        lower, upper = params.get('_source_line', (0, len(source)))
        self._source_context = source, start, end, lower, upper

    def get_width(self, start=0, end=None, chars=None, initial_cs=True):
        characters = self.characters[start:end] if chars is None else chars
        self._set_context(characters)
        return super().get_width(start, end, chars, initial_cs)

    def render_with_text_shaping(self, *args, **kwargs):
        self._set_context(self.characters)
        return super().render_with_text_shaping(*args, **kwargs)
