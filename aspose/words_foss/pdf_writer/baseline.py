"""Share native line metrics without changing glyph sizes or cursor flow."""

from contextlib import contextmanager

from fpdf.line_break import Fragment, TextLine


class _BaselineLine(TextLine):
    def get_ordered_fragments(self):
        fragments = super().get_ordered_fragments()
        state = fragments[-1].graphics_state.copy()
        state.font_size_pt = self.baseline_size
        state.char_spacing = 0
        state.underline = state.strikethrough = False
        # An empty metric fragment keeps native bidi order and glyph advances intact.
        return fragments + [Fragment([], state, fragments[-1].k)]


def mixed_size(sizes):
    sizes = set(sizes)
    return max(sizes) if len(sizes) > 1 else None


@contextmanager
def baseline_scope(pdf, size):
    nested = hasattr(pdf, "_text_baseline_size")
    if size is None and not nested:
        yield
        return
    previous = getattr(pdf, "_text_baseline_size", None)
    original = pdf._render_styled_text_line
    overridden = "_render_styled_text_line" in pdf.__dict__
    if not nested:
        # shortcut: native TextLine metric integration requires checks on fpdf2 upgrades.
        def draw_line(line, *args, **kwargs):
            target = pdf._text_baseline_size
            shifted = bool(target and line.fragments and
                           any(f.characters for f in line.fragments) and
                           target > max(f.graphics_state.font_size_pt for f in line.fragments))
            if shifted:
                line = _BaselineLine(*line)
                line.baseline_size = target
            try:
                return original(line, *args, **kwargs)
            finally:
                if shifted:
                    pdf.current_font_is_set_on_page = False

        pdf._render_styled_text_line = draw_line
    pdf._text_baseline_size = size
    try:
        yield
    finally:
        if nested:
            pdf._text_baseline_size = previous
        else:
            pdf.__dict__.pop("_text_baseline_size", None)
            if overridden:
                pdf._render_styled_text_line = original
            else:
                pdf.__dict__.pop("_render_styled_text_line", None)
