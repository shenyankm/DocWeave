"""Share native line metrics without changing glyph sizes or cursor flow."""

from contextlib import contextmanager

from fpdf.enums import CharVPos
from fpdf.line_break import Fragment, TextLine

from aspose.words_foss.pdf_writer.constants import DEFAULT_FONT_SIZE_PT


class _BaselineLine(TextLine):
    def get_ordered_fragments(self):
        fragments = super().get_ordered_fragments()
        state = fragments[-1].graphics_state.copy()
        state.font_size_pt = self.baseline_size
        state.char_spacing = 0
        state.underline = state.strikethrough = False
        # An empty metric fragment keeps native bidi order and glyph advances intact.
        return fragments + [Fragment([], state, fragments[-1].k)]


def baseline_size(runs):
    sizes = set()
    scripted = False
    for run in runs:
        if run.text:
            sizes.add(run.font.size if run.font.size > 0 else DEFAULT_FONT_SIZE_PT)
            scripted |= run.font.superscript or run.font.subscript
    return max(sizes) if sizes and (len(sizes) > 1 or scripted) else None


@contextmanager
def baseline_scope(pdf, size):
    previous_vpos = pdf.char_vpos
    nested = hasattr(pdf, "_text_baseline_size")
    if size is None and not nested:
        try:
            yield
        finally:
            pdf.char_vpos = previous_vpos
        return
    previous = getattr(pdf, "_text_baseline_size", None)
    original = pdf._render_styled_text_line
    overridden = "_render_styled_text_line" in pdf.__dict__
    if not nested:
        output_metrics = (0, None, None, None)
        originals = {}
        for name in ("_do_underline", "_do_strikethrough", "link"):
            originals[name] = (getattr(pdf, name), name in pdf.__dict__)
        for name in ("_do_underline", "_do_strikethrough"):
            def decorate(x, y, w, font=None, _draw=originals[name][0]):
                previous_size = pdf.font_size_pt
                if output_metrics[3] is not None:
                    pdf.font_size_pt = output_metrics[3] * pdf.k
                try:
                    return _draw(x, y + output_metrics[0], w, font)
                finally:
                    pdf.font_size_pt = previous_size
            setattr(pdf, name, decorate)

        def link(x, y, w, h, link, alt_text=None, **kwargs):
            delta, metric_size, base_size, glyph_size = output_metrics
            if metric_size is not None and abs(h - metric_size) < 1e-8:
                y += (h - base_size) / 2
                h = base_size
            if glyph_size is not None:
                y += 0.8 * (h - glyph_size)
                h = glyph_size
            return originals["link"][0](x, y + delta, w, h, link, alt_text, **kwargs)
        pdf.link = link
        # shortcut: native TextLine metric integration requires checks on fpdf2 upgrades.
        def draw_line(line, *args, **kwargs):
            nonlocal output_metrics
            previous_metrics = output_metrics
            output_metrics = (0, None, None, None)
            target = pdf._text_baseline_size
            shifted = bool(target and line.fragments and
                           any(f.characters for f in line.fragments) and
                           target > max(f.graphics_state.font_size_pt for f in line.fragments))
            if line.fragments:
                last = line.fragments[-1]
                scripted = any(f.char_vpos != CharVPos.LINE for f in line.fragments if f.characters)
                if scripted:
                    kwargs["prevent_font_change"] = True
                if shifted or scripted:
                    own_size = max(f.font_size for f in line.fragments)
                    metric_size = target / pdf.k if shifted else own_size
                    output_metrics = (0.3 * (metric_size - own_size) - last.lift / pdf.k,
                                      metric_size if shifted else None,
                                      last.font_size, last.font_size_pt / pdf.k)
            if shifted:
                line = _BaselineLine(*line)
                line.baseline_size = target
            try:
                return original(line, *args, **kwargs)
            finally:
                output_metrics = previous_metrics
                if shifted:
                    pdf.current_font_is_set_on_page = False

        pdf._render_styled_text_line = draw_line
    pdf._text_baseline_size = size
    if size is None:
        pdf.char_vpos = CharVPos.LINE
    try:
        yield
    finally:
        pdf.char_vpos = previous_vpos
        if nested:
            pdf._text_baseline_size = previous
        else:
            pdf.__dict__.pop("_text_baseline_size", None)
            for name, (method, instance_override) in originals.items():
                if instance_override:
                    setattr(pdf, name, method)
                else:
                    pdf.__dict__.pop(name, None)
            if overridden:
                pdf._render_styled_text_line = original
            else:
                pdf.__dict__.pop("_render_styled_text_line", None)
