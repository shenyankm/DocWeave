"""Formatted-run rendering for the PDF writer.

Handles left-aligned writes, aligned cell rows, highlight rectangles,
and strikethrough lines.
"""


import re
from itertools import groupby
from typing import Optional, Tuple, Union

from fpdf import FPDF
# ponytail: fpdf2 bidi/line-break internals require regression checks when upgrading fpdf2.
from fpdf.bidi import BidiParagraph
from fpdf.enums import Align, CharVPos, StrokeCapStyle
from fpdf.line_break import Fragment, MultiLineBreak, TextLine

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.baseline import baseline_scope, baseline_size
from aspose.words_foss.pdf_writer.color import parse_color
from aspose.words_foss.pdf_writer.constants import (
    DEFAULT_FONT_SIZE_PT,
    GAP_MIN_SPACES,
    HIGHLIGHT_HEIGHT_RATIO,
    HIGHLIGHT_Y_OFFSET_RATIO,
    HYPERLINK_TEXT_RGB,
    LINE_HEIGHT_FACTOR,
    PT_TO_MM,
)
from aspose.words_foss.pdf_writer.font import apply_run_font
from aspose.words_foss.pdf_writer.shaping_context import ContextualFont, ContextualFragment, SourceCharacter
from aspose.words_foss.pdf_writer.text import (
    apply_caps,
    extract_link_segments,
    safe_text,
)
from aspose.words_foss.docx_reader import PAGE_FIELD_SENTINEL
from aspose.words_foss.pdf_writer._context import PDFWriterContext


class _ShapedSegment(tuple):
    """Keep the existing five values plus the native direction/font context."""

    def __new__(cls, values, fragment):
        segment = super().__new__(cls, values)
        segment.fragment = fragment
        return segment


class RunRenderer:
    """Renders formatted runs (text segments with fonts, colors, links)."""

    def __init__(self, writer: PDFWriterContext) -> None:
        self._writer = writer

    # ------------------------------------------------------------------
    # Public rendering methods
    # ------------------------------------------------------------------

    def render_formatted_runs(
        self,
        pdf: FPDF,
        runs: list[ldm.Run],
        *,
        align: str = "L",
        newline: bool = True,
        line_h_override: Optional[float] = None,
        is_toc: bool = False,
        tab_stops: Optional[ldm.TabStopCollection] = None,
        default_tab_stop: float = 36.0,
        pf: Optional[ldm.ParagraphFormat] = None,
    ) -> None:
        """Render runs with inline formatting: bold, italic, underline, colors, sizes."""
        with baseline_scope(pdf, baseline_size(runs)):
            fs = DEFAULT_FONT_SIZE_PT

            # A ``\t`` between runs marks a Word tab stop — TOC entries put
            # the title before the tab and the page number after.  Route
            # TOC entries through the aligned path so the column-split
            # detector can anchor the trailing runs to the right margin
            # (Word's default TOC right-tab stop).  Ordinary numbered
            # paragraphs like ``4.1<tab>Heading`` must stay on the
            # left-aligned path — the aligned renderer would otherwise
            # split them at the tab and collide the trailing text with
            # the number.
            has_tab = any("\t" in (r.text or "") for r in runs)
            # For non-left alignment, use multi_cell which supports align
            if align != "L" or (has_tab and is_toc) or pdf.text_shaping or (
                    pf is not None and pf.widow_control and not has_tab and newline):
                self.render_formatted_runs_aligned(
                    pdf, runs, align=align, line_h_override=line_h_override, is_toc=is_toc, pf=pf
                )
                return

            for run in runs:
                text = run.text or ""
                if not text:
                    continue
                font = run.font
                size = font.size if font.size > 0 else fs
                line_h = (
                    line_h_override
                    if line_h_override is not None
                    else size * PT_TO_MM * LINE_HEIGHT_FACTOR
                )

                # Strip form-feed characters ("\f" = Word page-break marker).
                text = text.replace("\f", "")
                for piece_index, piece in enumerate(text.split("\t")):
                    if piece_index:
                        apply_run_font(pdf, font, default_size=fs)
                        self._advance_to_tab_stop(
                            pdf, tab_stops, default_tab_stop, line_h, upcoming_text=piece,
                        )
                    text = self._resolve_run_text(pdf, piece)
                    if not text:
                        continue
                    text = apply_caps(text, font)
                    for chunk, link in extract_link_segments(text):
                        if not chunk:
                            continue
                        apply_run_font(pdf, font, default_size=fs)
                        highlight = parse_color(font.highlight_color)
                        if link is not None:
                            # Typical hyperlink convention: blue text.
                            pdf.set_text_color(*HYPERLINK_TEXT_RGB)
                        safe = safe_text(chunk)
                        link_target = self._writer._link_target_for(pdf, link)
                        if highlight:
                            # Use cell-based rendering so the fill rectangle
                            # and text are always perfectly aligned — even
                            # when the run wraps across lines.
                            self._write_with_highlight(pdf, safe, line_h, highlight, run=run, link=link_target)
                        else:
                            pdf.write(h=line_h, text=safe, link=link_target)

            # Reset color
            pdf.set_text_color(0, 0, 0)
            if newline:
                pdf.ln()

    def _advance_to_tab_stop(
        self,
        pdf: FPDF,
        tab_stops: Optional[ldm.TabStopCollection],
        default_tab_stop: float,
        line_h: float,
        upcoming_text: str = "",
    ) -> None:
        """Advance cursor to the next tab stop position, drawing leader fill.

        Handles CENTER/RIGHT/DECIMAL alignment by offsetting the target
        position based on the width of *upcoming_text*.
        """
        w = self._writer
        current_x = pdf.get_x()
        margin_left = w._page_margin_left
        # Native write/cell starts glyphs one inner margin after the cursor.
        pos_pt = (current_x + pdf.c_margin - margin_left) / PT_TO_MM + 1e-7

        next_tab = tab_stops.after(pos_pt) if tab_stops else None
        if next_tab is not None:
            target_pt = next_tab.position
        else:
            step = default_tab_stop if default_tab_stop > 0 else 36.0
            target_pt = ((pos_pt // step) + 1) * step

        if next_tab and upcoming_text:
            alignment = next_tab.alignment
            text_w_mm = pdf.get_string_width(safe_text(upcoming_text))
            text_w_pt = text_w_mm / PT_TO_MM
            if alignment == 1:  # CENTER
                target_pt -= text_w_pt / 2
            elif alignment == 2:  # RIGHT
                target_pt -= text_w_pt
            elif alignment == 3:  # DECIMAL
                dot_idx = -1
                for ci, ch in enumerate(upcoming_text):
                    if ch in ".,":
                        dot_idx = ci
                        break
                if dot_idx >= 0:
                    before_w = pdf.get_string_width(safe_text(upcoming_text[:dot_idx])) / PT_TO_MM
                else:
                    before_w = text_w_pt
                target_pt -= before_w

        target_x = margin_left + target_pt * PT_TO_MM - pdf.c_margin
        gap = target_x - current_x
        if gap <= 0:
            return

        if next_tab and next_tab.leader > 0:
            _LEADER_CHARS = {1: ".", 2: "-", 3: "_", 4: "_", 5: "·"}
            ch = _LEADER_CHARS.get(next_tab.leader, ".")
            char_w = pdf.get_string_width(ch)
            if char_w > 0:
                count = int(gap / char_w)
                if count > 0:
                    pdf.write(h=line_h, text=ch * count)

        pdf.set_x(target_x)

    def render_formatted_runs_aligned(
        self,
        pdf: FPDF,
        runs: list[ldm.Run],
        *,
        align: str = "C",
        line_h_override: Optional[float] = None,
        is_toc: bool = False,
        pf: Optional[ldm.ParagraphFormat] = None,
    ) -> None:
        """Render runs with per-run formatting inside an aligned line.

        Computes the total width of all runs, then positions the cursor
        according to *align* before emitting each run with its own font,
        color and style via ``pdf.cell()``.
        """
        fs = DEFAULT_FONT_SIZE_PT
        usable_w = pdf.epw

        # Pre-compute chunk widths using each run's own font settings.
        segments: list[Tuple[ldm.Run, str, float, float, Optional[str]]] = []
        for run in runs:
            text = self._resolve_run_text(pdf, (run.text or "").replace("\f", ""))
            if not text:
                continue
            font = run.font
            size = font.size if font.size > 0 else fs
            apply_run_font(pdf, font, default_size=fs)
            if not is_toc and "\t" in text:
                text = text.replace("\t", " ")
            text = apply_caps(text, font)
            # Split runs on ``\t`` so each tab becomes its own segment
            pieces = text.split("\t")
            for idx, piece in enumerate(pieces):
                if idx > 0:
                    segments.append((run, "\t", 0.0, size, None))
                if not piece:
                    continue
                for chunk, link in extract_link_segments(piece):
                    if not chunk:
                        continue
                    safe_chunk = safe_text(chunk)
                    chunk_w = pdf.get_string_width(safe_chunk)
                    segments.append((run, safe_chunk, chunk_w, size, link))

        if not segments:
            return

        max_size = max(s for _, _, _, s, _ in segments)
        line_h = (
            line_h_override
            if line_h_override is not None
            else max_size * PT_TO_MM * LINE_HEIGHT_FACTOR
        )

        # Reserve a page-number area; wrapped TOC titles finish beside it.
        tab_idx = (
            next((i for i, (_, t, _, _, _) in enumerate(segments) if t == "\t"), None)
            if is_toc
            else None
        )
        trailing = []
        field_tab = None
        trailing_right = usable_w - pdf.c_margin
        text_width = usable_w - 2 * pdf.c_margin
        if tab_idx is not None:
            if pf is not None:
                # shortcut: one LEFT/CENTER/RIGHT/DECIMAL field; extend for full multi-tab layout.
                field_tab = min((tab for tab in pf.tab_stops
                                 if tab.alignment in (0, 1, 2, 3) and not tab.is_clear
                                 and tab.position > pf.left_indent),
                                key=lambda tab: tab.position, default=None)
                if field_tab is not None:
                    trailing_right = (field_tab.position - pf.left_indent) * PT_TO_MM
                    if not pdf.c_margin < trailing_right <= usable_w + pf.right_indent * PT_TO_MM + 1e-7:
                        raise ValueError("TOC tab is outside the usable text area")
                    text_width = trailing_right - pdf.c_margin
            visible = [s for s in segments[:tab_idx] if s[1] != "\t"]
            right = [s for s in segments[tab_idx + 1 :] if s[1] != "\t"]
            if right:
                field_width = text_width
                if field_tab is not None and field_tab.alignment in (0, 1, 3):
                    field_width = usable_w + pf.right_indent * PT_TO_MM - pdf.c_margin
                trailing_rows = self.wrap_segments(pdf, right, field_width)
                if len(trailing_rows) != 1:
                    raise ValueError("TOC page number is wider than the usable text area")
                trailing = trailing_rows[0]
                right_w = sum(s[2] for s in trailing)
                if field_tab is not None and field_tab.alignment in (0, 1, 3):
                    before = 0.0
                    if field_tab.alignment == 1:
                        before = right_w / 2
                    elif field_tab.alignment == 3:
                        before = self._decimal_prefix_width(pdf, trailing)
                    trailing_right += right_w - before
                    if (trailing_right > usable_w + pf.right_indent * PT_TO_MM + 1e-7 or
                            trailing_right - right_w < pdf.c_margin):
                        raise ValueError("TOC field is outside the usable text area")
                    text_width = trailing_right - pdf.c_margin
                if visible:
                    text_width -= right_w + pdf.c_margin
            align = "L"
        else:
            visible = [s for s in segments if s[1] != "\t"]
        if not visible and not trailing:
            return

        # Word footers frequently encode a two-column "left / right" line
        # as a single right-aligned paragraph with a long run of spaces
        # between the two halves.  Detect the padding idiom and render
        # the halves at the margins instead.
        if align == "R":
            split_idx = self._find_gap_split(visible, usable_w)
            if split_idx is not None:
                left = visible[:split_idx]
                right = visible[split_idx + 1 :]
                self._render_segment_row(pdf, left, line_h, fs, at_x=pdf.l_margin)
                right_w = sum(sw for _, _, sw, _, _ in right)
                right_x = max(
                    pdf.l_margin,
                    pdf.l_margin + usable_w - right_w - 2 * pdf.c_margin,
                )
                pdf.set_y(pdf.get_y() - line_h)  # stay on the same baseline
                self._render_segment_row(pdf, right, line_h, fs, at_x=right_x)
                pdf.set_text_color(0, 0, 0)
                return
        first_offset = pdf.get_x() - pdf.l_margin
        rows = (self.wrap_segments(pdf, visible, text_width,
                                   first_width=text_width - first_offset)
                if visible else [[]])
        for index, row in enumerate(rows):
            if (pf is not None and pf.widow_control and len(rows) > 1 and
                    index in (0, len(rows) - 2) and pdf.auto_page_break and
                    2 * line_h <= pdf.h - pdf.b_margin - pdf.t_margin + 1e-7 and
                    pdf.y + 2 * line_h > pdf.h - pdf.b_margin + 1e-7):
                getattr(pdf, "_advance_region", pdf.add_page)()
            row_w = sum(seg[2] for seg in row)
            offset = first_offset if index == 0 else 0.0
            x_start = pdf.l_margin + offset
            if align == "C":
                x_start += (usable_w - offset - row_w - 2 * pdf.c_margin) / 2
            elif align == "R":
                x_start += usable_w - offset - row_w - 2 * pdf.c_margin
            self._render_segment_row(pdf, row, line_h, fs, at_x=x_start)
            if trailing and index == len(rows) - 1:
                right_x = pdf.l_margin + trailing_right - right_w - pdf.c_margin
                pdf.set_y(pdf.get_y() - line_h)
                if row and field_tab is not None and field_tab.leader:
                    self._draw_toc_leader(
                        pdf, segments[tab_idx][0], field_tab.leader, line_h,
                        pdf.l_margin + offset + row_w + 2 * pdf.c_margin, right_x,
                    )
                self._render_segment_row(pdf, trailing, line_h, fs, at_x=right_x)
        pdf.set_text_color(0, 0, 0)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _decimal_prefix_width(pdf: FPDF, segments: list) -> float:
        width = 0.0
        for segment in segments:
            run, text, seg_w, _, _ = segment
            separator = next((i for i, char in enumerate(text) if char in ".,"), None)
            if separator is not None:
                if isinstance(segment, _ShapedSegment):
                    return width + segment.fragment.get_width(end=separator)
                apply_run_font(pdf, run.font, default_size=DEFAULT_FONT_SIZE_PT)
                return width + pdf.get_string_width(text[:separator])
            width += seg_w
        return width

    def _draw_toc_leader(self, pdf: FPDF, run: ldm.Run, leader: int,
                         line_h: float, x1: float, x2: float) -> None:
        styles = {1: (0.005, 0.16, -0.06, 0.05),
                  2: (0.22, 0.14, -0.25, 0.04),
                  3: (0, 0, 0.08, 0.04),
                  4: (0, 0, 0.08, 0.08),
                  5: (0.005, 0.16, -0.30, 0.05)}
        if leader not in styles or x2 <= x1:
            return
        dash, gap, shift, width = styles[leader]
        size = (run.font.size if run.font.size > 0 else DEFAULT_FONT_SIZE_PT) * PT_TO_MM
        x1 += width * size / 2
        x2 -= width * size / 2
        if x2 <= x1:
            return
        metric = getattr(pdf, "_text_baseline_size", None)
        baseline = metric * PT_TO_MM if metric else size
        y = pdf.y + line_h / 2 + 0.3 * baseline + shift * size
        # Vector decoration keeps leader characters out of text extraction.
        with self._writer._artifact(pdf), pdf.local_context(
                line_width=width * size, draw_color=parse_color(run.font.color) or (0, 0, 0),
                stroke_cap_style=StrokeCapStyle.ROUND if leader in (1, 5) else StrokeCapStyle.BUTT,
                dash_pattern={'dash': dash * size, 'gap': gap * size}):
            pdf.line(x1, y, x2, y)

    @staticmethod
    def wrap_segments(pdf: FPDF, segments: list, width: float, *, first_width=None) -> list:
        """Wrap styled Unicode runs using actual glyph widths, retaining every character."""
        if width <= 0:
            raise ValueError("No usable text width")
        if pdf.text_shaping:
            return RunRenderer._wrap_shaped_segments(pdf, segments, width, first_width)
        saved_vpos = pdf.char_vpos
        saved_color = pdf.text_color
        saved_font = (pdf.font_family, pdf.font_style + ("U" if pdf.underline else "") + ("S" if pdf.strikethrough else ""), pdf.font_size_pt)
        rows, row, used = [], [], 0.0
        available = width if first_width is None else max(0.0, first_width)
        try:
            # ponytail: oversized words use character breaks; typography-specific punctuation rules are deferred.
            for run, text, _, size, link in segments:
                apply_run_font(pdf, run.font, default_size=size)
                characters = set(text) - {'\n'}
                if not pdf._fallback_font_ids and pdf.char_vpos == CharVPos.LINE:
                    cache = getattr(pdf, '_plain_glyph_widths', None)
                    if cache is None:
                        cache = pdf._plain_glyph_widths = {}
                    context = (pdf.current_font, pdf.font_size_pt, pdf.font_stretching, pdf.char_spacing, pdf.k)
                    widths = {}
                    for char in characters:
                        key = (*context, char)
                        value = cache.get(key)
                        if value is None:
                            value = pdf.get_string_width(char)
                            if len(cache) >= 4096:
                                cache.clear()
                            cache[key] = value
                        widths[char] = value
                else:
                    widths = {char: pdf.get_string_width(char) for char in characters}
                for token in re.findall(r"\n|[^\S\n]+|[^\s]+", text):
                    token_width = sum(widths.get(char, 0) for char in token)
                    if (used and not token.isspace() and token_width <= width
                            and used + token_width > available + 1e-7):
                        rows.append(row)
                        row, used, available = [], 0.0, width
                    for char in token:
                        if char == "\n":
                            rows.append(row)
                            row, used, available = [], 0.0, width
                            continue
                        char_w = widths[char]
                        if char_w > width:
                            raise ValueError("A glyph is wider than the usable text area")
                        if used + char_w > available + 1e-7:
                            rows.append(row)
                            row, used, available = [], 0.0, width
                        if row and row[-1][0] is run and row[-1][4] == link:
                            previous = row[-1]
                            row[-1] = (run, previous[1] + char, previous[2] + char_w, size, link)
                        else:
                            row.append((run, char, char_w, size, link))
                        used += char_w
            rows.append(row)
        finally:
            pdf.char_vpos = saved_vpos
            pdf.text_color = saved_color
            if saved_font[0]:
                pdf.set_font(*saved_font)
        return rows

    @staticmethod
    def _wrap_shaped_segments(pdf: FPDF, segments: list, width: float, first_width) -> list:
        """Use fpdf2\'s shaping-aware line breaker, including fallback-font metrics."""
        grouped = []
        # Run boundaries with identical formatting must not break shaping context.
        for (_, size, link), group in groupby(segments, key=lambda item: (item[0].font, item[3], item[4])):
            items = list(group)
            grouped.append((items[0][0], "".join(item[1] for item in items), 0.0, size, link))
        segments = grouped
        paragraph_text = "".join(segment[1] for segment in segments)
        if not paragraph_text:
            return [[]]
        contextual = any(left[1] and right[1] and not left[1][-1].isspace()
                         and not right[1][0].isspace()
                         for left, right in zip(segments, segments[1:]))
        saved_vpos = pdf.char_vpos
        saved_color = pdf.text_color
        saved_font = (pdf.font_family, pdf.font_style + ("U" if pdf.underline else "") + ("S" if pdf.strikethrough else ""), pdf.font_size_pt)
        fragments = []
        try:
            paragraph = BidiParagraph(
                paragraph_text,
                base_direction=pdf.text_shaping["direction"], preserve_bn_chars=True,
                alias=pdf.str_alias_nb_pages,
            )
            pdf.text_shaping["paragraph_direction"] = paragraph.base_direction
            directions = iter(paragraph.get_bidi_fragments())
            directional_text, direction = next(directions)
            remaining = len(directional_text)
            source_index = 0
            for index, (run, text, _, size, _) in enumerate(segments):
                apply_run_font(pdf, run.font, default_size=size)
                offset = 0
                while offset < len(text):
                    count = min(len(text) - offset, remaining)
                    pdf.text_shaping["fragment_direction"] = direction
                    for fragment in pdf._preload_font_styles(text[offset:offset + count], False):
                        if contextual and type(fragment) is Fragment:
                            fragment.font.__class__ = ContextualFont
                            fragment.graphics_state.text_shaping['_source_text'] = paragraph_text
                            fragment = ContextualFragment(
                                [SourceCharacter(char, source_index + i)
                                 for i, char in enumerate(fragment.characters)],
                                fragment.graphics_state, fragment.k, fragment.link,
                            )
                        source_index += len(fragment.characters)
                        fragment.link = index  # Keep source-run identity across native fragment clones.
                        fragments.append(fragment)
                    offset += count
                    remaining -= count
                    if remaining == 0:
                        directional_text, direction = next(directions, ("", None))
                        remaining = len(directional_text)
            breaker = MultiLineBreak(fragments, width, margins=(0, 0),
                                     first_line_indent=0 if first_width is None else width - first_width)
            rows = []
            while line := breaker.get_line():
                row = []
                if contextual:
                    positions = [char.source_index for fragment in line.fragments for char in fragment.characters
                                 if isinstance(char, SourceCharacter)]
                    for fragment in line.fragments:
                        if isinstance(fragment, ContextualFragment) and positions:
                            fragment.graphics_state.text_shaping['_source_line'] = min(positions), max(positions) + 1
                for fragment in line.get_ordered_fragments():
                    if not fragment.characters:
                        continue  # A zero-width cell would advance to the right margin.
                    run, _, _, size, link = segments[fragment.link]
                    row.append(_ShapedSegment((run, fragment.string, fragment.get_width(), size, link), fragment))
                rows.append(row)
            return rows or [[]]
        finally:
            pdf.char_vpos = saved_vpos
            pdf.text_color = saved_color
            if saved_font[0]:
                pdf.set_font(*saved_font)

    def _render_segment_row(
        self,
        pdf: FPDF,
        segments: list[Tuple[ldm.Run, str, float, float, Optional[str]]],
        line_h: float,
        fs: float,
        *,
        at_x: float,
    ) -> None:
        """Emit *segments* on a single line starting at *at_x*."""
        offset = at_x - pdf.l_margin
        pdf.set_x(at_x)
        pdf.cell(w=0, h=line_h)  # Break pages/columns before painting the background.
        pdf.set_x(pdf.l_margin + offset)
        for segment in segments:
            run, safe, seg_w, size, link = segment
            apply_run_font(pdf, run.font, default_size=fs)
            highlight = parse_color(run.font.highlight_color)
            if highlight:
                self._draw_highlight(pdf, safe, line_h, highlight)
            if link is not None:
                pdf.set_text_color(*HYPERLINK_TEXT_RGB)
            link_target = self._writer._link_target_for(pdf, link)
            if isinstance(segment, _ShapedSegment):
                fragment = segment.fragment
                fragment.link = None  # Source indices are not PDF link IDs.
                fragment.graphics_state.text_color = pdf.text_color
                pdf._render_styled_text_line(
                    TextLine((fragment,), 0, 0, Align.L, line_h, seg_w), line_h, link=link_target)
            else:
                pdf.cell(w=seg_w, h=line_h, text=safe, link=link_target)
            # shortcut: fpdf2 fallback cells cache fonts inside q/Q; recheck on upgrades.
            if self._writer.options.fallback_fonts:
                pdf.current_font_is_set_on_page = False
        pdf.ln(line_h)

    @staticmethod
    def _find_gap_split(
        segments: list[Tuple[ldm.Run, str, float, float, Optional[str]]],
        usable_w: float,
    ) -> Optional[int]:
        """Return the index of the whitespace segment that splits a right-
        aligned line into left / right halves, or ``None`` if no such gap
        exists.
        """
        total_w = sum(w for _, _, w, _, _ in segments)
        if total_w <= usable_w:
            return None
        best = None
        best_w = 0.0
        for i, (_, text, w, _, _) in enumerate(segments):
            if len(text) < GAP_MIN_SPACES or not text.isspace():
                continue
            if i == 0 or i == len(segments) - 1:
                continue
            if w > best_w:
                best = i
                best_w = w
        return best

    @staticmethod
    def _draw_highlight(
        pdf: FPDF,
        text: str,
        line_h: float,
        rgb: Tuple[int, int, int],
    ) -> None:
        """Paint a single highlight rectangle behind *text*."""
        if not text:
            return
        w = pdf.get_string_width(text)
        if w <= 0:
            return
        # ``pdf.cell`` renders text shifted right by ``c_margin`` (1 mm by
        # default).  Align the fill rectangle with that text origin so the
        # highlight tracks the glyphs instead of leading them.
        x = pdf.get_x() + pdf.c_margin
        y = pdf.get_y()
        pdf.set_fill_color(*rgb)
        pdf.rect(x, y + line_h * HIGHLIGHT_Y_OFFSET_RATIO, w, line_h * HIGHLIGHT_HEIGHT_RATIO, "F")

    def _write_with_highlight(
        self,
        pdf: FPDF,
        text: str,
        line_h: float,
        rgb: Tuple[int, int, int],
        *,
        run: ldm.Run,
        link: Union[int, str] = "",
    ) -> None:
        """Write *text* with a filled highlight background.

        Paints the fill rectangle separately and writes the text on top so
        the rectangle stays aligned with the glyphs.  ``pdf.cell`` offsets
        text by ``c_margin`` (1 mm by default); using ``fill=True`` would
        anchor the fill at the cell's left edge instead, producing a
        half-character drift between text and highlight.
        """
        width = pdf.epw - 2 * pdf.c_margin
        first_width = pdf.w - pdf.r_margin - pdf.get_x() - 2 * pdf.c_margin
        segments = [(run, text, 0.0, pdf.font_size_pt, None)]
        rows = self.wrap_segments(pdf, segments, width, first_width=first_width)
        for index, row in enumerate(rows):
            if index:
                pdf.ln(line_h)
                pdf.set_x(pdf.l_margin)
            offset = pdf.get_x() - pdf.l_margin
            pdf.cell(w=0, h=line_h)
            pdf.set_x(pdf.l_margin + offset)
            chunk = "".join(seg[1] for seg in row)
            chunk_width = sum(seg[2] for seg in row)
            pdf.set_fill_color(*rgb)
            self._fill_text_rect(pdf, pdf.get_x(), pdf.get_y(), chunk_width, line_h)
            pdf.cell(w=chunk_width, h=line_h, text=chunk, link=link)
            if self._writer.options.fallback_fonts:
                pdf.current_font_is_set_on_page = False

    @staticmethod
    def _fill_text_rect(pdf: FPDF, x: float, y: float, w: float, h: float) -> None:
        """Paint the highlight rectangle behind a cell of width *w*.

        Shifts by ``pdf.c_margin`` to match the text origin used by
        ``pdf.cell``.
        """
        if w <= 0:
            return
        pdf.rect(x + pdf.c_margin, y, w, h, "F")

    def _resolve_run_text(self, pdf: FPDF, text: str) -> str:
        """Resolve reader-emitted sentinels to their live page-local values."""
        if PAGE_FIELD_SENTINEL in text:
            logical = pdf.page_no() + self._writer._page_number_offset
            text = text.replace(PAGE_FIELD_SENTINEL, str(logical))
        return text
