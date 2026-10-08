"""Formatted-run rendering for the PDF writer.

Handles left-aligned writes, aligned cell rows, highlight rectangles,
and strikethrough lines.
"""


import re
from typing import Optional, Tuple, Union

from fpdf import FPDF
# ponytail: fpdf2 bidi/line-break internals require regression checks when upgrading fpdf2.
from fpdf.line_break import MultiLineBreak

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.color import parse_color
from aspose.words_foss.pdf_writer.constants import (
    DEFAULT_FONT_NAME,
    DEFAULT_FONT_SIZE_PT,
    GAP_MIN_SPACES,
    HIGHLIGHT_HEIGHT_RATIO,
    HIGHLIGHT_Y_OFFSET_RATIO,
    HYPERLINK_TEXT_RGB,
    LINE_HEIGHT_FACTOR,
    PT_TO_MM,
    STRIKETHROUGH_Y_RATIO,
)
from aspose.words_foss.pdf_writer.font import apply_run_font
from aspose.words_foss.pdf_writer.text import (
    apply_caps,
    extract_link_segments,
    safe_text,
)
from aspose.words_foss.docx_reader import PAGE_FIELD_SENTINEL
from aspose.words_foss.pdf_writer._context import PDFWriterContext


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
        if align != "L" or (has_tab and is_toc) or pdf.text_shaping:
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
            if "\t" in text:
                if tab_stops and tab_stops.tab_stops:
                    apply_run_font(pdf, font, default_size=fs)
                    pieces = text.split("\t")
                    for p_idx, piece in enumerate(pieces):
                        if p_idx > 0:
                            self._advance_to_tab_stop(
                                pdf, tab_stops, default_tab_stop, line_h,
                                upcoming_text=piece,
                            )
                        if piece:
                            resolved = self._resolve_run_text(pdf, piece)
                            if resolved:
                                resolved = apply_caps(resolved, font)
                                pdf.write(h=line_h, text=safe_text(resolved))
                    continue
                text = text.replace("\t", " ")
            text = self._resolve_run_text(pdf, text)
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
                elif font.strike_through:
                    x_before = pdf.get_x()
                    y_before = pdf.get_y()
                    pdf.write(h=line_h, text=safe, link=link_target)
                    x_after = pdf.get_x()
                    strike_y = y_before + size * PT_TO_MM * STRIKETHROUGH_Y_RATIO
                    pdf.line(x_before, strike_y, x_after, strike_y)
                else:
                    pdf.write(h=line_h, text=safe, link=link_target)

        # Reset color
        pdf.set_text_color(0, 0, 0)
        if newline:
            pdf.ln()

    def _advance_to_tab_stop(
        self,
        pdf: FPDF,
        tab_stops: ldm.TabStopCollection,
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
        pos_from_margin = current_x - margin_left
        pos_pt = pos_from_margin / PT_TO_MM

        next_tab = tab_stops.after(pos_pt)
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

        target_x = margin_left + target_pt * PT_TO_MM
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

        # Tab-based two-column layout (TOC entries): title on the left,
        # trailing content pinned to the right margin, same baseline.
        tab_idx = (
            next((i for i, (_, t, _, _, _) in enumerate(segments) if t == "\t"), None)
            if is_toc
            else None
        )
        if tab_idx is not None:
            left = [s for s in segments[:tab_idx] if s[1] != "\t"]
            right = [s for s in segments[tab_idx + 1 :] if s[1] != "\t"]
            self._render_segment_row(pdf, left, line_h, fs, at_x=pdf.l_margin)
            if right:
                right_w = sum(sw for _, _, sw, _, _ in right)
                right_x = max(
                    pdf.l_margin,
                    pdf.l_margin + usable_w - right_w,
                )
                pdf.set_y(pdf.get_y() - line_h)  # stay on the same baseline
                self._render_segment_row(pdf, right, line_h, fs, at_x=right_x)
            pdf.set_text_color(0, 0, 0)
            return

        # Recompute visible-only total after tab segments have been ruled out.
        visible = [s for s in segments if s[1] != "\t"]
        if not visible:
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
                    pdf.l_margin + usable_w - right_w,
                )
                pdf.set_y(pdf.get_y() - line_h)  # stay on the same baseline
                self._render_segment_row(pdf, right, line_h, fs, at_x=right_x)
                pdf.set_text_color(0, 0, 0)
                return
        first_offset = pdf.get_x() - pdf.l_margin
        rows = self.wrap_segments(pdf, visible, usable_w - 2 * pdf.c_margin,
                                  first_width=usable_w - first_offset - 2 * pdf.c_margin)
        for index, row in enumerate(rows):
            row_w = sum(seg[2] for seg in row)
            offset = first_offset if index == 0 else 0.0
            x_start = pdf.l_margin + offset
            if align == "C":
                x_start += (usable_w - offset - row_w) / 2
            elif align == "R":
                x_start += usable_w - offset - row_w - 2 * pdf.c_margin
            self._render_segment_row(pdf, row, line_h, fs, at_x=x_start)
        pdf.set_text_color(0, 0, 0)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    @staticmethod
    def wrap_segments(pdf: FPDF, segments: list, width: float, *, first_width=None) -> list:
        """Wrap styled Unicode runs using actual glyph widths, retaining every character."""
        if width <= 0:
            raise ValueError("No usable text width")
        if pdf.text_shaping:
            return RunRenderer._wrap_shaped_segments(pdf, segments, width, first_width)
        saved_font = (pdf.font_family, pdf.font_style + ("U" if pdf.underline else ""), pdf.font_size_pt)
        rows, row, used = [], [], 0.0
        available = width if first_width is None else max(0.0, first_width)
        try:
            # ponytail: oversized words use character breaks; typography-specific punctuation rules are deferred.
            for run, text, _, size, link in segments:
                style = ("B" if run.font.bold else "") + ("I" if run.font.italic else "")
                pdf.set_font(DEFAULT_FONT_NAME, style=style, size=size)
                widths = {char: pdf.get_string_width(char) for char in set(text) if char != "\n"}
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
            if saved_font[0]:
                pdf.set_font(*saved_font)
        return rows

    @staticmethod
    def _wrap_shaped_segments(pdf: FPDF, segments: list, width: float, first_width) -> list:
        """Use fpdf2\'s shaping-aware line breaker, including fallback-font metrics."""
        saved_font = (pdf.font_family, pdf.font_style + ("U" if pdf.underline else ""), pdf.font_size_pt)
        fragments = []
        try:
            for index, (run, text, _, size, _) in enumerate(segments):
                apply_run_font(pdf, run.font, default_size=size)
                for fragment in pdf._preload_bidirectional_text(text, False):
                    fragment.link = index  # Keep source-run identity across native fragment clones.
                    fragments.append(fragment)
            breaker = MultiLineBreak(fragments, width, margins=(0, 0),
                                     first_line_indent=0 if first_width is None else width - first_width)
            rows = []
            while line := breaker.get_line():
                row = []
                for fragment in line.get_ordered_fragments():
                    run, _, _, size, link = segments[fragment.link]
                    row.append((run, fragment.string, fragment.get_width(), size, link))
                rows.append(row)
            return rows or [[]]
        finally:
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
        for run, safe, seg_w, size, link in segments:
            apply_run_font(pdf, run.font, default_size=fs)
            highlight = parse_color(run.font.highlight_color)
            if highlight:
                self._draw_highlight(pdf, safe, line_h, highlight)
            if link is not None:
                pdf.set_text_color(*HYPERLINK_TEXT_RGB)
            link_target = self._writer._link_target_for(pdf, link)
            if run.font.strike_through:
                x_before = pdf.get_x()
                y_before = pdf.get_y()
                pdf.cell(w=seg_w, h=line_h, text=safe, link=link_target)
                x_after = pdf.get_x()
                strike_y = y_before + size * PT_TO_MM * STRIKETHROUGH_Y_RATIO
                pdf.line(x_before, strike_y, x_after, strike_y)
            else:
                pdf.cell(w=seg_w, h=line_h, text=safe, link=link_target)
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
