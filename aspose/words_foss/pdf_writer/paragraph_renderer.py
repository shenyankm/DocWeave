"""Paragraph rendering for the PDF writer.

Contains the unified styled-block rendering logic used by both the
legacy path and the content-sequence (mixed image+text) path,
eliminating the previous code duplication.
"""


from contextlib import contextmanager
from typing import Optional

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._visible_runs import is_horizontal_rule_shape, visible_runs
from aspose.words_foss.model.wrap_type import WrapType
from aspose.words_foss.pdf_writer.baseline import baseline_scope, mixed_size
from aspose.words_foss.pdf_writer.constants import (
    CODE_BLOCK_BG_RGB,
    DEFAULT_FONT_NAME,
    DEFAULT_FONT_SIZE_PT,
    DEFAULT_QUOTE_INDENT_MM,
    FPDF_ALIGN,
    LINE_HEIGHT_FACTOR,
    BOTTOM_BORDER_SLOT,
    HORIZONTAL_RULE_CHARS,
    HORIZONTAL_RULE_MIN_WIDTH_PT,
    HORIZONTAL_RULE_RGB,
    LIST_INDENT_PER_LEVEL_MM,
    PT_TO_MM,
    QUOTE_TEXT_RGB,
    WORD_LINE_SPACING_DIVISOR,
)
from aspose.words_foss.pdf_writer._context import PDFWriterContext
from aspose.words_foss.pdf_writer.font import apply_run_font, reset_font
from aspose.words_foss.pdf_writer.page_bands import register_bookmarks
from aspose.words_foss.model.enums import LineSpacingRule
from aspose.words_foss.pdf_writer.text import (
    apply_caps,
    extract_link_segments,
    get_dominant_color,
    get_dominant_font_size,
    get_line_font_size,
    is_pure_page_break,
    is_toc_style,
    safe_text,
)


# Mirror of ``NumberStyle``; unknown styles fall back to decimal.
def _format_number(num: int, number_style: int) -> str:
    if number_style == 0:  # ARABIC
        return str(num)
    if number_style == 1:  # UPPERCASE_ROMAN
        return _to_roman(num).upper()
    if number_style == 2:  # LOWERCASE_ROMAN
        return _to_roman(num).lower()
    if number_style == 3:  # UPPERCASE_LETTER
        return _to_alpha(num).upper()
    if number_style == 4:  # LOWERCASE_LETTER
        return _to_alpha(num).lower()
    return str(num)


def _to_roman(num: int) -> str:
    if num <= 0:
        return ""
    pairs = (
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    )
    out = []
    for value, sym in pairs:
        while num >= value:
            out.append(sym)
            num -= value
    return "".join(out)


def _to_alpha(num: int) -> str:
    if num <= 0:
        return ""
    out = ""
    while num > 0:
        num, rem = divmod(num - 1, 26)
        out = chr(ord("a") + rem) + out
    return out


def _expand_number_format(
    fmt: str,
    lst: ldm.DocList,
    counters: dict[tuple[int, int], int],
    current_level: int,
    current_style: int,
) -> str:
    """Replace ``%N`` placeholders with the counter for level ``N-1``."""
    out = []
    i = 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == "%" and i + 1 < len(fmt) and fmt[i + 1].isdigit():
            lvl_idx = int(fmt[i + 1]) - 1
            if 0 <= lvl_idx < len(lst.list_levels):
                style = lst.list_levels[lvl_idx].number_style
            else:
                style = current_style
            counter = counters.get((lst.list_id, lvl_idx), 0)
            if counter == 0 and lvl_idx == current_level:
                counter = 1
            out.append(_format_number(counter, style))
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)


class ParagraphRenderer:
    """Renders LDM paragraphs into PDF."""

    def __init__(self, writer: PDFWriterContext) -> None:
        self._writer = writer

    # ------------------------------------------------------------------
    # Line height computation
    # ------------------------------------------------------------------

    @staticmethod
    def line_height_mm(size_pt: float, pf: ldm.ParagraphFormat) -> float:
        """Return the line height in mm for a run of *size_pt* inside *pf*.

        Honours Word's ``w:spacing w:lineRule="exact"`` (``rule==1``), which
        pins the line advance to a fixed point value.
        ``atLeast`` (0) acts as a floor above the natural leading;
        ``multiple`` (2) scales the natural leading by ``line_spacing/240``.
        """

        natural = size_pt * PT_TO_MM * LINE_HEIGHT_FACTOR
        ls = pf.line_spacing
        if ls <= 0:
            return natural
        if pf.line_spacing_rule == LineSpacingRule.EXACTLY:
            return ls * PT_TO_MM
        if pf.line_spacing_rule == LineSpacingRule.AT_LEAST:
            return max(natural, ls * PT_TO_MM)
        if pf.line_spacing_rule == LineSpacingRule.MULTIPLE:
            return natural * (ls / WORD_LINE_SPACING_DIVISOR) if ls else natural
        return natural

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def render_paragraph(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Render a single paragraph into the PDF."""
        if not para.list_format or not para.list_format.is_list_item:
            self._writer._end_list(pdf)
        in_band = getattr(pdf, '_in_header_render', False) or getattr(pdf, '_in_footer_render', False)
        if not in_band and para.paragraph_format.page_break_before and (
                pdf.get_y() > pdf.t_margin + 1e-7 or
                self._writer._page_margin_left > self._writer._page_full_margin_left + 1e-7):
            pdf.add_page()
        if is_pure_page_break(para):
            if not in_band:
                pdf.add_page()
            return

        if not in_band and para.paragraph_format.keep_together:
            self._maybe_keep_together(pdf, para)

        # Capture the paragraph's top Y before body rendering — needed
        # to position anchored shapes with
        # ``relative_vertical_position == Paragraph`` ().
        paragraph_top_y = pdf.get_y()

        with self._track_navigation_start(pdf, para):
            register_bookmarks(pdf, para, self._writer._anchor_links)
            self._emit_bookmark_outlines(pdf, para)
            self._render_paragraph_body(pdf, para)

        self._render_paragraph_relative_shapes(pdf, para, paragraph_top_y)

        if not in_band and any("\f" in (run.text or "") for run in visible_runs(para)):
            pdf.add_page()

    @contextmanager
    def _track_navigation_start(self, pdf: FPDF, para: ldm.Paragraph):
        if getattr(pdf, '_in_repeated_page_band', False) or (
                not para.paragraph_format.is_heading and not any(
                    isinstance(item, ldm.BookmarkStart) and item.name for item in para._children)):
            yield
            return
        outline_start = len(pdf._outline)
        original = pdf._perform_page_break_if_need_be
        overridden = "_perform_page_break_if_need_be" in pdf.__dict__
        captured = checking = False

        # shortcut: fpdf2 has no public first-draw hook; verify this integration on upgrades.
        def after_break(h):
            nonlocal captured, checking
            if captured or checking:
                return original(h)
            # A break can render header/footer outlines; retain only the paragraph's entries.
            sections = list(pdf._outline[outline_start:])
            checking = True
            try:
                result = original(h)
            finally:
                checking = False
            captured = True
            register_bookmarks(pdf, para, self._writer._anchor_links)
            for section in sections:
                section.page_number = section.dest.page_number = pdf.page_no()
                section.dest.top = pdf.h_pt - pdf.get_y() * pdf.k
                section.dest.left = pdf.get_x() * pdf.k
            return result

        pdf._perform_page_break_if_need_be = after_break
        try:
            yield
        finally:
            if overridden:
                pdf._perform_page_break_if_need_be = original
            else:
                pdf.__dict__.pop("_perform_page_break_if_need_be", None)

    def _render_paragraph_relative_shapes(
        self, pdf: FPDF, para: ldm.Paragraph, paragraph_top_y: float
    ) -> None:
        """Draw shapes anchored ``relative_vertical_position == Paragraph``.

        The shape's ``top`` field carries the raw OOXML
        ``<wp:positionV/posOffset>`` in points; convert to mm and add
        the paragraph's actual top Y so the shape lands where Word
        places it.
        """
        extras = [
            e for e in para._children
            if isinstance(e, ldm.Shape) and e._is_positioned
            and e.relative_vertical_position == 2
        ]
        if not extras:
            return
        saved_x, saved_y = pdf.get_x(), pdf.get_y()
        for shape in extras:
            y_mm = paragraph_top_y + shape.top * PT_TO_MM
            self._writer._shape_renderer.render_positioned_shape(
                pdf, shape, y_override=y_mm,
            )
        pdf.set_xy(saved_x, saved_y)

    # ------------------------------------------------------------------
    # Keep-together support
    # ------------------------------------------------------------------

    def _maybe_keep_together(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """If the paragraph won't fit on this page, move to the next one."""
        w = self._writer
        pf = para.paragraph_format
        runs = visible_runs(para)
        text = self._plain_text_from_runs(runs)
        if not text.strip():
            return
        usable_w = w._page_width - w._page_margin_left - w._page_margin_right
        est_height = w._estimate_paragraph_height(para, usable_w)
        remaining = w._page_height - w._page_margin_bottom - pdf.get_y()
        if est_height > remaining + 1e-7 and est_height <= (w._page_height - w._page_margin_bottom - pdf.t_margin) + 1e-7:
            getattr(pdf, "_advance_region", pdf.add_page)()

    # ------------------------------------------------------------------
    # Outline helpers
    # ------------------------------------------------------------------

    def _effective_headings_outline_levels(self) -> int:
        """Return the effective headings outline depth."""
        oo = self._writer.options.outline_options
        if oo.headings_outline_levels > 0:
            return oo.headings_outline_levels
        if self._writer.options.export_document_structure:
            return 9
        return 0

    def _effective_bookmarks_outline_level(self) -> int:
        """Return the effective default bookmark outline level."""
        oo = self._writer.options.outline_options
        if oo.default_bookmarks_outline_level > 0 or oo.bookmarks_outline_levels:
            return oo.default_bookmarks_outline_level
        return 1 if self._writer.options.export_bookmarks_outline else 0

    def _start_section_safe(self, pdf: FPDF, text: str, level: int) -> None:
        """Map source levels to a contiguous PDF hierarchy, optionally filling gaps."""
        if getattr(pdf, '_in_repeated_page_band', False):
            return
        w = self._writer
        levels = w._outline_levels
        if w.options.outline_options.create_missing_outline_levels:
            for gap in range((levels[-1] if levels else 0) + 1, level):
                levels.append(gap)
                pdf.start_section("", level=len(levels) - 1)
                pdf._outline[-1].dest.left = pdf.get_x() * pdf.k
        while levels and levels[-1] >= level:
            levels.pop()
        levels.append(level)
        pdf.start_section(text, level=len(levels) - 1)
        pdf._outline[-1].dest.left = pdf.get_x() * pdf.k

    def _emit_heading_outline(self, pdf: FPDF, text: str, level: int) -> None:
        """Emit a PDF outline entry for a heading if allowed by OutlineOptions."""
        w = self._writer
        oo = w.options.outline_options
        max_level = self._effective_headings_outline_levels()
        if level > max_level:
            return
        if w._in_table and not oo.create_outlines_for_headings_in_tables:
            return
        self._start_section_safe(pdf, text, level)

    def _emit_bookmark_outlines(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Emit outline entries for bookmarks in this paragraph."""
        if para.paragraph_format.is_heading:
            return
        oo = self._writer.options.outline_options
        default_level = self._effective_bookmarks_outline_level()
        for extra in para._children:
            if not isinstance(extra, ldm.BookmarkStart) or not extra.name:
                continue
            if extra.name.startswith("_"):
                continue
            bm_level = oo.bookmarks_outline_levels.get(extra.name, default_level)
            if bm_level <= 0:
                continue
            self._start_section_safe(pdf, extra.name, bm_level)

    # ------------------------------------------------------------------
    # Body dispatcher
    # ------------------------------------------------------------------

    def _render_paragraph_body(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Render paragraph body: floating images, then styled content."""
        w = self._writer

        # Draw floating (wrapNone) images at their anchor coordinates
        w._shape_renderer.render_floating_images(pdf, para)

        # Anchored (is_inline=False) wrapped images float — don't take inline space.
        w._shape_renderer.render_anchored_wrapped_shapes(pdf, para)

        # ``is_inline=None`` (untagged shapes from older fixtures) is treated as inline.
        has_mixed = any(
            isinstance(i, ldm.Shape)
            and i.has_image
            and i.is_inline is not False
            and not i._is_positioned
            and i.wrap_type != WrapType.NONE
            for i in para._children
        )
        if has_mixed:
            self._render_content_sequence(pdf, para)
            return

        # Legacy path: no inline images — emit any standalone shapes first.
        for item in para._children:
            if isinstance(item, ldm.Shape):
                if item._is_positioned or item.wrap_type == WrapType.NONE:
                    continue
                if item.is_inline is False:
                    # Anchored/floating; already drawn above.
                    continue
                if item.has_image and item.image_data is not None:
                    w._shape_renderer.render_shape(pdf, item)
                elif item.text_box or is_horizontal_rule_shape(item):
                    w._shape_renderer.render_shape(pdf, item)

        pf = para.paragraph_format
        align = FPDF_ALIGN.get(pf.alignment, "L")

        # Ensure each paragraph starts at the left margin.
        pdf.set_x(w._page_margin_left)

        self._render_styled_block(
            pdf, visible_runs(para), pf, para.list_format, para.list_label, align
        )

    # ------------------------------------------------------------------
    # Unified styled-block rendering (DRY fix)
    # ------------------------------------------------------------------

    def _render_styled_block(
        self,
        pdf: FPDF,
        runs: list[ldm.Run],
        pf: ldm.ParagraphFormat,
        list_format: Optional[ldm.ListFormat],
        list_label: Optional[ldm.ListLabel],
        align: str,
    ) -> None:
        """Apply paragraph margins, including after page/column changes."""
        w = self._writer
        previous = w._paragraph_insets
        left = pf.left_indent * PT_TO_MM
        if not pf.is_heading and "Quote" in (pf.style_name or "") and left <= 0:
            left = DEFAULT_QUOTE_INDENT_MM
        if list_format and list_format.is_list_item:
            left = 0.0  # The list branch positions its own marker and text.
        w._paragraph_insets = (left, pf.right_indent * PT_TO_MM)
        pdf.set_left_margin(w._page_margin_left + left)
        pdf.set_right_margin(w._page_margin_right + pf.right_indent * PT_TO_MM)
        pdf.set_x(pdf.l_margin)
        try:
            self._render_styled_content(pdf, runs, pf, list_format, list_label, align)
        finally:
            w._paragraph_insets = previous
            pdf.set_left_margin(w._page_margin_left + previous[0])
            pdf.set_right_margin(w._page_margin_right + previous[1])

    def _render_styled_content(
        self, pdf: FPDF, runs: list[ldm.Run], pf: ldm.ParagraphFormat,
        list_format: Optional[ldm.ListFormat], list_label: Optional[ldm.ListLabel], align: str,
    ) -> None:
        """Shared heading/code/quote/list/normal rendering logic.

        This single method replaces the duplicated branching that
        previously existed in both ``_render_paragraph_body`` and
        ``_render_content_sequence``'s ``_flush_styled`` closure.
        """
        w = self._writer
        fs = DEFAULT_FONT_SIZE_PT
        style_name = pf.style_name
        if (pf.is_heading or any(name in (style_name or '') for name in ('Code', 'code', 'Quote')) or
                not list_format or not list_format.is_list_item or self._is_horizontal_rule(pf, runs)):
            w._end_list(pdf)

        # Apply space before
        self._apply_space_before(pdf, pf)

        if self._is_horizontal_rule(pf, runs):
            self._draw_horizontal_rule(pdf, pf)
            return

        # Heading
        if pf.is_heading:
            level = min(pf.outline_level + 1, 6)
            run_size = get_dominant_font_size(runs)
            size = run_size if run_size > 0 else fs + (6 - level) * 2
            line_h = self.line_height_mm(size, pf)
            text = self._plain_text_from_runs(runs)
            if text.strip():
                self._emit_heading_outline(pdf, text.strip(), level)
            with w._tag(pdf, f"/H{level}", title=text.strip()):
                heading_color = get_dominant_color(runs)
                if runs:
                    apply_run_font(pdf, runs[0].font, default_size=size)
                else:
                    pdf.set_font(DEFAULT_FONT_NAME, style="B", size=size)
                if heading_color:
                    pdf.set_text_color(*heading_color)
                pdf.multi_cell(w=0, h=line_h, text=safe_text(text), align=align)
            self._apply_space_after(pdf, pf)
            reset_font(pdf)
            return

        # Code block
        is_code = bool(style_name and ("Code" in style_name or "code" in style_name))
        if is_code:
            code_size = fs - 1
            line_h = self.line_height_mm(code_size, pf)
            text = self._plain_text_from_runs(runs)
            with w._tag(pdf, "/Code"):
                pdf.set_fill_color(*CODE_BLOCK_BG_RGB)
                pdf.set_font(DEFAULT_FONT_NAME, size=code_size)
                usable_w = pdf.epw
                pdf.multi_cell(
                    w=usable_w,
                    h=line_h,
                    text=safe_text(text),
                    fill=True,
                    align=align,
                )
            self._apply_space_after(pdf, pf)
            reset_font(pdf)
            return

        # Block quote
        if style_name and "Quote" in style_name:
            line_h = self.line_height_mm(fs, pf)
            text = self._plain_text_from_runs(runs)
            with w._tag(pdf, "/BlockQuote"):
                pdf.set_font(DEFAULT_FONT_NAME, style="I", size=fs)
                pdf.set_text_color(*QUOTE_TEXT_RGB)
                pdf.multi_cell(w=0, h=line_h, text=safe_text(text), align=align)
            self._apply_space_after(pdf, pf)
            reset_font(pdf)
            return

        # List item
        if list_format and list_format.is_list_item:
            level = list_format.list_level_number
            # Marker at ``npos`` (left_indent + first_line_indent), text at ``tpos`` (left_indent).
            text_indent_mm = (
                pf.left_indent * PT_TO_MM
                if pf.left_indent > 0
                else LIST_INDENT_PER_LEVEL_MM * (level + 1)
            )
            marker_indent_mm = max(0.0, text_indent_mm + pf.first_line_indent * PT_TO_MM)
            label = list_label.label_string if list_label and list_label.label_string else ""
            if not label:
                label = self._compute_list_label(list_format)
            baseline = mixed_size(run.font.size if run.font.size > 0 else fs for run in runs if run.text)
            with baseline_scope(pdf, baseline), w._list_structure(pdf, list_format):
                run_size = get_dominant_font_size(runs)
                effective_fs = run_size if run_size > 0 else fs
                line_h = self.line_height_mm(get_line_font_size(runs), pf)
                pdf.set_font(DEFAULT_FONT_NAME, size=effective_fs)
                label_text = safe_text(f"{label} ") if label else ""
                label_width = pdf.get_string_width(label_text) if label_text else 0.0
                text_indent_mm = max(text_indent_mm, marker_indent_mm + label_width)
                if (text_indent_mm + 2 * pdf.c_margin >= pdf.epw or
                        marker_indent_mm + label_width + 2 * pdf.c_margin > pdf.epw):
                    raise ValueError("No usable text width after body list indents")
                if pdf.auto_page_break:
                    needed = line_h
                    region_height = pdf.h - pdf.b_margin - pdf.t_margin
                    if (pf.widow_control and not any('\t' in (run.text or '') for run in runs)
                            and 2 * line_h <= region_height + 1e-7
                            and pdf.y + 2 * line_h > pdf.h - pdf.b_margin + 1e-7):
                        body_pf = pf.model_copy(update={'left_indent': text_indent_mm / PT_TO_MM,
                            'first_line_indent': 0.0, 'space_before': 0.0, 'space_after': 0.0})
                        body = ldm.Paragraph(children=runs, paragraph_format=body_pf)
                        width = w._page_width - w._page_margin_left - w._page_margin_right
                        if w._estimate_paragraph_height(body, width) > line_h + 1e-7:
                            needed = 2 * line_h
                    if needed <= region_height + 1e-7 and pdf.y + needed > pdf.h - pdf.b_margin + 1e-7:
                        getattr(pdf, '_advance_region', pdf.add_page)()
                pdf.set_x(pdf.l_margin + marker_indent_mm)
                if label:
                    with w._tag(pdf, '/Lbl'):
                        pdf.cell(w=label_width + 2 * pdf.c_margin, h=line_h, text=label_text)
                previous = w._paragraph_insets
                w._paragraph_insets = (text_indent_mm, previous[1])
                pdf.set_left_margin(w._page_margin_left + text_indent_mm)
                pdf.set_x(pdf.l_margin)
                try:
                    with w._tag(pdf, '/LBody'):
                        w._run_renderer.render_formatted_runs(
                            pdf, runs, align=align, newline=True,
                            line_h_override=line_h, pf=pf,
                        )
                finally:
                    w._paragraph_insets = previous
                    pdf.set_left_margin(w._page_margin_left + previous[0])
            self._apply_space_after(pdf, pf)
            return

        # Normal paragraph — render with inline formatting
        text = self._plain_text_from_runs(runs)
        if not text.strip():
            run_size = get_dominant_font_size(runs) or fs
            pdf.ln(self.line_height_mm(run_size, pf))
            return

        # Apply first-line indent.
        pdf.set_x(pdf.l_margin + pf.first_line_indent * PT_TO_MM)

        line_h = self.line_height_mm(get_line_font_size(runs), pf)

        with w._tag(pdf, "/P"):
            w._run_renderer.render_formatted_runs(
                pdf,
                runs,
                align=align,
                line_h_override=line_h,
                is_toc=is_toc_style(style_name),
                tab_stops=pf.tab_stops if pf.tab_stops else None,
                default_tab_stop=w._default_tab_stop,
                pf=pf,
            )
        self._apply_space_after(pdf, pf)

    # ------------------------------------------------------------------
    # Content-sequence path (mixed image+text paragraphs)
    # ------------------------------------------------------------------

    def _render_content_sequence(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Render a paragraph using content_sequence to preserve XML element order."""
        w = self._writer
        pf = para.paragraph_format
        align = FPDF_ALIGN.get(pf.alignment, "L")

        self._apply_space_before(pdf, pf)

        pending_runs: list[ldm.Run] = []

        def _flush_styled() -> None:
            if not pending_runs:
                return
            runs_snapshot = list(pending_runs)
            pending_runs.clear()
            plain = safe_text("".join(r.text or "" for r in runs_snapshot))
            if not plain:
                return
            self._render_styled_block(
                pdf, runs_snapshot, pf, para.list_format, para.list_label, align
            )

        for item in para._children:
            if isinstance(item, ldm.Shape):
                if item._is_positioned or item.wrap_type == WrapType.NONE:
                    continue  # drawn by positioned/floating pass
                if item.is_inline is False:
                    continue  # anchored/floating; drawn separately
                if item.has_image and item.image_data is not None:
                    _flush_styled()
                    w._shape_renderer.render_shape(pdf, item)
                elif item.text_box:
                    _flush_styled()
                    w._shape_renderer.render_shape(pdf, item)
            elif isinstance(item, ldm.Run):
                pending_runs.append(item)

        _flush_styled()

    # ------------------------------------------------------------------
    # List label helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _is_horizontal_rule(pf: ldm.ParagraphFormat, runs: list[ldm.Run]) -> bool:
        if not pf.borders or len(pf.borders) <= BOTTOM_BORDER_SLOT:
            return False
        bottom = pf.borders[BOTTOM_BORDER_SLOT]
        if bottom.line_style <= 0 or bottom.line_width < HORIZONTAL_RULE_MIN_WIDTH_PT:
            return False
        return not any((run.text or "").strip(HORIZONTAL_RULE_CHARS) for run in runs)

    def _draw_horizontal_rule(self, pdf: FPDF, pf: ldm.ParagraphFormat) -> None:
        w = self._writer
        y = pdf.get_y() + self.line_height_mm(DEFAULT_FONT_SIZE_PT, pf) / 2
        bottom = pf.borders[BOTTOM_BORDER_SLOT]
        pdf.set_line_width(bottom.line_width * PT_TO_MM)
        pdf.set_draw_color(*HORIZONTAL_RULE_RGB)
        pdf.line(w._page_margin_left, y, w._page_width - w._page_margin_right, y)
        pdf.set_y(y)
        self._apply_space_after(pdf, pf)

    def _compute_list_label(
        self, list_format: ldm.ListFormat
    ) -> Optional[str]:
        """Generate the marker for an ordered/bulleted list item."""
        w = self._writer
        doc = getattr(w, "_doc", None)
        if doc is None:
            return None
        list_id = list_format.list_id
        level = list_format.list_level_number
        lst = doc.get_list(list_id)
        if lst is None or not lst.list_levels:
            return None
        lvl = doc.resolve_list_level(list_id, level)
        if lvl is None:
            return None
        if level >= len(lst.list_levels):
            level = len(lst.list_levels) - 1
        if lvl.number_style in (23, 255):
            return lvl.number_format

        key = (list_id, level)
        counters = w._list_counters
        if key not in counters:
            counters[key] = max(lvl.start_at, 1)
        else:
            counters[key] += 1
        # Advancing a level resets all deeper-level counters.
        for k in list(counters.keys()):
            if k[0] == list_id and k[1] > level:
                counters.pop(k, None)

        fmt = lvl.number_format or "%1."
        return _expand_number_format(fmt, lst, counters, level, lvl.number_style)

    # ------------------------------------------------------------------
    # Spacing helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _apply_space_before(pdf: FPDF, pf: ldm.ParagraphFormat) -> None:
        """Add vertical space before a paragraph based on space_before."""
        if pf.space_before > 0:
            pdf.ln(pf.space_before * PT_TO_MM)

    @staticmethod
    def _apply_space_after(pdf: FPDF, pf: ldm.ParagraphFormat) -> None:
        """Add vertical space after a paragraph based on space_after."""
        if pf.space_after > 0:
            pdf.ln(pf.space_after * PT_TO_MM)

    # ------------------------------------------------------------------
    # Text helpers (thin wrappers for para-level use)
    # ------------------------------------------------------------------

    @staticmethod
    def _plain_text_from_runs(runs: list[ldm.Run]) -> str:
        """Return plain text from a list of runs, stripping Markdown links."""

        return "".join(
            apply_caps(chunk, run.font)
            for run in runs
            for chunk, _ in extract_link_segments(run.text or "")
        )
