"""Table rendering for the PDF writer."""


from dataclasses import dataclass, field, replace
from io import BytesIO
from typing import Optional

from fpdf import FPDF
from aspose.words_foss.diagnostics import warn
from aspose.words_foss._io import MAX_TABLE_COLUMNS
from aspose.words_foss.pdf_writer.page_bands import register_bookmarks
from aspose.words_foss.pdf_writer.diagnostics import PdfContentLossWarning

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._visible_runs import visible_runs
from aspose.words_foss.model.enums import LineStyle, ParagraphAlignment
from aspose.words_foss.pdf_writer.color import parse_color
from aspose.words_foss.pdf_writer.constants import (
    DEFAULT_CELL_LINE_H_FACTOR,
    DEFAULT_SHAPE_DIM_PT,
    FPDF_ALIGN,
    LINE_HEIGHT_FACTOR,
    DEFAULT_CELL_PAD_LEFT_MM,
    DEFAULT_CELL_PAD_TOP_MM,
    DEFAULT_FONT_SIZE_PT,
    MIN_LINE_WIDTH_MM,
    MIN_ROW_HEIGHT_FACTOR,
    POST_TABLE_SPACING_MM,
    PT_TO_MM,
)
from aspose.words_foss.pdf_writer.font import apply_run_font, reset_font
from aspose.words_foss.pdf_writer._context import PDFWriterContext
from aspose.words_foss.pdf_writer.text import apply_caps, cell_text, extract_link_segments, safe_text

# Border slot indices in the LDM `borders` list, matching the canonical
# BorderType layout shared with the DOCX writer:
#   0=Bottom, 1=Left, 2=Right, 3=Top, 4=insideH, 5=insideV
_B_BOTTOM, _B_LEFT, _B_RIGHT, _B_TOP = 0, 1, 2, 3
_B_INSIDE_H, _B_INSIDE_V = 4, 5


# CellFormat.orientation values that require rotated text rendering.
# 1=btLr (bottom-to-top), 2=tbRl (top-to-bottom), 3=lrTbV, 4=tbRlV, 5=tbLrV
_VERTICAL_ORIENTATIONS = {1, 2, 3, 4, 5}

# Maps orientation enum to clockwise rotation degrees for fpdf2.
_ORIENTATION_TO_ANGLE: dict[int, float] = {
    1: 90.0,    # btLr: text reads bottom-to-top
    2: -90.0,   # tbRl: text reads top-to-bottom
    3: 90.0,    # lrTbV
    4: -90.0,   # tbRlV
    5: 90.0,    # tbLrV
}


@dataclass
class _CellLine:
    height: float
    segments: list = field(default_factory=list)
    image: Optional[ldm.Shape] = None
    width: float = 0.0
    align: str = "L"
    nested: Optional[tuple] = None
    rotated_text: Optional[str] = None
    paragraph: Optional[ldm.Paragraph] = None


class TableRenderer:
    """Renders LDM tables into PDF."""

    def __init__(self, writer: PDFWriterContext) -> None:
        self._writer = writer

    def render_table(self, pdf: FPDF, table: ldm.Table) -> None:
        """Keep rich cell content in a grid and split long rows at content-line boundaries."""
        if not table.rows:
            return
        w = self._writer
        attrs = table._tblp_pr_attrs
        if (attrs and table.text_wrapping == 1
                and attrs.get("tblpXSpec") in ("right", "center", "inside", "outside")):
            self._render_floating_table(pdf, table)
            return
        num_cols = max(sum(cell.cell_format.grid_span for cell in row.cells) for row in table.rows)
        if not num_cols:
            return
        usable_w = w._page_width - w._page_margin_left - w._page_margin_right - max(0, table.left_indent * PT_TO_MM)
        widths = self._compute_col_widths(table, num_cols, usable_w)
        layouts = [self._layout_row(pdf, row, widths) for row in table.rows]
        for row, layout in zip(table.rows, layouts):
            if (row.row_format.height_rule == 1 and row.row_format.height > 0
                    and self._row_height(None, layout) > row.row_format.height * PT_TO_MM + 1e-7):
                warn("Exact table row height is expanded to avoid clipping content",
                     PdfContentLossWarning, code="pdf.table_row_height")
        headers = []
        for i, row in enumerate(table.rows):
            if not row.row_format.heading_format:
                break
            headers.append(i)
        header_height = sum(self._row_height(table.rows[i], layouts[i]) for i in headers)
        in_hf = getattr(pdf, "_in_header_render", False) or getattr(pdf, "_in_footer_render", False)
        bottom = pdf.h - pdf.b_margin
        full_height = bottom - pdf.t_margin
        if not in_hf and header_height >= full_height:
            raise ValueError("Table headers do not fit the printable page area")
        pdf.ln(table.top_padding * PT_TO_MM)

        def table_x():
            free = usable_w - sum(widths)
            offset = free / 2 if table.alignment == ParagraphAlignment.CENTER else free if table.alignment == ParagraphAlignment.RIGHT else 0
            return w._page_margin_left + max(0.0, offset) + table.left_indent * PT_TO_MM

        def advance(repeat_headers=True):
            pdf._perform_page_break_if_need_be(pdf.h - pdf.get_y() + 1)
            if sum(widths) > w._page_width - w._page_margin_left - w._page_margin_right + 1e-7:
                raise ValueError("Table cannot flow into a narrower column")
            if repeat_headers:
                for index in headers:
                    height = self._row_height(table.rows[index], layouts[index])
                    self._paint_row(pdf, table, index, layouts[index], table_x(), pdf.get_y(), height)
                    pdf.set_y(pdf.get_y() + height)

        for row_index, row in enumerate(table.rows):
            remaining = [(cell, col, span, width, list(lines))
                         for cell, col, span, width, lines in layouts[row_index]]
            first_fragment = True
            while True:
                height = self._row_height(row, remaining) if first_fragment else self._row_height(None, remaining)
                available = bottom - pdf.get_y()
                capacity = full_height - (header_height if row_index not in headers else 0)
                if in_hf or height <= available + 1e-7:
                    self._paint_row(pdf, table, row_index, remaining, table_x(), pdf.get_y(), height)
                    pdf.set_y(pdf.get_y() + height)
                    break
                if height <= capacity + 1e-7:
                    advance(row_index not in headers)
                    continue
                if row_index in headers:
                    raise ValueError("Table header row cannot be split across pages")
                if first_fragment and row.row_format.height * PT_TO_MM > capacity:
                    warn("Table row minimum height exceeds a page; content is laid out without that minimum",
                         PdfContentLossWarning, code="pdf.table_row_height")
                if first_fragment and not row.row_format.allow_break_across_pages:
                    warn("A table row taller than a page must be split despite cantSplit",
                         PdfContentLossWarning, code="pdf.table_row_split")
                fragment, rest, taken = [], [], 0
                for cell, col, span, width, lines in remaining:
                    top, right, bottom_pad, left = self._padding(cell)
                    room = available - top - bottom_pad
                    selected, used = [], 0.0
                    for line in lines:
                        if line.image and line.height > capacity - top - bottom_pad:
                            scale = (capacity - top - bottom_pad) / line.height
                            if scale <= 0:
                                raise ValueError("No usable image height in table cell")
                            line = replace(line, height=line.height * scale, width=line.width * scale)
                        if used + line.height > room + 1e-7:
                            break
                        selected.append(line)
                        used += line.height
                    taken += len(selected)
                    fragment.append((cell, col, span, width, selected))
                    rest.append((cell, col, span, width, lines[len(selected):]))
                if not taken:
                    if available >= capacity - 1e-7:
                        raise ValueError("A table cell content line is taller than the printable page area")
                    advance()
                    continue
                height = self._row_height(None, fragment)
                self._paint_row(pdf, table, row_index, fragment, table_x(), pdf.get_y(), height)
                pdf.set_y(pdf.get_y() + height)
                remaining = rest
                if not any(lines for _, _, _, _, lines in remaining):
                    break
                first_fragment = False
                advance()
        pdf.set_x(w._page_margin_left)
        pdf.ln(table.bottom_padding * PT_TO_MM or POST_TABLE_SPACING_MM)

    @staticmethod
    def _padding(cell):
        cf = cell.cell_format
        return (cf.top_padding * PT_TO_MM or DEFAULT_CELL_PAD_TOP_MM,
                cf.right_padding * PT_TO_MM or DEFAULT_CELL_PAD_LEFT_MM,
                cf.bottom_padding * PT_TO_MM or DEFAULT_CELL_PAD_TOP_MM,
                cf.left_padding * PT_TO_MM or DEFAULT_CELL_PAD_LEFT_MM)

    @staticmethod
    def _row_cells(row):
        return ldm.iter_grid_cells(row)

    def _layout_row(self, pdf, row, widths):
        layout = []
        for cell, col, span in self._row_cells(row):
            width = sum(widths[col:col + span])
            top, right, bottom, left = self._padding(cell)
            inner_width = width - left - right
            if inner_width <= 0:
                raise ValueError("No usable text width in table cell")
            lines = self._cell_lines(pdf, cell, inner_width)
            layout.append((cell, col, span, width, lines))
        return layout

    def _cell_lines(self, pdf, cell, width):
        lines = []
        w = self._writer
        if cell.cell_format.orientation in _VERTICAL_ORIENTATIONS:
            if cell.tables or any(isinstance(item, ldm.Shape) for para in cell.paragraphs for item in para._children):
                raise ValueError("Rotated table cells with shapes or nested tables are unsupported")
            text = safe_text(cell_text(cell))
            height = max(self._measure_text_width(pdf, text, cell), DEFAULT_FONT_SIZE_PT * MIN_ROW_HEIGHT_FACTOR)
            return [_CellLine(height, rotated_text=text)]
        for child in cell.children:
            if isinstance(child, ldm.Table):
                count = max((sum(c.cell_format.grid_span for c in row.cells)
                             for row in child.rows), default=0)
                if count:
                    widths = self._compute_col_widths(child, count, width)
                    for i, row in enumerate(child.rows):
                        layout = self._layout_row(pdf, row, widths)
                        lines.append(_CellLine(self._row_height(row, layout), nested=(child, i, layout)))
                continue
            para = child
            first_line = len(lines)
            segments = []
            align = FPDF_ALIGN.get(para.paragraph_format.alignment, "L")

            def flush(segments=segments, para=para, align=align):
                if not segments:
                    return
                rows = w._run_renderer.wrap_segments(pdf, segments, width) if cell.cell_format.wrap_text else [list(segments)]
                for row in rows:
                    size = max((item[3] for item in row), default=DEFAULT_FONT_SIZE_PT)
                    height = w._paragraph_renderer.line_height_mm(size, para.paragraph_format)
                    lines.append(_CellLine(height, segments=row, align=align))
                segments.clear()

            if para.paragraph_format.space_before:
                lines.append(_CellLine(para.paragraph_format.space_before * PT_TO_MM))
            visible = {id(run) for run in visible_runs(para)}
            for item in para._children:
                if isinstance(item, ldm.Run) and id(item) in visible and not item.font.hidden:
                    apply_run_font(pdf, item.font)
                    text = w._run_renderer._resolve_run_text(pdf, item.text).replace("\t", " ").replace("\f", "")
                    for chunk, link in extract_link_segments(apply_caps(text, item.font)):
                        chunk = safe_text(chunk)
                        if chunk:
                            segments.append((item, chunk, pdf.get_string_width(chunk), pdf.font_size_pt, link))
                elif isinstance(item, ldm.Shape):
                    flush()
                    if item.has_image and item.image_data and item.image_data.image_bytes:
                        image_w = (item.width or DEFAULT_SHAPE_DIM_PT) * PT_TO_MM
                        image_h = (item.height or DEFAULT_SHAPE_DIM_PT) * PT_TO_MM
                        if image_w <= 0 or image_h <= 0:
                            raise ValueError("Table image dimensions must be positive")
                        scale = min(1.0, width / image_w)
                        lines.append(_CellLine(image_h * scale, image=item, width=image_w * scale, align=align))
                    if item.text_box:
                        nested_cell = ldm.Cell(paragraphs=item.text_box.get("paragraphs", []))
                        lines.extend(self._cell_lines(pdf, nested_cell, width))
            flush()
            if not para._children:
                lines.append(_CellLine(DEFAULT_FONT_SIZE_PT * PT_TO_MM * LINE_HEIGHT_FACTOR))
            if para.paragraph_format.space_after:
                lines.append(_CellLine(para.paragraph_format.space_after * PT_TO_MM))
            if len(lines) > first_line:
                lines[first_line].paragraph = para
        reset_font(pdf)
        return lines

    def _row_height(self, row, layout):
        height = DEFAULT_FONT_SIZE_PT * MIN_ROW_HEIGHT_FACTOR
        for cell, _, _, _, lines in layout:
            top, _, bottom, _ = self._padding(cell)
            height = max(height, sum(line.height for line in lines) + top + bottom)
        return max(height, row.row_format.height * PT_TO_MM if row else 0)

    def _paint_row(self, pdf, table, row_index, layout, x, y, height):
        w = self._writer
        table_borders = self._table_level_borders(table)
        style_borders = self._inherited_table_borders(table)
        num_cols = max(sum(c.cell_format.grid_span for c in row.cells) for row in table.rows)
        previous_auto, previous_margin = pdf.auto_page_break, pdf.b_margin
        previous_cell_margin = pdf.c_margin
        pdf.set_auto_page_break(False)
        pdf.c_margin = 0
        try:
            for cell, col, span, width, lines in layout:
                top, right, bottom, left = self._padding(cell)
                cf = cell.cell_format
                bg_color = parse_color(cf.shading.background_pattern_color)
                if bg_color:
                    pdf.set_fill_color(*bg_color)
                    pdf.rect(x, y, width, height, "F")
                next_merged = False
                if row_index + 1 < len(table.rows):
                    next_merged = any(c == col and candidate.cell_format.vertical_merge == 2
                                      for candidate, c, _ in self._row_cells(table.rows[row_index + 1]))
                self._draw_cell_borders(pdf, cell, table_borders, style_borders,
                    row_index, col, len(table.rows), num_cols, x, y, width, height,
                    last_col=col + span - 1, suppress_top=cf.vertical_merge == 2,
                    suppress_bottom=next_merged and cf.vertical_merge in (1, 2))
                content_height = sum(line.height for line in lines)
                free = max(0, height - top - bottom - content_height)
                at_y = y + top + (free / 2 if cf.vertical_alignment == 1 else free if cf.vertical_alignment == 2 else 0)
                with w._tag(pdf, "/TH" if table.rows[row_index].row_format.heading_format else "/TD"):
                    with pdf.rect_clip(x + left, y + top, width - left - right, height - top - bottom):
                        # q/Q restores the PDF font, but fpdf2 can retain the clip's cached font flag.
                        pdf.current_font_is_set_on_page = False
                        for line in lines:
                            offset = max(0, width - left - right - (line.width if line.image else sum(s[2] for s in line.segments)))
                            at_x = x + left + (offset / 2 if line.align == "C" else offset if line.align == "R" else 0)
                            pdf.set_xy(at_x, at_y)
                            if line.paragraph:
                                register_bookmarks(pdf, line.paragraph, w._anchor_links)
                                w._paragraph_renderer._emit_bookmark_outlines(pdf, line.paragraph)
                                pf = line.paragraph.paragraph_format
                                if pf.is_heading and w.options.outline_options.create_outlines_for_headings_in_tables:
                                    w._paragraph_renderer._emit_heading_outline(pdf, cell_text(cell), pf.outline_level + 1)
                            if line.rotated_text is not None:
                                font = self._get_cell_first_font(cell)
                                if font:
                                    apply_run_font(pdf, font)
                                self._render_rotated_cell(pdf, line.rotated_text, x, y, width, height,
                                                          left, top, DEFAULT_FONT_SIZE_PT * DEFAULT_CELL_LINE_H_FACTOR,
                                                          cf.orientation)
                            elif line.image:
                                data = w._shape_renderer.compress_image_bytes(line.image.image_data.image_bytes)
                                pdf.image(BytesIO(data), x=at_x, y=at_y, w=line.width, h=line.height)
                            elif line.nested:
                                nested_table, nested_index, nested_layout = line.nested
                                self._paint_row(pdf, nested_table, nested_index, nested_layout, x + left, at_y, line.height)
                            elif line.segments:
                                w._run_renderer._render_segment_row(pdf, line.segments, line.height, DEFAULT_FONT_SIZE_PT, at_x=at_x)
                            at_y += line.height
                x += width
        finally:
            pdf.set_auto_page_break(previous_auto, previous_margin)
            pdf.c_margin = previous_cell_margin
            pdf.set_y(y)
            reset_font(pdf)
            pdf.current_font_is_set_on_page = False

    def _render_floating_table(self, pdf: FPDF, table: ldm.Table) -> None:
        """Draw a `w:tblpPr` floating table without disturbing the cursor.

        Honours ``tblpXSpec`` (``left``/``center``/``right``); ``tblpY``
        is added to the live cursor Y. The text cursor is restored so
        the column flow continues underneath instead of being pushed
        below the table.
        """
        w = self._writer
        saved_x, saved_y = pdf.get_x(), pdf.get_y()

        attrs = table._tblp_pr_attrs
        spec = attrs.get("tblpXSpec", "left")
        # Page-level margins (not the current column's): floating tables
        # position themselves against the page or section margin, not
        # the column edge.
        page_left = getattr(w, "_page_full_margin_left", w._page_margin_left)
        page_right = getattr(w, "_page_full_margin_right", w._page_margin_right)
        page_usable_w = w._page_width - page_left - page_right
        num_cols = max(len(row.cells) for row in table.rows)
        col_widths = self._compute_col_widths(table, num_cols, page_usable_w)
        total_w = sum(col_widths)

        if spec == "right":
            float_x = w._page_width - page_right - total_w
        elif spec == "center":
            float_x = page_left + (page_usable_w - total_w) / 2
        else:
            float_x = saved_x

        tblpY = attrs.get("tblpY")
        y_offset_mm = 0.0
        if tblpY:
            try:
                y_offset_mm = int(tblpY) / 20.0 * PT_TO_MM  # twips → pt → mm
            except ValueError:
                pass
        # When the table pins to a side column, take the anchor Y from
        # that column's live cursor — not the writer cursor, which may
        # have hopped to a different column during section reflow.
        col_y_map = getattr(w, "_active_col_y", None)
        anchor_y = saved_y
        if spec == "right" and col_y_map:
            anchor_y = col_y_map.get(max(col_y_map), saved_y)
        float_y = anchor_y + y_offset_mm

        prev_left = w._page_margin_left
        prev_right = w._page_margin_right
        pdf.set_left_margin(float_x)
        pdf.set_right_margin(max(0.0, w._page_width - float_x - total_w))
        w._page_margin_left = float_x
        w._page_margin_right = max(0.0, w._page_width - float_x - total_w)
        pdf.set_xy(float_x, float_y)

        attrs_holder = table._tblp_pr_attrs
        table_bottom = saved_y
        table._tblp_pr_attrs = {}  # avoid recursion
        try:
            self.render_table(pdf, table)
            table_bottom = pdf.get_y()
        finally:
            table._tblp_pr_attrs = attrs_holder
            pdf.set_left_margin(prev_left)
            pdf.set_right_margin(prev_right)
            w._page_margin_left = prev_left
            w._page_margin_right = prev_right
            # If the inline cursor is already in the same column the
            # table just landed in, skip past it so the next paragraph
            # doesn't render on top of the table. (Cursor was always
            # in another column when reflow had moved us elsewhere.)
            if pdf.l_margin == w._page_margin_left and saved_y < table_bottom:
                pdf.set_xy(saved_x, table_bottom)
            else:
                pdf.set_xy(saved_x, saved_y)
        # Push the float's target column cursor past the table so the
        # next paragraphs in that column don't render on top of it.
        if col_y_map is not None and spec == "right":
            tgt_col = max(col_y_map)
            col_y_map[tgt_col] = max(col_y_map.get(tgt_col, 0.0), table_bottom)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _compute_col_widths(table: ldm.Table, num_cols: int, usable_w: float) -> list[float]:
        """Distribute spanning-cell widths across their actual grid columns."""
        if num_cols > MAX_TABLE_COLUMNS:
            raise ValueError("Table has too many grid columns")
        if usable_w <= 0:
            raise ValueError("No usable table width")
        preferred = table.preferred_width
        if preferred.type == 1 and preferred.value > 0:
            usable_w *= min(preferred.value, 100) / 100
        elif preferred.type == 2 and preferred.value > 0:
            usable_w = min(usable_w, preferred.value * PT_TO_MM)
        widths = [0.0] * num_cols
        for row in table.rows:
            for cell, col, span in TableRenderer._row_cells(row):
                if cell.cell_format.width > 0:
                    each = cell.cell_format.width * PT_TO_MM / span
                    for index in range(col, min(col + span, num_cols)):
                        widths[index] = max(widths[index], each)
        total = sum(widths)
        if not total:
            return [usable_w / num_cols] * num_cols
        zero_count = widths.count(0)
        if zero_count:
            fill = max(usable_w - total, usable_w / num_cols * zero_count) / zero_count
            widths = [width or fill for width in widths]
        total = sum(widths)
        if total > usable_w:
            widths = [width * usable_w / total for width in widths]
        return widths

    def _compute_row_height(
        self, pdf: FPDF, row: ldm.Row, col_widths: list[float], num_cols: int
    ) -> float:
        return self._row_height(row, self._layout_row(pdf, row, col_widths))

    @staticmethod
    def _render_rotated_cell(
        pdf: FPDF,
        text: str,
        cell_x: float,
        row_y: float,
        cw: float,
        rh: float,
        pad_left: float,
        pad_top: float,
        line_h: float,
        orientation: int,
    ) -> None:
        """Render text rotated 90/-90 degrees inside a cell."""
        angle = _ORIENTATION_TO_ANGLE.get(orientation, 90.0)
        cx = cell_x + cw / 2
        cy = row_y + rh / 2
        text = safe_text(text)
        with pdf.rotation(angle, cx, cy):
            tw = pdf.get_string_width(text)
            pdf.set_xy(cx - tw / 2, cy - line_h / 2)
            pdf.cell(w=tw, h=line_h, text=text)

    @staticmethod
    def _get_cell_first_font(cell: ldm.Cell) -> Optional[ldm.Font]:
        """Get the font from the first non-empty run in a cell."""
        for para in cell.paragraphs:
            for run in visible_runs(para):
                if run.text:
                    return run.font
        return None

    def _measure_text_width(
        self, pdf: FPDF, text: str, cell: ldm.Cell
    ) -> float:
        """Width of *text* in mm under the cell's first font."""
        if not text:
            return 0.0
        # Strip characters the core PDF font can't render so fpdf2's
        # encoder doesn't raise mid-measurement.
        text = safe_text(text)
        font = self._get_cell_first_font(cell)
        prev_family = pdf.font_family
        prev_style = pdf.font_style
        prev_size = pdf.font_size_pt
        try:
            if font:
                apply_run_font(pdf, font)
            else:
                reset_font(pdf)
            return float(pdf.get_string_width(text))
        finally:
            pdf.set_font(prev_family, prev_style, prev_size)

    @staticmethod
    def _cell_has_borders(cell: ldm.Cell) -> bool:
        """Kept for the writer\'s backward-compatible helper alias."""
        return any(border.line_style != LineStyle.NONE or border.line_width > 0
                   for border in cell.cell_format.borders or ())

    @staticmethod
    def _border_is_visible(border: Optional[ldm.Border]) -> bool:
        if border is None:
            return False
        return border.line_style != LineStyle.NONE or border.line_width > 0

    def _inherited_table_borders(
        self, table: ldm.Table
    ) -> Optional[list[ldm.Border]]:
        """Return borders from the table's style chain, if any."""
        doc = getattr(self._writer, "_doc", None)
        if doc is None or not table.style_name:
            return None
        seen: set[str] = set()
        name = table.style_name
        while name and name not in seen:
            seen.add(name)
            style = doc.find_style(name)
            if style is None:
                return None
            fmt = style.table_style_format
            if fmt and fmt.borders:
                return list(fmt.borders)
            name = style.base_style_name
        return None

    @staticmethod
    def _table_level_borders(table: ldm.Table) -> Optional[list[ldm.Border]]:
        """Return the table-level w:tblBorders (carried per-row by the reader)."""
        for row in table.rows:
            if row.row_format.borders:
                return list(row.row_format.borders)
        return None

    def _resolve_cell_side(
        self,
        cell_borders: list[ldm.Border],
        table_level_borders: Optional[list[ldm.Border]],
        style_borders: Optional[list[ldm.Border]],
        outer_slot: int,
        inside_slot: int,
        is_outer: bool,
    ) -> Optional[ldm.Border]:
        """Resolve a single cell side: explicit cell > table-level > style."""
        cb = cell_borders[outer_slot] if outer_slot < len(cell_borders) else None
        if self._border_is_visible(cb):
            return cb
        slot = outer_slot if is_outer else inside_slot
        for source in (table_level_borders, style_borders):
            if source is None:
                continue
            if slot >= len(source):
                continue
            # An explicitly-set ``none`` entry in this layer is a deliberate
            # override that should hide the side rather than letting the
            # next layer leak through.
            return source[slot] if self._border_is_visible(source[slot]) else None
        return None

    def _draw_cell_borders(
        self,
        pdf: FPDF,
        cell: ldm.Cell,
        table_level_borders: Optional[list[ldm.Border]],
        style_borders: Optional[list[ldm.Border]],
        row_idx: int,
        col_idx: int,
        num_rows: int,
        num_cols: int,
        x: float,
        y: float,
        w: float,
        h: float,
        *,
        last_col: Optional[int] = None,
        suppress_top: bool = False,
        suppress_bottom: bool = False,
    ) -> None:
        """Draw cell sides using cell → table-level → style cascade."""
        cell_borders = cell.cell_format.borders or []

        # Interior edges fall back to insideH / insideV; outer edges use
        # the outer slot. The outer-slot border can still hide an interior
        # edge if a cell explicitly sets its own border.
        top = self._resolve_cell_side(
            cell_borders, table_level_borders, style_borders,
            _B_TOP, _B_INSIDE_H, row_idx == 0,
        )
        bottom = self._resolve_cell_side(
            cell_borders, table_level_borders, style_borders,
            _B_BOTTOM, _B_INSIDE_H, row_idx == num_rows - 1,
        )
        left = self._resolve_cell_side(
            cell_borders, table_level_borders, style_borders,
            _B_LEFT, _B_INSIDE_V, col_idx == 0,
        )
        right = self._resolve_cell_side(
            cell_borders, table_level_borders, style_borders,
            _B_RIGHT, _B_INSIDE_V, (last_col if last_col is not None else col_idx) == num_cols - 1,
        )

        if suppress_top:
            top = None
        if suppress_bottom:
            bottom = None
        if not any((top, left, bottom, right)):
            return

        prev_lw = pdf.line_width

        def draw_line(b: ldm.Border, x1: float, y1: float, x2: float, y2: float) -> None:
            rgb = parse_color(b.color) or (0, 0, 0)
            pdf.set_draw_color(*rgb)
            lw_mm = max(b.line_width * PT_TO_MM, MIN_LINE_WIDTH_MM)
            pdf.set_line_width(lw_mm)
            pdf.line(x1, y1, x2, y2)

        if top:
            draw_line(top, x, y, x + w, y)
        if bottom:
            draw_line(bottom, x, y + h, x + w, y + h)
        if left:
            draw_line(left, x, y, x, y + h)
        if right:
            draw_line(right, x + w, y, x + w, y + h)

        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(prev_lw)
