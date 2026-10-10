"""Text utility functions for the PDF writer."""


from typing import Optional, Tuple

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._visible_runs import visible_children, visible_runs
from aspose.words_foss.pdf_writer.constants import DEFAULT_FONT_SIZE_PT, LINE_HEIGHT_FACTOR
from aspose.words_foss._links import INLINE_LINK_RE, decode_link
from aspose.words_foss.pdf_writer.color import parse_color

def safe_text(text: str) -> str:
    """Preserve Unicode, mapping Word symbol-font bullets to real codepoints."""
    return text.replace("\uf0b7", "\u2022").replace("\uf0a7", "\u25aa")


def extract_link_segments(text: str) -> list[Tuple[str, Optional[str]]]:
    """Split ``text`` into ``(chunk, url)`` segments around Markdown links."""
    if "[" not in text:
        return [(text, None)]
    segments: list[Tuple[str, Optional[str]]] = []
    idx = 0
    for m in INLINE_LINK_RE.finditer(text):
        if m.start() > idx:
            segments.append((text[idx : m.start()], None))
        display, url = decode_link(m)
        segments.append((display, url))
        idx = m.end()
    if idx < len(text):
        segments.append((text[idx:], None))
    return segments or [(text, None)]


def apply_caps(text: str, font: ldm.Font) -> str:
    """Uppercase *text* when the run requests ``w:caps`` / ``w:smallCaps``."""
    if text and (font.all_caps or font.small_caps):
        return text.upper()
    return text


def plain_text(para: ldm.Paragraph) -> str:
    """Return paragraph text with Markdown link syntax stripped and ``w:caps`` applied."""
    return "".join(
        apply_caps(chunk, run.font)
        for run in visible_runs(para) if not run.font.render_hidden
        for chunk, _ in extract_link_segments(run.text or "")
    )


def source_plain_text(para: ldm.Paragraph) -> str:
    """Visible source text without interpreting ordinary runs as Markdown links."""
    return "".join(apply_caps(chunk, run.font) for run in visible_runs(para) if not run.font.render_hidden
                   for chunk, _ in (extract_link_segments(run.text) if run.is_hyperlink
                                    else [(run.text, None)]))


def cell_text(cell: ldm.Cell) -> str:
    """Flatten a cell's paragraphs to plain text.

    Paragraphs are joined with newlines so ``multi_cell`` renders each
    one on its own line (important for layout tables that use a single
    cell to hold multi-line blocks like address blocks).  Markdown link
    syntax is stripped so the ``[text](url)`` form from the readers does
    not leak into the cell.  Per-run ``w:caps`` is honoured via
    :func:`apply_caps`; hidden runs and field instructions are excluded.
    """
    lines = (plain_text(para) for para in cell.paragraphs)
    return "\n".join(line for line in lines if line)


def is_pure_page_break(para: ldm.Paragraph) -> bool:
    """True when the paragraph contains only form-feeds and no images."""
    for item in visible_children(para):
        if isinstance(item, ldm.Shape) and item.has_image:
            return False
    visible = False
    saw_form_feed = False
    for run in visible_runs(para):
        if run.font.render_hidden:
            continue
        text = run.text or ""
        for ch in text:
            if ch == "\f":
                saw_form_feed = True
            elif not ch.isspace():
                visible = True
    return saw_form_feed and not visible


def is_toc_style(style_name: str) -> bool:
    """True when *style_name* names a Word Table-of-Contents style."""
    if not style_name:
        return False
    lowered = style_name.lower()
    return lowered.startswith("toc") or "toc " in lowered or "table of contents" in lowered


def get_dominant_font_size(runs: list[ldm.Run]) -> float:
    """Get the font size from the first run that has a non-zero size."""
    for run in runs:
        if run.font.size > 0:
            return run.font.size
    return 0.0


def get_line_font_size(runs: list[ldm.Run]) -> float:
    """Line advance must accommodate the largest visible run, not just the first."""
    # Native SUP (0.7 scale, 0.4 rise) needs 1.1 em above the common baseline.
    return max(((run.font.size if run.font.size > 0 else DEFAULT_FONT_SIZE_PT)
                * (1.6 / LINE_HEIGHT_FACTOR if run.font.superscript else 1) for run in runs),
               default=DEFAULT_FONT_SIZE_PT)


def get_dominant_color(runs: list[ldm.Run]) -> Optional[Tuple[int, int, int]]:
    """Get the color from the first run that has a parseable color."""

    for run in runs:
        rgb = parse_color(run.font.render_color)
        if rgb:
            return rgb
    return None
