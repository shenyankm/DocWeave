"""Build :class:`ldm.Paragraph` (with runs, hyperlinks, fields, drawings).

The paragraph is the most complex element in the LDM: it stitches
together inline runs, bookmarks, field-code state, embedded
drawings and TOC-aware hyperlink splitting.  Each concern lives in
its own small method here so the top-level :meth:`ParagraphBuilder.build`
reads like a flat dispatch loop.
"""

import re
from typing import Optional
from xml.etree import ElementTree as ET

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._links import format_link
from aspose.words_foss._visible_runs import HORIZONTAL_RULE_SHAPE_TYPE
from aspose.words_foss.docx_reader.constants import (
    PAGE_FIELD_SENTINEL,
    R_NS,
    W_NS,
)
from aspose.words_foss.docx_reader.utils import _collect_run_text

from ._context import ReaderContext
from ._helpers import is_truthy_onoff
from .cascading import (
    FontResolver,
    ParagraphFormatResolver,
    StyleChainResolver,
)


_HEADING_RE = re.compile(r"[Hh]eading\s*(\d+)")
_PAGE_FIELD_RE = re.compile(r"\s*PAGE(\s|$)")

_VML_RECT_TAG = "{urn:schemas-microsoft-com:vml}rect"
_O_HR_ATTR = "{urn:schemas-microsoft-com:office:office}hr"
_VML_LENGTH_RE = re.compile(r"(-?[\d.]+)([a-z]*)")
_VML_UNIT_TO_PT = {"": 1.0, "pt": 1.0, "in": 72.0, "cm": 72.0 / 2.54, "mm": 72.0 / 25.4, "px": 0.75}


def _vml_style_lengths(style: str) -> dict[str, float]:
    out: dict[str, float] = {}
    for declaration in style.split(";"):
        prop, _, raw = declaration.partition(":")
        match = _VML_LENGTH_RE.match(raw.strip())
        if match:
            unit = _VML_UNIT_TO_PT.get(match.group(2), 1.0)
            out[prop.strip()] = float(match.group(1)) * unit
    return out


def _build_horizontal_rule_shape(rect: ET.Element) -> ldm.Shape:
    lengths = _vml_style_lengths(rect.get("style") or "")
    return ldm.Shape(
        shape_type=HORIZONTAL_RULE_SHAPE_TYPE,
        width=lengths.get("width", 432.0),
        height=lengths.get("height", 1.5),
        is_inline=True,
        fill_color="808080",
    )


class RunBuilder:
    """Build a single :class:`ldm.Run` with a resolved font cascade."""

    def __init__(self, font_resolver: FontResolver):
        self._font_resolver = font_resolver

    def build(self, r_elem: ET.Element, para_style_id: str = "",
              implicit_paragraph_style: bool = False) -> ldm.Run:
        """Translate one ``<w:r>`` into a Run; *para_style_id* feeds the font cascade."""
        run = ldm.Run()
        run.text = _collect_run_text(r_elem)
        run.font = self._font_resolver.resolve(r_elem.find(f"{W_NS}rPr"), para_style_id, implicit_paragraph_style)
        return run


class _FieldTracker:
    """Track ``<w:fldChar>``/``<w:instrText>`` state across a paragraph.

    Mirrors Aspose's view of a field: ``<w:r>`` elements between ``begin``
    and ``separate`` carry the field code as regular ``Run`` nodes whose
    text comes from ``<w:instrText>`` rather than ``<w:t>``; the writer
    re-emits the correct element based on the run's position relative to
    the markers, so this tracker only needs to know whether we're still
    in the field-code window (for PAGE-sentinel detection)."""

    __slots__ = ("depth", "page_pending", "suppress_cached", "in_code")

    def __init__(self) -> None:
        self.depth: int = 0
        self.page_pending: bool = False
        self.suppress_cached: bool = False
        # True while the current innermost field is between begin and
        # separate (i.e. emitting field-code runs).
        self.in_code: bool = False

    def on_begin(self) -> None:
        self.depth += 1
        self.page_pending = False
        self.in_code = True

    def on_separate(self) -> bool:
        """Return True if the *current* field is the PAGE sentinel."""
        self.in_code = False
        if not self.page_pending:
            return False
        self.suppress_cached = True
        self.page_pending = False
        return True

    def on_end(self) -> None:
        self.depth = max(0, self.depth - 1)
        self.suppress_cached = False
        self.page_pending = False
        self.in_code = False

    def note_instr_text(self, text: str) -> None:
        if self.depth > 0 and _PAGE_FIELD_RE.match(text):
            self.page_pending = True


class ParagraphBuilder:
    """Build :class:`ldm.Paragraph` from a ``<w:p>`` element."""

    def __init__(
        self,
        ctx: ReaderContext,
        font_resolver: FontResolver,
        pf_resolver: ParagraphFormatResolver,
        chains: StyleChainResolver,
        run_builder: RunBuilder,
    ):
        self._ctx = ctx
        self._font_resolver = font_resolver
        self._pf_resolver = pf_resolver
        self._chains = chains
        self._run_builder = run_builder
        self._default_style_id = next((sid for sid, style in reversed(ctx._style_elem_cache.items())
            if style.get(f"{W_NS}type") == "paragraph"
            and is_truthy_onoff(style.get(f"{W_NS}default", ""))), "")

    def build(
        self,
        p_elem: ET.Element,
        image_rels: Optional[dict[str, str]] = None,
    ) -> ldm.Paragraph:
        """Translate one ``<w:p>`` into a Paragraph.

        *image_rels* maps ``rId`` → media path for embedded drawings;
        defaults to the document-level relationships on the reader.
        """
        if image_rels is None:
            image_rels = self._ctx._doc_image_rels

        para = ldm.Paragraph(source_location=self._ctx._source_locations.get(p_elem))
        pPr = p_elem.find(f"{W_NS}pPr")
        explicit_style_id = self._read_style_id(pPr)
        implicit_paragraph_style = not explicit_style_id
        para_style_id = explicit_style_id or self._default_style_id

        para.paragraph_format = self._pf_resolver.resolve(pPr, para_style_id)
        self._apply_paragraph_meta(para.paragraph_format, para_style_id)
        para.paragraph_format.style_explicit = bool(explicit_style_id)
        self._apply_list_format(para, pPr, para_style_id)

        tracker = _FieldTracker()
        for child in self._ctx._resolve_paragraph_children(p_elem):
            tag = child.tag
            if tag == f"{W_NS}bookmarkStart":
                self._handle_bookmark_start(child, para)
            elif tag == f"{W_NS}bookmarkEnd":
                self._handle_bookmark_end(child, para)
            elif tag == f"{W_NS}r":
                self._handle_run(child, para, tracker, para_style_id, image_rels, implicit_paragraph_style)
            elif tag == f"{W_NS}hyperlink":
                self._handle_hyperlink(child, para, para_style_id, tracker, image_rels, implicit_paragraph_style)
        if not para._children:
            mark = self._font_resolver.resolve(None, para_style_id, implicit_paragraph_style)
            if para.paragraph_break_font is not None:
                self._font_resolver.merge(mark, para.paragraph_break_font)
            para.paragraph_break_font = mark
        return para

    @staticmethod
    def _read_style_id(pPr: Optional[ET.Element]) -> str:
        if pPr is None:
            return ""
        pStyle = pPr.find(f"{W_NS}pStyle")
        return pStyle.get(f"{W_NS}val", "") if pStyle is not None else ""

    def _apply_paragraph_meta(
        self,
        pf: ldm.ParagraphFormat,
        para_style_id: str,
    ) -> None:
        if para_style_id:
            pf.style_name = self._ctx._resolve_style_name(para_style_id)
        if pf.is_heading or not pf.style_name:
            return
        match = _HEADING_RE.search(pf.style_name)
        if match:
            pf.is_heading = True
            pf.outline_level = int(match.group(1)) - 1

    def _apply_list_format(
        self,
        para: ldm.Paragraph,
        pPr: Optional[ET.Element],
        para_style_id: str,
    ) -> None:
        direct_numPr = pPr.find(f"{W_NS}numPr") if pPr is not None else None
        numPr = direct_numPr
        if numPr is None and para_style_id:
            numPr = self._chains.numPr_from_style(para_style_id)
        if numPr is None:
            return
        numId_elem = numPr.find(f"{W_NS}numId")
        num_id_val = int(numId_elem.get(f"{W_NS}val", "0")) if numId_elem is not None else 0
        if num_id_val <= 0:
            return
        lf = ldm.ListFormat()
        lf.is_list_item = True
        lf.list_id = num_id_val
        ilvl = numPr.find(f"{W_NS}ilvl")
        if ilvl is not None:
            lf.list_level_number = int(ilvl.get(f"{W_NS}val", "0"))
        para.list_format = lf
        para.paragraph_format.is_list_item = True

    def _handle_bookmark_start(self, child: ET.Element, para: ldm.Paragraph) -> None:
        name = child.get(f"{W_NS}name", "")
        if not name or name.startswith("_GoBack"):
            return
        bm_id = child.get(f"{W_NS}id", "")
        if bm_id:
            self._ctx._bookmark_id_to_name[bm_id] = name
        para._children.append(ldm.BookmarkStart(name=name))

    def _handle_bookmark_end(self, child: ET.Element, para: ldm.Paragraph) -> None:
        bm_id = child.get(f"{W_NS}id", "")
        bm_name = self._ctx._bookmark_id_to_name.get(bm_id, "")
        if bm_name:
            para._children.append(ldm.BookmarkEnd(name=bm_name))

    def _handle_run(
        self,
        child: ET.Element,
        para: ldm.Paragraph,
        tracker: _FieldTracker,
        para_style_id: str,
        image_rels: dict[str, str],
        implicit_paragraph_style: bool = False,
    ) -> None:
        fld_char = child.find(f"{W_NS}fldChar")
        if fld_char is not None:
            self._handle_field_char(fld_char, child, para, tracker, para_style_id, implicit_paragraph_style)
            return
        instr = child.find(f"{W_NS}instrText")
        if instr is not None:
            tracker.note_instr_text(instr.text or "")
            if tracker.page_pending:
                return
        if tracker.suppress_cached:
            return
        self._extract_drawings(child, para, image_rels)
        if any(item.tag in {W_NS + "footnoteReference", W_NS + "endnoteReference"} for item in child):
            source_run = self._run_builder.build(child, para_style_id, implicit_paragraph_style)
            fragment = ET.Element(W_NS + "r")
            def flush():
                text = _collect_run_text(fragment)
                if text:
                    run = source_run.model_copy(deep=True)
                    run.text = text
                    para._children.append(run)
                fragment.clear()
            for item in child:
                if item.tag in {W_NS + "footnoteReference", W_NS + "endnoteReference"}:
                    flush()
                    identifier = item.get(W_NS + "id", "")
                    if identifier:
                        anchor = ldm.NoteReference(kind=item.tag[len(W_NS):-9], identifier=identifier,
                                                   hidden=source_run.font.hidden)
                        para.note_references.append(anchor)
                        para._children.append(anchor)
                else:
                    fragment.append(item)
            flush()
            return
        run = self._run_builder.build(child, para_style_id, implicit_paragraph_style)
        if not run.text:
            return
        para._children.append(run)

    def _handle_field_char(
        self,
        fld_char: ET.Element,
        run_elem: ET.Element,
        para: ldm.Paragraph,
        tracker: _FieldTracker,
        para_style_id: str,
        implicit_paragraph_style: bool = False,
    ) -> None:
        ft = fld_char.get(f"{W_NS}fldCharType", "")
        if ft == "begin":
            tracker.on_begin()
            para._children.append(ldm.FieldStart())
            return
        if ft == "separate":
            is_page = tracker.on_separate()
            para._children.append(ldm.FieldSeparator())
            if is_page:
                sentinel = self._build_page_sentinel(run_elem, para_style_id, implicit_paragraph_style)
                para._children.append(sentinel)
            return
        if ft == "end":
            tracker.on_end()
            para._children.append(ldm.FieldEnd())

    def _build_page_sentinel(self, run_elem: ET.Element, para_style_id: str,
                             implicit_paragraph_style: bool = False) -> ldm.Run:
        sentinel = ldm.Run()
        sentinel.text = PAGE_FIELD_SENTINEL
        sentinel.font = self._font_resolver.resolve(
            run_elem.find(f"{W_NS}rPr"), para_style_id, implicit_paragraph_style
        )
        return sentinel

    def _extract_drawings(
        self,
        run_elem: ET.Element,
        para: ldm.Paragraph,
        image_rels: dict[str, str],
    ) -> None:
        for drawing in self._ctx._iter_effective_drawings(run_elem):
            positioned = self._ctx._extract_positioned_shapes(drawing, image_rels)
            if positioned:
                para._children.extend(positioned)
                continue
            shape = self._ctx._build_drawing_shape(drawing, image_rels)
            if shape is not None:
                para._children.append(shape)
        for pict in run_elem.findall(f"{W_NS}pict"):
            rect = pict.find(_VML_RECT_TAG)
            if rect is not None and rect.get(_O_HR_ATTR) in ("t", "true", "1"):
                para._children.append(_build_horizontal_rule_shape(rect))

    def _handle_hyperlink(
        self,
        child: ET.Element,
        para: ldm.Paragraph,
        para_style_id: str,
        tracker: _FieldTracker,
        image_rels: dict[str, str],
        implicit_paragraph_style: bool = False,
    ) -> None:
        url = self._resolve_hyperlink_url(child)
        head_runs, tail_runs = self._split_at_first_tab(child)

        if head_runs:
            if any(r.find(W_NS + kind + "Reference") is not None
                   for r in head_runs for kind in ("footnote", "endnote")):
                for r in head_runs:
                    start = len(para._children)
                    self._handle_run(r, para, tracker, para_style_id, image_rels, implicit_paragraph_style)
                    for run in para._children[start:]:
                        if isinstance(run, ldm.Run) and url:
                            run.text = format_link(run.text, url)
                            run.is_hyperlink = True
                head_runs = []
        if head_runs:
            link_text = "".join(_collect_run_text(r) for r in head_runs)
            if link_text:
                head_rPr = head_runs[0].find(f"{W_NS}rPr")
                self._append_link_run(para, link_text, url, head_rPr, para_style_id, implicit_paragraph_style)

        for r_elem in tail_runs:
            self._handle_run(r_elem, para, tracker, para_style_id, image_rels, implicit_paragraph_style)

    def _resolve_hyperlink_url(self, child: ET.Element) -> str:
        r_id = child.get(f"{R_NS}id", "")
        url = self._ctx._rels.get(r_id, "")
        anchor = child.get(f"{W_NS}anchor", "")
        if anchor:
            return f"{url}#{anchor}" if url else f"#{anchor}"
        return url

    @staticmethod
    def _split_at_first_tab(
        child: ET.Element,
    ) -> tuple[list[ET.Element], list[ET.Element]]:
        runs = list(child.findall(f"{W_NS}r"))
        for i, r_elem in enumerate(runs):
            if r_elem.find(f"{W_NS}tab") is not None:
                return runs[:i], runs[i:]
        return runs, []

    def _append_link_run(
        self,
        para: ldm.Paragraph,
        link_text: str,
        url: str,
        head_rPr: Optional[ET.Element],
        para_style_id: str,
        implicit_paragraph_style: bool = False,
    ) -> None:
        run = ldm.Run()
        run.text = format_link(link_text, url) if url else link_text
        run.is_hyperlink = bool(url)
        run.font = self._font_resolver.resolve(head_rPr, para_style_id, implicit_paragraph_style)
        para._children.append(run)

    @staticmethod
    def paragraph_ends_first_page(para: ldm.Paragraph) -> bool:
        """True when *para* closes Word's first body page (form-feed"""
        for run in para.runs:
            if "\f" in (run.text or ""):
                return True
        return para.paragraph_format.page_break_before
