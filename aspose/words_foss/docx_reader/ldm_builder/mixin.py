"""Facade mixin that wires the LDM builders into :class:`DocumentReader`.

The mixin owns one set of builder instances per document, instantiates
them lazily in :meth:`to_light_document`, and orchestrates the
top-level build in document order: styles → lists → sections →
header/footer parts.  No XML parsing happens here — every method
delegates to a specialised builder.
"""

from typing import Optional
from xml.etree import ElementTree as ET

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader.constants import (
    COLOR_EMPTY,
    W_NS,
    _DEFAULT_TAB_STOP_PT,
)
from aspose.words_foss.docx_reader.utils import (
    _canonicalize_style_name,
    _hex_to_ldm_color,
    parse_onoff,
)

from .cascading import (
    FontBuilder,
    FontResolver,
    ParagraphFormatBuilder,
    ParagraphFormatResolver,
    StyleChainResolver,
)
from .definitions import ListBuilder, StyleBuilder
from .paragraphs import ParagraphBuilder, RunBuilder
from .sections import PageSetupBuilder, SectionBuilder
from .tables import TableBuilder


class LdmBuilderMixin:
    """Mixin that adds :meth:`to_light_document` to :class:`DocumentReader`."""

    def to_light_document(self) -> ldm.Document:
        """Build a :class:`ldm.Document` from the loaded DOCX."""
        doc = ldm.Document()
        doc.source_theme = self._source_theme
        doc.source_font_table = self._source_font_table
        doc.page_color = self._get_page_color()
        doc.default_tab_stop = self._get_default_tab_stop()
        doc.compatibility_mode = 12
        if self._settings_xml is not None:
            flags = {}
            for field, tag in [('embed_true_type_fonts', 'embedTrueTypeFonts'),
                               ('do_not_embed_system_fonts', 'doNotEmbedSystemFonts'),
                               ('embed_system_fonts', 'embedSystemFonts'),
                               ('save_subset_fonts', 'saveSubsetFonts')]:
                flag = self._settings_xml.find(W_NS + tag)
                if flag is not None:
                    flags[field] = parse_onoff(flag)
            if flags:
                doc.font_embedding = ldm.FontEmbeddingSettings(**flags)
            shift_return_flags = self._settings_xml.findall(
                f"{W_NS}compat/{W_NS}doNotExpandShiftReturn")
            doc.do_not_expand_shift_return = parse_onoff(
                shift_return_flags[-1] if shift_return_flags else None)
            for setting in self._settings_xml.findall(f"{W_NS}compat/{W_NS}compatSetting"):
                if (setting.get(f"{W_NS}name") == "compatibilityMode" and
                        setting.get(f"{W_NS}uri") == "http://schemas.microsoft.com/office/word"):
                    value = setting.get(f"{W_NS}val", "")
                    if value.isascii() and value.isdecimal():
                        doc.compatibility_mode = int(value)
                    break
        compat = self._settings_xml.find(W_NS + 'compat') if self._settings_xml is not None else None
        ui_flags = self._settings_xml.findall(W_NS+'uiCompat97To2003') if self._settings_xml is not None else []
        doc.source_settings = ldm.SourceSettings(
            data=ET.tostring(compat) if compat is not None else None,
            compatibility_mode=doc.compatibility_mode,
            do_not_expand_shift_return=doc.do_not_expand_shift_return,
            ui_compat_97_to_2003=parse_onoff(ui_flags[-1]) if ui_flags else None,
            ui_compat_data=ET.tostring(ui_flags[-1]) if ui_flags else None)
        self._bookmark_id_to_name = {}

        self._build_name_to_style_id_map()
        self._ensure_builders_installed()

        # build_all() synthesizes a default Normal when styles.xml is
        # absent, matching the writer's emit-Normal-by-default behaviour.
        doc.styles = self._styles_builder.build_all()
        if self._numbering_xml is not None:
            doc.lists = self._lists_builder.build_all()
        doc.sections = self._sections_builder.build_all()
        even_flags = (self._settings_xml.findall(W_NS + "evenAndOddHeaders")
                      if self._settings_xml is not None else [])
        for section in doc.sections:
            section.page_setup.odd_and_even_pages_header_footer = parse_onoff(
                even_flags[-1] if even_flags else None)

        if self._doc_default_rPr is not None:
            doc.doc_defaults_font = self._fonts_builder.build(self._doc_default_rPr)
        doc.doc_defaults_rpr_present = (self._styles_xml is not None and
                                       self._styles_xml.find(f"{W_NS}docDefaults/{W_NS}rPrDefault") is not None)

        self._populate_source_stories(doc)
        self._populate_headers_footers(doc)
        from aspose.words_foss._comments import validate_comments, validate_source_anchors
        validate_comments(doc)
        validate_source_anchors(self, doc)
        from aspose.words_foss._font_sizes import bind_font_sizes
        bind_font_sizes(doc)
        return doc

    # -- Reader-side state helpers ---------------------------------------------

    def _get_default_tab_stop(self) -> float:
        if self._settings_xml is None:
            return _DEFAULT_TAB_STOP_PT
        elements = self._settings_xml.findall(f"{W_NS}defaultTabStop")
        if not elements:
            return _DEFAULT_TAB_STOP_PT
        from aspose.words_foss.utils.tab_stops import parse_default_tab_stop
        return parse_default_tab_stop(elements[-1].get(f"{W_NS}val", ""))

    def _get_page_color(self) -> str:
        if self._document_xml is None:
            return COLOR_EMPTY
        bg = self._document_xml.find(f"{W_NS}background")
        if bg is None:
            return COLOR_EMPTY
        color = bg.get(f"{W_NS}color", "")
        return _hex_to_ldm_color(color) if color else COLOR_EMPTY

    def _build_name_to_style_id_map(self) -> None:
        self._name_to_style_id = {}
        if self._styles_xml is None:
            return
        for style_elem in self._styles_xml.findall(f"{W_NS}style"):
            sid = style_elem.get(f"{W_NS}styleId", "")
            name_elem = style_elem.find(f"{W_NS}name")
            if name_elem is None or not sid:
                continue
            raw_name = name_elem.get(f"{W_NS}val", "")
            self._name_to_style_id[raw_name] = sid
            is_custom = style_elem.get(f"{W_NS}customStyle", "") == "1"
            if is_custom:
                continue
            canonical = _canonicalize_style_name(raw_name)
            if canonical != raw_name:
                self._name_to_style_id[canonical] = sid

    def _install_builders(self) -> None:
        """Instantiate per-document builders and wire their dependencies."""
        if not hasattr(self, "_bookmark_id_to_name"):
            self._bookmark_id_to_name = {}
        if not hasattr(self, "_current_table_style_id"):
            self._current_table_style_id = ""

        # The mixin is only ever used through DocumentReader, so `self`
        # carries the parser-side state the builders read from.
        reader = self
        chains = StyleChainResolver(reader)
        fonts = FontBuilder(reader)
        pfs = ParagraphFormatBuilder(reader, fonts)
        font_resolver = FontResolver(reader, fonts, chains)
        pf_resolver = ParagraphFormatResolver(reader, pfs, chains)
        run_builder = RunBuilder(font_resolver)
        paragraph_builder = ParagraphBuilder(
            reader, font_resolver, pf_resolver, chains, run_builder
        )
        table_builder = TableBuilder(reader, paragraph_builder.build)
        page_setup_builder = PageSetupBuilder()
        sections_builder = SectionBuilder(
            reader, paragraph_builder, table_builder.build, page_setup_builder
        )

        self._chains = chains
        self._fonts_builder = fonts
        self._pfs_builder = pfs
        self._font_resolver = font_resolver
        self._pf_resolver = pf_resolver
        self._paragraph_builder = paragraph_builder
        self._table_builder = table_builder
        self._page_setup_builder = page_setup_builder
        self._sections_builder = sections_builder
        self._styles_builder = StyleBuilder(
            reader, fonts, font_resolver, pfs, pf_resolver
        )
        self._lists_builder = ListBuilder(reader, fonts)

    def _build_ldm_paragraph(
        self,
        p_elem: ET.Element,
        image_rels: Optional[dict[str, str]] = None,
    ) -> ldm.Paragraph:
        """Compatibility shim for :class:`ShapeParserMixin`."""
        self._ensure_builders_installed()
        return self._paragraph_builder.build(p_elem, image_rels)

    def _ensure_builders_installed(self) -> None:
        if getattr(self, "_paragraph_builder", None) is None:
            self._install_builders()

    def _populate_source_stories(self, doc):
        stories = []
        for kind, name, identifier, root, references in self._source_story_data:
            if kind in {"header", "footer"} and references:
                # Section references determine the role, even for mismatched source parts.
                for role in ("header", "footer"):
                    refs = [ref for ref in references if ref.get("kind", kind) == role]
                    if refs:
                        stories.append((role, name, identifier, root, refs))
            else:
                stories.append((kind, name, identifier, root, references))
        for kind, name, identifier, root, references in stories:
            _, children = self._build_part_children(
                [(root, self._part_images[name])],
                anchor_mode=kind if kind in {"header", "footer"} else "body",
                hyperlink_rels=self._part_links[name],
            )
            doc.source_stories.append(ldm.SourceStory(
                kind=kind, part_name=name, identifier=identifier,
                references=references, children=children,
                comment_metadata=(ldm.CommentMetadata(author=root.get(W_NS + "author"),
                                  initial=root.get(W_NS + "initials"), date=root.get(W_NS + "date"),
                                  **self._comment_threads.get(identifier, {}))
                                  if kind == "comment" else None),
            ))

    def _populate_headers_footers(self, doc: ldm.Document) -> None:
        """Keep explicit section references; absent references inherit in OOXML."""
        variants = {"default": 0, "first": 2, "even": 4}
        for story in doc.source_stories:
            if story.kind not in {"header", "footer"}:
                continue
            for reference in story.references:
                if reference["inherited"]:
                    continue
                index = reference["section"]
                variant = variants.get(reference["variant"])
                if variant is None or not 0 <= index < len(doc.sections):
                    raise ValueError("Invalid header/footer section reference")
                kind = variant + (story.kind == "footer")
                doc.sections[index].headers_footers.append(ldm.HeaderFooter(
                    header_footer_type=kind, children=story.children))

    def _build_part_children(
        self,
        parts: list[tuple[ET.Element, dict[str, str]]],
        *,
        anchor_mode: str,
        hyperlink_rels=None,
    ) -> tuple[list[ldm.Paragraph], list[ldm.BodyChild]]:
        paragraphs: list[ldm.Paragraph] = []
        children: list[ldm.BodyChild] = []
        self._anchor_y_base_mode = anchor_mode  # type: ignore[attr-defined]
        previous_links, previous_images = self._rels, self._doc_image_rels
        try:
            if hyperlink_rels is not None:
                self._rels = hyperlink_rels
            for part_xml, part_rels in parts:
                self._doc_image_rels = part_rels
                for elem in self._resolve_body_children(part_xml):
                    if elem.tag == f"{W_NS}p":
                        p = self._paragraph_builder.build(elem, part_rels)
                        paragraphs.append(p)
                        children.append(p)
                    elif elem.tag == f"{W_NS}tbl":
                        children.append(self._table_builder.build(elem))
        finally:
            self._rels, self._doc_image_rels = previous_links, previous_images
            self._anchor_y_base_mode = "body"  # type: ignore[attr-defined]
        return paragraphs, children

    # -- Type-only declarations of host-provided state -------------------------
    # The mixin reads these from the host :class:`DocumentReader`; the
    # annotations here are for IDEs / type checkers and are never executed
    # (they shadow nothing because the host always sets the real values).
    # Keep this block in sync with :class:`ReaderContext` in ``_context``.
    _document_xml: Optional[ET.Element]
    _numbering_xml: Optional[ET.Element]
    _styles_xml: Optional[ET.Element]
    _settings_xml: Optional[ET.Element]
    _rels: dict[str, str]
    _doc_image_rels: dict[str, str]
    _style_elem_cache: dict[str, ET.Element]
    _style_id_to_name: dict[str, str]
    _name_to_style_id: dict[str, str]
    _theme_fonts: dict[str, str]
    _theme_colors: dict[str, str]
    _doc_default_rPr: Optional[ET.Element]
    _doc_default_pPr: Optional[ET.Element]
    _header_data: list[tuple[ET.Element, dict[str, str]]]
    _footer_data: list[tuple[ET.Element, dict[str, str]]]
    _current_page_setup: Optional[ldm.PageSetup]
    _current_table_style_id: str
    _bookmark_id_to_name: dict[str, str]
