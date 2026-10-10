"""Cascading Font/ParagraphFormat builders.

Resolution order: docDefaults → table style chain → para/char style chain → direct.
"""

from typing import Optional
from functools import lru_cache
import re
from xml.etree import ElementTree as ET

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.utils.xml_helpers import combine_style_toggle, parse_font_size, resolve_font_names
from aspose.words_foss.docx_reader.constants import (
    COLOR_EMPTY,
    W_NS,
    _ALIGNMENT_MAP,
    _HALF_PT_DIVISOR,
    _HIGHLIGHT_COLOR_MAP,
    _LINE_RULE_MAP,
    _OUTLINE_LEVEL_BODY,
    _TAB_ALIGNMENT_MAP,
    _TAB_LEADER_MAP,
    _TWIPS_PER_PT,
    _UNDERLINE_MAP,
)
from aspose.words_foss.docx_reader.field_mappings import (
    BASELINE_ALIGN_MAP,
    DROP_CAP_POSITION_MAP,
    EMPHASIS_MARK_MAP,
    MERGE_FONT_FIELDS,
    MERGE_PF_FIELDS,
    PARA_ONOFF_FLAGS,
    RUN_ONOFF_FLAGS,
    TEXT_EFFECT_MAP,
)
from aspose.words_foss.docx_reader.utils import (
    _apply_theme_color_modifiers,
    _empty_borders,
    _hex_to_ldm_color,
    apply_onoff_attrs,
    parse_onoff,
)
from aspose.words_foss.model.style_identifiers import resolve_style_identifier

from ._helpers import (
    build_borders,
    build_frame,
    build_shading,
    find_val,
    is_truthy_onoff,
    parse_int,
    parse_universal_measure,
    read_twip,
)
from ._context import ReaderContext


_DEFAULT_PARAGRAPH_FONT_NAME = "Default Paragraph Font"


class StyleChainResolver:
    """Walk the ``<w:basedOn>`` graph for a styleId."""

    def __init__(self, ctx: ReaderContext):
        self._ctx = ctx

    def chain(self, style_id_or_name: str) -> list[str]:
        ctx = self._ctx
        if ctx._styles_xml is None:
            return []
        sid = ctx._name_to_style_id.get(style_id_or_name, style_id_or_name)
        chain: list[str] = []
        visited: set[str] = set()
        while sid and sid not in visited:
            visited.add(sid)
            chain.append(sid)
            elem = ctx._style_elem_cache.get(sid)
            if elem is None:
                break
            based_on = elem.find(f"{W_NS}basedOn")
            sid = based_on.get(f"{W_NS}val", "") if based_on is not None else ""
        chain.reverse()
        return chain

    def numPr_from_style(self, style_id: str) -> Optional[ET.Element]:
        """Find the first ``<w:numPr>`` while walking the chain back to root."""
        for sid in reversed(self.chain(style_id)):
            elem = self._ctx._style_elem_cache.get(sid)
            if elem is None:
                continue
            pPr = elem.find(f"{W_NS}pPr")
            if pPr is None:
                continue
            numPr = pPr.find(f"{W_NS}numPr")
            if numPr is not None:
                return numPr
        return None


class FontBuilder:
    """Build a non-cascaded :class:`ldm.Font` from one ``<w:rPr>``."""

    def __init__(self, ctx: ReaderContext):
        self._ctx = ctx

    def build(self, rPr: ET.Element) -> ldm.Font:
        """Translate one ``<w:rPr>`` into a value-only :class:`ldm.Font`."""
        font = ldm.Font(font_names_explicit=False)
        font.size_explicit = False
        font.color_explicit = rPr.find(f"{W_NS}color") is not None
        for tag, field, _ in RUN_ONOFF_FLAGS:
            setattr(font, field + "_explicit", rPr.find(f"{W_NS}{tag}") is not None)
        self._apply_name(rPr, font)
        self._apply_size(rPr, font)
        apply_onoff_attrs(font, rPr, RUN_ONOFF_FLAGS)
        self._apply_underline(rPr, font)
        self._apply_color(rPr, font)
        self._apply_vert_align(rPr, font)
        self._apply_highlight(rPr, font)
        self._apply_effects(rPr, font)
        self._apply_style_ref(rPr, font)
        self._apply_shading(rPr, font)
        self._apply_kerning(rPr, font)
        self._apply_locale(rPr, font)
        return font

    def _apply_name(self, rPr: ET.Element, font: ldm.Font) -> None:
        rFonts = rPr.find(f"{W_NS}rFonts")
        if rFonts is None:
            return
        names = resolve_font_names(rFonts.attrib, self._ctx._theme_fonts)
        ascii_name, hAnsi_name, cs_name, ea_name = (names.get(channel, '') for channel in ('ascii', 'hAnsi', 'cs', 'eastAsia'))
        explicit = ascii_name or hAnsi_name or cs_name
        if explicit:
            super(ldm.Font, font).__setattr__('name', explicit)
        if hAnsi_name:
            font.name_other = hAnsi_name
        if ascii_name:
            font.name_ascii = ascii_name
        if cs_name:
            font.name_bi = cs_name
        if ea_name:
            font.name_far_east = ea_name
        font.source_font_names = self.source_font_names(rPr)
        font.font_names_explicit = True

    @staticmethod
    def source_font_names(rPr: ET.Element | None) -> ldm.SourceFontNames | None:
        rFonts = rPr.find(f'{W_NS}rFonts') if rPr is not None else None
        if rFonts is None:
            return None
        attributes = {key.removeprefix(W_NS): value for key, value in rFonts.attrib.items()
                      if key.startswith(W_NS) and key.removeprefix(W_NS) in
                      {'ascii', 'hAnsi', 'cs', 'eastAsia', 'asciiTheme', 'hAnsiTheme',
                       'cstheme', 'eastAsiaTheme', 'hint'}}
        return ldm.SourceFontNames(attributes=attributes)

    @staticmethod
    def _apply_size(rPr: ET.Element, font: ldm.Font) -> None:
        size = rPr.find(f"{W_NS}sz")
        if size is not None:
            font.size = parse_font_size(size.get(f"{W_NS}val", ""))
            font.size_explicit = True

    @staticmethod
    def _apply_underline(rPr: ET.Element, font: ldm.Font) -> None:
        u = rPr.find(f"{W_NS}u")
        if u is None:
            return
        u_val = u.get(f"{W_NS}val", "none")
        font.underline = _UNDERLINE_MAP.get(u_val, 1 if u_val != "none" else 0)

    def _apply_color(self, rPr: ET.Element, font: ldm.Font) -> None:
        color = rPr.find(f"{W_NS}color")
        if color is None:
            return
        val = color.get(f"{W_NS}val", "")
        font.color = _hex_to_ldm_color(val)
        font.source_color = ldm.SourceColor(
            value=color.get(f"{W_NS}val"), theme_color=color.get(f"{W_NS}themeColor"),
            theme_tint=color.get(f"{W_NS}themeTint"), theme_shade=color.get(f"{W_NS}themeShade"))
        theme_name = color.get(f"{W_NS}themeColor", "")
        resolved = self._ctx._resolve_theme_color(theme_name) if theme_name else ""
        tint, shade = color.get(f"{W_NS}themeTint"), color.get(f"{W_NS}themeShade")
        # Native 26.9 paints the raw fallback when both modifiers are present.
        if resolved and not (tint is not None and shade is not None):
            font.color_rendering = _hex_to_ldm_color(_apply_theme_color_modifiers(
                resolved, tint=tint, shade=shade))

    @staticmethod
    def _apply_vert_align(rPr: ET.Element, font: ldm.Font) -> None:
        va = rPr.find(f"{W_NS}vertAlign")
        if va is None:
            return
        va_val = va.get(f"{W_NS}val", "")
        font.superscript = va_val == "superscript"
        font.subscript = va_val == "subscript"

    @staticmethod
    def _apply_highlight(rPr: ET.Element, font: ldm.Font) -> None:
        hl = rPr.find(f"{W_NS}highlight")
        if hl is None:
            return
        font.highlight_color = _HIGHLIGHT_COLOR_MAP.get(
            hl.get(f"{W_NS}val", ""), COLOR_EMPTY
        )

    @staticmethod
    def _apply_effects(rPr: ET.Element, font: ldm.Font) -> None:
        effect = rPr.find(f"{W_NS}effect")
        if effect is not None:
            font.text_effect = TEXT_EFFECT_MAP.get(effect.get(f"{W_NS}val", ""), 0)
        em = rPr.find(f"{W_NS}em")
        if em is not None:
            font.emphasis_mark = EMPHASIS_MARK_MAP.get(em.get(f"{W_NS}val", ""), 0)

    def _apply_style_ref(self, rPr: ET.Element, font: ldm.Font) -> None:
        rStyle = rPr.find(f"{W_NS}rStyle")
        if rStyle is None:
            return
        style_id = rStyle.get(f"{W_NS}val", "")
        resolved = self._ctx._resolve_style_name(style_id)
        if resolved != _DEFAULT_PARAGRAPH_FONT_NAME:
            font.style_name = resolved
        sid = resolve_style_identifier(style_id, resolved)
        if sid >= 0:
            font.style_identifier = sid

    @staticmethod
    def _apply_shading(rPr: ET.Element, font: ldm.Font) -> None:
        shd = rPr.find(f"{W_NS}shd")
        if shd is not None:
            font.shading = build_shading(shd)

    @staticmethod
    def _apply_kerning(rPr: ET.Element, font: ldm.Font) -> None:
        # ST_HpsMeasure: bare number (half-points) or ST_UniversalMeasure suffix.
        val = find_val(rPr, "kern")
        if not val:
            return
        hp = parse_universal_measure(val, "hp", default=float("nan"))
        if hp != hp:    # NaN: parsing failed
            return
        font.kerning = hp / _HALF_PT_DIVISOR

    @staticmethod
    def _apply_locale(rPr: ET.Element, font: ldm.Font) -> None:
        lang = rPr.find(f"{W_NS}lang")
        if lang is None:
            return
        val = lang.get(f"{W_NS}val", "")
        if val:
            font.locale_id = _LANG_TAG_TO_LCID.get(val.lower(), 0)
        bidi_val = lang.get(f"{W_NS}bidi", "")
        if bidi_val:
            font.locale_id_bi = _LANG_TAG_TO_LCID.get(bidi_val.lower(), 0)
        ea_val = lang.get(f"{W_NS}eastAsia", "")
        if ea_val:
            font.locale_id_far_east = _LANG_TAG_TO_LCID.get(ea_val.lower(), 0)


_LANG_TAG_TO_LCID: dict[str, int] = {
    "en-us": 1033, "ru-ru": 1049, "zh-cn": 2052, "zh-tw": 1028,
    "ja-jp": 1041, "ko-kr": 1042, "fr-fr": 1036, "de-de": 1031,
    "es-es": 3082, "it-it": 1040, "pt-br": 1046, "ar-sa": 1025,
    "he-il": 1037, "th-th": 1054, "vi-vn": 1066, "pl-pl": 1045,
    "uk-ua": 1058, "cs-cz": 1029, "nl-nl": 1043, "sv-se": 1053,
    "da-dk": 1030, "fi-fi": 1035, "nb-no": 1044, "hu-hu": 1038,
    "ro-ro": 1048, "tr-tr": 1055, "el-gr": 1032, "bg-bg": 1026,
    "hr-hr": 1050, "sk-sk": 1051, "sl-si": 1060,
}


class ParagraphFormatBuilder:
    """Build a non-cascaded :class:`ldm.ParagraphFormat` from one ``<w:pPr>``."""

    def __init__(self, ctx: ReaderContext, font_builder: FontBuilder):
        self._ctx = ctx
        self._fonts = font_builder

    def build(self, pPr: ET.Element) -> ldm.ParagraphFormat:
        """Translate one ``<w:pPr>`` into a value-only :class:`ldm.ParagraphFormat`."""
        pf = ldm.ParagraphFormat()
        self._apply_style_name(pPr, pf)
        self._apply_alignment(pPr, pf)
        self._apply_indent(pPr, pf)
        self._apply_spacing(pPr, pf)
        apply_onoff_attrs(pf, pPr, PARA_ONOFF_FLAGS)
        self._apply_text_alignment(pPr, pf)
        self._apply_conditional(pPr, pf)
        self._apply_outline_level(pPr, pf)
        self._apply_shading(pPr, pf)
        self._apply_borders(pPr, pf)
        self._apply_tab_stops(pPr, pf)
        self._apply_mark_font(pPr, pf)
        self._apply_frame(pPr, pf)
        return pf

    def _apply_style_name(self, pPr: ET.Element, pf: ldm.ParagraphFormat) -> None:
        pStyle = pPr.find(f"{W_NS}pStyle")
        if pStyle is not None:
            pf.style_name = self._ctx._resolve_style_name(pStyle.get(f"{W_NS}val", ""))

    @staticmethod
    def _apply_alignment(pPr: ET.Element, pf: ldm.ParagraphFormat) -> None:
        jc = pPr.find(f"{W_NS}jc")
        if jc is not None:
            pf.alignment = _ALIGNMENT_MAP.get(jc.get(f"{W_NS}val", "left"), 0)

    @staticmethod
    def _apply_indent(pPr: ET.Element, pf: ldm.ParagraphFormat) -> None:
        ind = pPr.find(f"{W_NS}ind")
        if ind is None:
            return
        dimensions = {f"{W_NS}{name}": target for name, target in
                      (("left", "left_indent"), ("start", "left_indent"),
                       ("right", "right_indent"), ("end", "right_indent"),
                       ("firstLine", "first_line_indent"), ("hanging", "first_line_indent"))}
        for attribute in ind.attrib:
            if attribute in dimensions:
                value = read_twip(ind, attribute[len(W_NS):])
                if value is not None:
                    setattr(pf, dimensions[attribute], -value if attribute == f"{W_NS}hanging" else value)
        ParagraphFormatBuilder.apply_character_indents(ind, pf)

    @staticmethod
    def apply_character_indents(ind: ET.Element, pf: ldm.ParagraphFormat | ldm.ListLevel) -> None:
        characters = {f"{W_NS}{name}": target for name, target in
                      (("leftChars", "character_unit_left_indent"), ("startChars", "character_unit_left_indent"),
                       ("rightChars", "character_unit_right_indent"), ("endChars", "character_unit_right_indent"),
                       ("firstLineChars", "character_unit_first_line_indent"), ("hangingChars", "character_unit_first_line_indent"))}
        for attribute, raw in ind.attrib.items():
            if attribute in characters:
                try:
                    if re.fullmatch(r"[+-]?[0-9]+", raw.strip(" \t\r\n")) is None:
                        raise ValueError
                    value = int(raw) / 100
                except (ValueError, OverflowError):
                    raise ValueError("Character indents require integer hundredths") from None
                setattr(pf, characters[attribute], -value if attribute == f"{W_NS}hangingChars" else value)

    @staticmethod
    def _apply_spacing(pPr: ET.Element, pf: ldm.ParagraphFormat) -> None:
        spacing = pPr.find(f"{W_NS}spacing")
        if spacing is None:
            return
        for attr, target in (
            ("before", "space_before"),
            ("after", "space_after"),
            ("line", "line_spacing"),
        ):
            value = read_twip(spacing, attr)
            if value is not None:
                setattr(pf, target, value)
        line_rule = spacing.get(f"{W_NS}lineRule", "")
        if line_rule:
            pf.line_spacing_rule = _LINE_RULE_MAP.get(line_rule, 0)
        if is_truthy_onoff(spacing.get(f"{W_NS}beforeAutospacing", "")):
            pf.space_before_auto = True
        if is_truthy_onoff(spacing.get(f"{W_NS}afterAutospacing", "")):
            pf.space_after_auto = True

    @staticmethod
    def _apply_text_alignment(pPr: ET.Element, pf: ldm.ParagraphFormat) -> None:
        ta = pPr.find(f"{W_NS}textAlignment")
        if ta is not None:
            pf.baseline_alignment = BASELINE_ALIGN_MAP.get(ta.get(f"{W_NS}val", ""), 0)

    @staticmethod
    def _apply_conditional(pPr: ET.Element, pf: ldm.ParagraphFormat) -> None:
        cnf = pPr.find(f"{W_NS}cnfStyle")
        if cnf is not None:
            pf.conditional_style = ldm.ConditionalStyleMask.from_val(cnf.get(f"{W_NS}val", ""))

    @staticmethod
    def _apply_outline_level(pPr: ET.Element, pf: ldm.ParagraphFormat) -> None:
        outline = pPr.find(f"{W_NS}outlineLvl")
        if outline is None:
            return
        pf.outline_level = parse_int(outline.get(f"{W_NS}val", "9"), 9)
        if pf.outline_level < _OUTLINE_LEVEL_BODY:
            pf.is_heading = True

    @staticmethod
    def _apply_shading(pPr: ET.Element, pf: ldm.ParagraphFormat) -> None:
        shd = pPr.find(f"{W_NS}shd")
        if shd is not None:
            pf.shading = build_shading(shd)

    @staticmethod
    def _apply_borders(pPr: ET.Element, pf: ldm.ParagraphFormat) -> None:
        pBdr = pPr.find(f"{W_NS}pBdr")
        if pBdr is not None:
            pf.borders = build_borders(pBdr)

    @staticmethod
    def _apply_tab_stops(pPr: ET.Element, pf: ldm.ParagraphFormat) -> None:
        tabs_elem = pPr.find(f"{W_NS}tabs")
        if tabs_elem is None:
            return
        tab_stops = ldm.TabStopCollection()
        for tab in tabs_elem.findall(f"{W_NS}tab"):
            val = tab.get(f"{W_NS}val", "left")
            tab_stops.tab_stops.append(
                ldm.TabStop(
                    position=int(tab.get(f"{W_NS}pos", "0")) / _TWIPS_PER_PT,
                    alignment=_TAB_ALIGNMENT_MAP.get(val, 0),
                    leader=_TAB_LEADER_MAP.get(tab.get(f"{W_NS}leader", "none"), 0),
                    is_clear=val == "clear",
                )
            )
        pf.tab_stops = tab_stops

    def _apply_mark_font(self, pPr: ET.Element, pf: ldm.ParagraphFormat) -> None:
        mark_rPr = pPr.find(f"{W_NS}rPr")
        if mark_rPr is not None:
            pf.paragraph_break_font = self._fonts.build(mark_rPr)

    @staticmethod
    def _apply_frame(pPr: ET.Element, pf: ldm.ParagraphFormat) -> None:
        frame_pr = pPr.find(f"{W_NS}framePr")
        if frame_pr is None:
            return
        pf.frame_format = build_frame(frame_pr)
        raw_lines = frame_pr.get(f"{W_NS}lines")
        if raw_lines is not None:
            try:
                pf.lines_to_drop = int(raw_lines)
            except ValueError:
                pass
        raw_drop_cap = frame_pr.get(f"{W_NS}dropCap")
        if raw_drop_cap is not None:
            pf.drop_cap_position = DROP_CAP_POSITION_MAP.get(raw_drop_cap, 0)


class FontResolver:
    """Compose a fully cascaded :class:`ldm.Font`."""

    def __init__(
        self,
        ctx: ReaderContext,
        fonts: FontBuilder,
        chains: StyleChainResolver,
    ):
        self._ctx = ctx
        self._fonts = fonts
        self._chains = chains
        # Instance-owned cache avoids retaining previous documents through a global method cache.
        self._inherited_font = lru_cache(maxsize=128)(self._inherited_font)

    def resolve(self, rPr: Optional[ET.Element], style_id: str = "", implicit_paragraph_style: bool = False) -> ldm.Font:
        """Compose a Font, walking docDefaults, table style, *style_id* chain, then *rPr*."""
        character_style = rPr.find(f"{W_NS}rStyle") if rPr is not None else None
        character_id = character_style.get(f"{W_NS}val", "") if character_style is not None else ""
        table_id = getattr(self._ctx, "_current_table_style_id", "")
        base = self._inherited_font(table_id, style_id, character_id, implicit_paragraph_style).model_copy(deep=True)
        if rPr is not None:
            self.merge(base, self._fonts.build(rPr))
        if not base.style_name:
            base.style_name = _DEFAULT_PARAGRAPH_FONT_NAME
        if not base.color:
            base.color = COLOR_EMPTY
        if not base.highlight_color:
            base.highlight_color = COLOR_EMPTY
        base.source_font_names = self._fonts.source_font_names(rPr)
        base.font_names_explicit = base.source_font_names is not None
        base.size_explicit = rPr is not None and rPr.find(f"{W_NS}sz") is not None
        base.color_explicit = rPr is not None and rPr.find(f"{W_NS}color") is not None
        base._name_context = (self._ctx._resolve_style_name(table_id),
                              self._ctx._resolve_style_name(style_id),
                              self._ctx._resolve_style_name(character_id),
                              tuple(getattr(self._ctx, "_current_table_conditions", ())))
        for tag, field, _ in RUN_ONOFF_FLAGS:
            setattr(base, field + "_explicit", rPr is not None and rPr.find(f"{W_NS}{tag}") is not None)
        return base

    def _inherited_font(self, table_id: str, style_id: str, character_id: str,
                        implicit_paragraph_style: bool = False) -> ldm.Font:
        ctx = self._ctx
        base = ldm.Font(color=COLOR_EMPTY, highlight_color=COLOR_EMPTY)

        if ctx._doc_default_rPr is not None:
            base = self._fonts.build(ctx._doc_default_rPr)
            if not base.color:
                base.color = COLOR_EMPTY
            if not base.highlight_color:
                base.highlight_color = COLOR_EMPTY

        if "size" not in base.model_fields_set:
            present = (ctx._styles_xml is not None and
                       ctx._styles_xml.find(f"{W_NS}docDefaults/{W_NS}rPrDefault") is not None)
            base.size = 10.0 if present else 11.0

        if table_id:
            self._merge_style_chain(base, table_id)
        initial = {tag: getattr(base, field) for tag, field, _ in RUN_ONOFF_FLAGS}
        for inherited in (style_id, character_id):
            if inherited:
                self._merge_style_chain(base, inherited)
        for tag, field, _ in RUN_ONOFF_FLAGS:
            setattr(base, field, combine_style_toggle(initial[tag], self._style_toggle(style_id, tag),
                                                      self._style_toggle(character_id, tag)))
        paragraph = self._style_toggle(style_id, 'vanish')
        character = self._style_toggle(character_id, 'vanish')
        rendered = base.hidden
        if implicit_paragraph_style:
            rendered = (character if character is not None else paragraph
                        if paragraph is not None else initial['vanish'])
        elif paragraph is not None and character is not None:
            rendered = initial['vanish'] ^ paragraph ^ character
        if rendered != base.hidden:
            base.hidden_rendering = rendered
        return base

    def _style_toggle(self, style_id: str, tag: str) -> bool | None:
        for sid in reversed(self._chains.chain(style_id)):
            style = self._ctx._style_elem_cache.get(sid)
            element = style.find(f"{W_NS}rPr/{W_NS}{tag}") if style is not None else None
            if element is not None:
                return parse_onoff(element)
        return None

    def _merge_style_chain(self, base: ldm.Font, style_id: str) -> None:
        for sid in self._chains.chain(style_id):
            elem = self._ctx._style_elem_cache.get(sid)
            if elem is None:
                continue
            srPr = elem.find(f"{W_NS}rPr")
            if srPr is not None:
                self.merge(base, self._fonts.build(srPr))

    @staticmethod
    def merge(base: ldm.Font, override: ldm.Font) -> None:
        """Merge *override* into *base* using Pydantic ``model_fields_set``."""
        _set: set[str] = override.model_fields_set
        name_fields = ('name', 'name_ascii', 'name_other', 'name_bi', 'name_far_east')
        source = override.source_font_names
        if source is not None:
            # Preserve effective aliases before a partial ASCII declaration changes the primary.
            old_names = {field: getattr(base, field) or base.name for field in name_fields[1:]}
            for field, value in old_names.items():
                super(ldm.Font, base).__setattr__(field, value)
            for field, literal, theme in (
                ('name_ascii', 'ascii', 'asciiTheme'), ('name_other', 'hAnsi', 'hAnsiTheme'),
                ('name_bi', 'cs', 'cstheme'), ('name_far_east', 'eastAsia', 'eastAsiaTheme')):
                if literal not in source.attributes and theme not in source.attributes:
                    continue
                value = getattr(override, field)
                if value:
                    super(ldm.Font, base).__setattr__(field, value)
                    if field == 'name_ascii':
                        super(ldm.Font, base).__setattr__('name', value)
        for field in MERGE_FONT_FIELDS:
            if source is not None and field in name_fields:
                continue
            if field not in _set:
                continue
            value = getattr(override, field)
            # Only an undeclared zero is unspecified after JSON validation.
            if field == "size" and value == 0 and override.size_explicit is not True:
                continue
            if (field in ("color", "highlight_color") and value == COLOR_EMPTY
                    and not (field == "color" and override.color_explicit is True)):
                continue
            if field in name_fields:
                # Cascade resolution is not a user's explicit font edit.
                super(ldm.Font, base).__setattr__(field, value)
            else:
                setattr(base, field, value)


class ParagraphFormatResolver:
    """Compose a fully cascaded :class:`ldm.ParagraphFormat`."""

    def __init__(
        self,
        ctx: ReaderContext,
        pfs: ParagraphFormatBuilder,
        chains: StyleChainResolver,
    ):
        self._ctx = ctx
        self._pfs = pfs
        self._chains = chains

    def resolve(
        self,
        pPr: Optional[ET.Element],
        style_id: str = "",
    ) -> ldm.ParagraphFormat:
        """Compose a ParagraphFormat through docDefaults → style chain → *pPr*."""
        ctx = self._ctx
        base = ldm.ParagraphFormat(borders=_empty_borders())
        if ctx._doc_default_pPr is not None:
            base = self._pfs.build(ctx._doc_default_pPr)

        table_style_id = getattr(ctx, "_current_table_style_id", "")
        if table_style_id:
            self._merge_style_chain(base, table_style_id)

        if style_id:
            self._merge_style_chain(base, style_id)

        if pPr is not None:
            self.merge(base, self._pfs.build(pPr))

        if style_id and not base.style_name:
            base.style_name = ctx._resolve_style_name(style_id)
        if not base.borders:
            base.borders = _empty_borders()
        return base

    def _merge_style_chain(self, base: ldm.ParagraphFormat, style_id: str) -> None:
        for sid in self._chains.chain(style_id):
            elem = self._ctx._style_elem_cache.get(sid)
            if elem is None:
                continue
            spPr = elem.find(f"{W_NS}pPr")
            if spPr is not None:
                self.merge(base, self._pfs.build(spPr))

    @staticmethod
    def merge(base: ldm.ParagraphFormat, override: ldm.ParagraphFormat) -> None:
        """Merge *override* into *base* with tab-stop clear-position semantics."""
        _set: set[str] = override.model_fields_set
        for field in MERGE_PF_FIELDS:
            if field not in _set:
                continue
            if field == "paragraph_break_font" and base.paragraph_break_font is not None and override.paragraph_break_font is not None:
                mark = base.paragraph_break_font.model_copy(deep=True)
                FontResolver.merge(mark, override.paragraph_break_font)
                base.paragraph_break_font = mark
                continue
            if field == "shading":
                val = getattr(override, field)
                if val.background_pattern_color in ("", "auto") and not val.foreground_pattern_color:
                    continue
            setattr(base, field, getattr(override, field))
        if "tab_stops" not in _set:
            return
        if not override.tab_stops:
            base.tab_stops = override.tab_stops
            return
        clear_positions = {t.position for t in override.tab_stops if t.is_clear}
        if clear_positions:
            base.tab_stops.tab_stops = [
                t for t in base.tab_stops.tab_stops
                if t.position not in clear_positions
            ]
        for t in override.tab_stops:
            if t.is_clear:
                continue
            base.tab_stops.remove_by_position(t.position)
            base.tab_stops.tab_stops.append(t)
        base.tab_stops.tab_stops.sort(key=lambda x: x.position)
