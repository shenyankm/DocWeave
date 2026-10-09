"""Run (``<w:r>``) and run-properties (``<w:rPr>``) rendering.

The reader flattens hyperlinks into runs whose text reads
``[display](url)``.  We detect that pattern and re-emit it as a real
``<w:hyperlink>`` element so the round-trip is lossless.
"""


import re
from typing import Mapping, Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader.constants import PAGE_FIELD_SENTINEL
from aspose.words_foss.docx_writer.constants import (
    HIGHLIGHT_NAME_BY_COLOR,
    UNDERLINE_VAL,
    pt_to_half_pt,
)
from aspose.words_foss.docx_writer.xml_utils import el, text_run
from aspose.words_foss.model.style_identifiers import IDENTIFIER_TO_STYLE_ID
from aspose.words_foss._links import INLINE_LINK_RE, decode_link

_HEX_COLOR_RE = re.compile(r"Color \[A=(?P<a>\d+), R=(?P<r>\d+), G=(?P<g>\d+), B=(?P<b>\d+)\]")
_RAW_HEX_RE = re.compile(r"^[0-9A-Fa-f]{6}$")

_LCID_TO_LANG_TAG: dict[int, str] = {
    1033: "en-US", 1049: "ru-RU", 2052: "zh-CN", 1028: "zh-TW",
    1041: "ja-JP", 1042: "ko-KR", 1036: "fr-FR", 1031: "de-DE",
    3082: "es-ES", 1040: "it-IT", 1046: "pt-BR", 1025: "ar-SA",
    1037: "he-IL", 1054: "th-TH", 1066: "vi-VN", 1045: "pl-PL",
    1058: "uk-UA", 1029: "cs-CZ", 1043: "nl-NL", 1053: "sv-SE",
    1030: "da-DK", 1035: "fi-FI", 1044: "nb-NO", 1038: "hu-HU",
    1048: "ro-RO", 1055: "tr-TR", 1032: "el-GR", 1026: "bg-BG",
    1050: "hr-HR", 1051: "sk-SK", 1060: "sl-SI",
}

#: ``EmphasisMark`` int → OOXML ``w:em/@val`` token.
_EMPHASIS_MARK_TOKEN: dict[int, str] = {
    0: "none",
    1: "dot",       # OverSolidCircle
    2: "comma",     # OverComma
    3: "circle",    # OverWhiteCircle
    4: "underDot",  # UnderSolidCircle
}

#: ``TextEffect`` int → OOXML ``w:effect/@val`` token.
_TEXT_EFFECT_TOKEN: dict[int, str] = {
    0: "none",
    1: "lights",            # LasVegasLights
    2: "blinkBackground",   # BlinkingBackground
    3: "sparkle",           # SparkleText
    4: "antsBlack",         # MarchingBlackAnts
    5: "antsRed",           # MarchingRedAnts
    6: "shimmer",           # Shimmer
}

# Schemes the writer accepts as a real hyperlink target.  The list mirrors
# what the reader produces from ``w:hyperlink`` (HTTP/S, mailto, tel,
# file, ftp) plus pure fragment anchors (``#bookmark``).  Any run text
# matching ``[x](y)`` whose ``y`` doesn't satisfy this gate is left as a
# literal string, so inline-code samples like ``[1](note)`` aren't
# silently rewritten into a Word hyperlink.
_URL_RE = re.compile(r"^(?:[a-zA-Z][a-zA-Z0-9+.\-]*:|#|/|\.{1,2}/)")


def color_to_hex(color: str) -> Optional[str]:
    """Coerce an LDM color string to upper-case ``RRGGBB`` hex.

    The reader stores colors in two shapes — ``"Color [A=255, R=…, G=…, B=…]"``
    for explicit colors (font.color, paragraph shading) and the raw 6-hex
    string Word puts in ``<w:shd w:fill>`` (font shading).  Returns ``None``
    for empty / sentinel values so the caller can omit the corresponding
    XML element entirely.
    """
    if not color or color == "Color [Empty]":
        return None
    m = _HEX_COLOR_RE.match(color)
    if m:
        r = int(m.group("r"))
        g = int(m.group("g"))
        b = int(m.group("b"))
        return f"{r:02X}{g:02X}{b:02X}"
    if _RAW_HEX_RE.match(color):
        return color.upper()
    return None


_DEFAULT_FONT: Optional[ldm.Font] = None


def _default_font() -> ldm.Font:
    """Lazy-cached zero-filled :class:`Font` for the no-base default."""
    global _DEFAULT_FONT
    if _DEFAULT_FONT is None:
        _DEFAULT_FONT = ldm.Font()
    return _DEFAULT_FONT


def _bool_toggle(name: str, pf_value: bool, base_value: bool) -> Optional[str]:
    """Emit a CT_OnOff toggle when ``pf_value`` differs from ``base_value``.

    Returns ``None`` when no emission is needed.  Used for ``<w:b>``,
    ``<w:i>``, ``<w:caps>``, ``<w:smallCaps>``, ``<w:strike>``,
    ``<w:vanish>`` — each of which the reader treats as boolean and the
    schema lets us flip off via ``w:val="0"``.
    """
    if pf_value == base_value:
        return None
    return el(name) if pf_value else el(name, {"w:val": "0"})


def render_rPr(
    font: ldm.Font,
    *,
    style_id: str = "",
    base: Optional[ldm.Font] = None,
    for_style: bool = False,
    preserve_explicit_off: bool = False,
    style_id_map: Optional[Mapping[str, str]] = None,
) -> str:
    """Render a ``<w:rPr>`` element for a run.

    Children are emitted in the order required by the OOXML schema
    (``CT_RPr``): rStyle, rFonts, b, i, caps, smallCaps, strike, vanish,
    color, sz, szCs, highlight, u, vertAlign, shd.  ECMA-376 strict
    consumers reject out-of-order properties.

    ``preserve_explicit_off=True`` retains explicit false flags in raw style fonts.

    ``for_style=True`` switches to ``CT_StyleRPr`` rules (used inside
    ``<w:style>``), which forbids ``<w:highlight>``.

    ``base`` is the resolved-font baseline this run inherits from
    (paragraph style chain for body runs, basedOn style for character /
    paragraph styles).  When ``pf == base`` for a field nothing is
    emitted; when they differ we emit the value — including the explicit
    "off" toggle (``w:val="0"``) needed to disable a boolean the
    inherited chain would otherwise turn on.  ``base=None`` falls back
    to a zero default for callers without chain context.

    Returns an empty string if every field matches the base — callers
    should only wrap this in an outer element when the result is
    non-empty.
    """
    if base is None:
        base = _default_font()
    children: list[str] = []

    # When no explicit ``style_id`` is provided, derive it from the LDM
    # font.  Prefer the locale-independent ``style_identifier`` (maps to
    # the canonical English ``w:styleId``); fall back to ``style_name``
    # for custom / user-defined character styles.
    if not style_id and style_id_map is not None and font.style_name not in _DEFAULT_FONT_STYLE_NAMES:
        style_id = style_id_map.get(font.style_name, "")
    if not style_id and font.style_identifier and font.style_identifier in IDENTIFIER_TO_STYLE_ID:
        style_id = IDENTIFIER_TO_STYLE_ID[font.style_identifier]
    if not style_id and font.style_name and font.style_name not in _DEFAULT_FONT_STYLE_NAMES:
        from aspose.words_foss.docx_writer.styles_part import _sanitize_style_id
        # Sanitize so non-alphanumeric characters in a custom style name
        # (e.g. ``"*Hyperlink"``) don't produce an unaddressable
        # ``w:rStyle w:val`` that MS Word rejects.
        style_id = _sanitize_style_id(font.style_name.replace(" ", ""))
    if style_id:
        children.append(el("w:rStyle", {"w:val": style_id}))
    rfont_changed = (
        font.name != base.name
        or font.name_ascii != base.name_ascii
        or font.name_bi != base.name_bi
        or font.name_far_east != base.name_far_east
    )
    if rfont_changed:
        ascii_name = font.name_ascii or font.name
        hAnsi_name = font.name
        cs_name = font.name_bi or font.name
        ea_name = font.name_far_east or font.name
        attrs: dict[str, str] = {}
        if ascii_name:
            attrs["w:ascii"] = ascii_name
        if hAnsi_name:
            attrs["w:hAnsi"] = hAnsi_name
        if cs_name:
            attrs["w:cs"] = cs_name
        if ea_name:
            attrs["w:eastAsia"] = ea_name
        if attrs:
            children.append(el("w:rFonts", attrs))
    for tag, field in (
        ("w:b", "bold"), ("w:bCs", "bold_bi"),
        ("w:i", "italic"), ("w:iCs", "italic_bi"),
        ("w:caps", "all_caps"), ("w:smallCaps", "small_caps"),
        ("w:strike", "strike_through"), ("w:outline", "outline"),
        ("w:shadow", "shadow"), ("w:emboss", "emboss"),
        ("w:imprint", "engrave"), ("w:noProof", "no_proofing"),
        ("w:vanish", "hidden"),
    ):
        val = getattr(font, field)
        if field in ("bold", "italic") and getattr(font, field + "_explicit") is True:
            toggle = el(tag, {"w:val": "1" if val else "0"})
        elif preserve_explicit_off and field in font.model_fields_set and not val:
            toggle = el(tag, {"w:val": "0"})
        else:
            toggle = _bool_toggle(tag, val, getattr(base, field))
        if toggle is not None:
            children.append(toggle)
    if font.color != base.color:
        color_hex = color_to_hex(font.color)
        if color_hex:
            children.append(el("w:color", {"w:val": color_hex}))
        else:
            children.append(el("w:color", {"w:val": "auto"}))
    if font.kerning != base.kerning:
        children.append(el("w:kern", {"w:val": pt_to_half_pt(font.kerning)}))
    if font.size > 0 and (font.size != base.size or font.size_explicit is True):
        sz = pt_to_half_pt(font.size)
        children.append(el("w:sz", {"w:val": sz}))
        children.append(el("w:szCs", {"w:val": sz}))
    if not for_style and font.highlight_color != base.highlight_color:
        highlight = HIGHLIGHT_NAME_BY_COLOR.get(font.highlight_color, "none")
        children.append(el("w:highlight", {"w:val": highlight}))
    if font.underline != base.underline:
        underline_val = UNDERLINE_VAL.get(font.underline, "none")
        children.append(el("w:u", {"w:val": underline_val}))
    if font.text_effect != base.text_effect:
        children.append(
            el("w:effect", {"w:val": _TEXT_EFFECT_TOKEN.get(font.text_effect, "none")})
        )
    if font.shading.background_pattern_color != base.shading.background_pattern_color:
        shading_hex = color_to_hex(font.shading.background_pattern_color)
        if shading_hex:
            children.append(
                el(
                    "w:shd",
                    {"w:val": "clear", "w:color": "auto", "w:fill": shading_hex},
                )
            )
        elif font.shading.background_pattern_color:
            children.append(
                el(
                    "w:shd",
                    {
                        "w:val": "clear",
                        "w:color": "auto",
                        "w:fill": font.shading.background_pattern_color,
                    },
                )
            )
        else:
            children.append(el("w:shd", {"w:val": "nil"}))
    if font.superscript != base.superscript or font.subscript != base.subscript:
        if font.superscript:
            children.append(el("w:vertAlign", {"w:val": "superscript"}))
        elif font.subscript:
            children.append(el("w:vertAlign", {"w:val": "subscript"}))
        else:
            children.append(el("w:vertAlign", {"w:val": "baseline"}))
    if font.emphasis_mark != base.emphasis_mark:
        children.append(
            el("w:em", {"w:val": _EMPHASIS_MARK_TOKEN.get(font.emphasis_mark, "none")})
        )
    lang_changed = (
        (font.locale_id != base.locale_id and font.locale_id)
        or (font.locale_id_bi != base.locale_id_bi and font.locale_id_bi)
        or (font.locale_id_far_east != base.locale_id_far_east and font.locale_id_far_east)
    )
    if lang_changed:
        lang_attrs: dict[str, str] = {}
        if font.locale_id and font.locale_id != base.locale_id:
            tag = _LCID_TO_LANG_TAG.get(font.locale_id)
            if tag:
                lang_attrs["w:val"] = tag
        if font.locale_id_bi and font.locale_id_bi != base.locale_id_bi:
            tag = _LCID_TO_LANG_TAG.get(font.locale_id_bi)
            if tag:
                lang_attrs["w:bidi"] = tag
        if font.locale_id_far_east and font.locale_id_far_east != base.locale_id_far_east:
            tag = _LCID_TO_LANG_TAG.get(font.locale_id_far_east)
            if tag:
                lang_attrs["w:eastAsia"] = tag
        if lang_attrs:
            children.append(el("w:lang", lang_attrs))

    if not children:
        return ""
    return el("w:rPr", None, children)


_DEFAULT_FONT_STYLE_NAMES = frozenset({"", "Default Paragraph Font"})


def _render_plain_run(
    run: ldm.Run,
    *,
    style_id: str = "",
    base_font: Optional[ldm.Font] = None,
    instr: bool = False,
    style_id_map: Optional[Mapping[str, str]] = None,
) -> str:
    """Render a single ``<w:r>``.

    ``style_id`` overrides the run's ``font.style_name``; the hyperlink
    path uses it to force ``rStyle="Hyperlink"`` even on runs that
    didn't carry it.  When unset, ``render_rPr`` derives the rStyle
    from the font itself.

    ``base_font`` is the paragraph style chain's resolved font — the
    diff against this lets the writer emit explicit toggles for
    properties the run overrides relative to its style (e.g. an
    ``rPr`` carrying ``<w:b w:val="0"/>`` when the chain set bold).

    ``instr=True`` switches the text element to ``<w:instrText>`` —
    used for runs that sit inside a field's code section (between
    ``begin`` and ``separate``).
    """
    rpr = render_rPr(run.font, style_id=style_id, base=base_font, style_id_map=style_id_map)
    parts: list[str] = []
    if rpr:
        parts.append(rpr)
    parts.append(text_run(run.text, instr=instr))
    return el("w:r", None, parts)


# Word's built-in character style for hyperlink-decorated text.  Word
# itself defines this in styles.xml as blue + underlined; we attach
# rStyle="Hyperlink" so consumers that respect the style chain pick up
# that visual treatment automatically, even when the inner LDM run
# doesn't carry an explicit color/underline.
_HYPERLINK_STYLE_ID = "Hyperlink"


def _render_hyperlink_run(
    text: str,
    url: str,
    run: ldm.Run,
    rels: dict,
    *,
    base_font: Optional[ldm.Font] = None,
    style_id_map: Optional[Mapping[str, str]] = None,
) -> str:
    """Emit ``<w:hyperlink>`` for an inline ``[text](url)`` run.

    Targets that begin with ``#`` are treated as bookmark anchors and
    emitted via ``w:anchor``; Word would otherwise route an external
    rel of ``Target="#name"`` through the URL handler instead of the
    in-document jump.  No relationship row is created for anchors.

    ``rStyle="Hyperlink"`` is attached only when the LDM run already
    carried that character style — preserving the source's choice
    (some authors leave hyperlinks unstyled and rely on direct
    color/underline) so the round-trip doesn't synthesise a new style
    reference and pollute the LDM ``font.style_name``.
    """
    use_hyperlink_style = (run.font.style_name or "").replace(" ", "").lower() == "hyperlink"
    inner = ldm.Run(text=text, font=run.font)
    inner_xml = _render_plain_run(
        inner,
        style_id=(style_id_map or {}).get(run.font.style_name, _HYPERLINK_STYLE_ID) if use_hyperlink_style else "",
        base_font=base_font,
        style_id_map=style_id_map,
    )
    if url.startswith("#"):
        anchor = url[1:]
        return el(
            "w:hyperlink",
            {"w:anchor": anchor, "w:history": "1"},
            inner_xml,
        )
    r_id = rels.setdefault(url, f"rIdLink{len(rels) + 1}")
    return el(
        "w:hyperlink",
        {"r:id": r_id, "w:history": "1"},
        inner_xml,
    )


def _render_page_field_run(run: ldm.Run, style_id_map: Optional[Mapping[str, str]] = None,
                           base_font: Optional[ldm.Font] = None) -> str:
    """Re-emit a ``PAGE`` field whose body the reader collapsed to a sentinel.

    The reader replaces a four-run field chain (``begin`` → ``instrText`` →
    ``separate`` → cached value → ``end``) with one run carrying
    ``PAGE_FIELD_SENTINEL`` (``\\x00PAGE\\x00``).  Reconstructing the chain
    on write means the next read sees the same sentinel rather than literal
    ``"PAGE"`` text or a stale page number.

    The sentinel is emitted by the reader at the ``separate`` boundary, so
    the ``rPr`` of the LDM run must travel with that specific ``<w:r>``.
    Other field-chain runs carry no formatting.
    """
    rpr = render_rPr(run.font, style_id_map=style_id_map, base=base_font)
    rpr_part = rpr if rpr else ""
    begin = el("w:r", None, el("w:fldChar", {"w:fldCharType": "begin"}))
    instr = el(
        "w:r",
        None,
        el("w:instrText", {"xml:space": "preserve"}, "PAGE"),
    )
    separate = el(
        "w:r",
        None,
        [rpr_part, el("w:fldChar", {"w:fldCharType": "separate"})],
    )
    end = el("w:r", None, el("w:fldChar", {"w:fldCharType": "end"}))
    return begin + instr + separate + end


def render_run(
    run: ldm.Run,
    rels: dict,
    *,
    base_font: Optional[ldm.Font] = None,
    instr: bool = False,
    style_id_map: Optional[Mapping[str, str]] = None,
) -> str:
    """Render a single LDM run to OOXML.

    ``rels`` is a mutable dict shared across the document used to
    accumulate hyperlink relationships (URL → relationship id).  Each
    new URL gets a fresh ``rIdLink<N>``.

    The Markdown-link rewrite is gated on the URL satisfying
    :data:`_URL_RE` (a scheme like ``http:``, ``mailto:``, a leading
    ``#``/``/``/``./``).  This stops literal text such as ``[1](note)``
    or inline-code samples from being silently turned into hyperlinks.

    ``base_font`` is the resolved style-chain font for the paragraph
    that contains this run; passing it lets the writer emit explicit
    toggles (``<w:b w:val="0"/>`` etc.) only when the run actually
    overrides the inherited value, instead of suppressing every
    override and silently inheriting the chain on re-read.
    """
    if not run.text:
        return ""
    if run.text == PAGE_FIELD_SENTINEL:
        return _render_page_field_run(run, style_id_map, base_font)
    if not instr:
        token, separator, suffix = run.text.partition("\t")
        match = INLINE_LINK_RE.fullmatch(token)
        if match:
            label, target = decode_link(match)
            if _URL_RE.match(target):
                parts = _render_hyperlink_run(label, target, run, rels, base_font=base_font,
                                              style_id_map=style_id_map)
                if separator:
                    suffix_run = ldm.Run(text=separator + suffix, font=run.font)
                    parts += _render_plain_run(suffix_run, base_font=base_font, style_id_map=style_id_map)
                return parts
    return _render_plain_run(run, base_font=base_font, instr=instr, style_id_map=style_id_map)
