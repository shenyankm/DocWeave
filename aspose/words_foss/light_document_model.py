"""
Lightweight document model for PDF rendering via fpdf2.

The model is genuinely recursive (Cell → list[Table] → list[Row] →
list[Cell]), so this is one of the rare modules that legitimately
needs ``from __future__ import annotations`` to keep all field
type-hints lazy — pydantic's ``model_rebuild`` then stitches the
forward references at the bottom of the file.
"""

from __future__ import annotations

from base64 import b64decode, b64encode
from collections.abc import Mapping
from typing import Annotated, Any, Literal, Optional, Union

from pydantic import (
    BaseModel, BeforeValidator, ConfigDict, Field, PrivateAttr, field_serializer, field_validator,
    model_serializer, model_validator,
)

from aspose.words_foss._io import MAX_TABLE_COLUMNS
from aspose.words_foss.model.enums.image import ImageType as _IT, _IMAGE_TYPE_TO_MIME, _MIME_TO_IMAGE_TYPE
from aspose.words_foss.model.enums.table import PreferredWidthType as _PWT


class NodeType:
    """Node type discriminators, valued as in ``aspose.words``."""

    ANY = 0
    DOCUMENT = 1
    SECTION = 2
    BODY = 3
    HEADER_FOOTER = 4
    TABLE = 5
    ROW = 6
    CELL = 7
    PARAGRAPH = 8
    BOOKMARK_START = 9
    BOOKMARK_END = 10
    SHAPE = 18
    RUN = 21
    FIELD_START = 22
    FIELD_SEPARATOR = 23
    FIELD_END = 24


_NODE_TYPE_BY_CLASS = {
    "Body": NodeType.BODY,
    "BookmarkEnd": NodeType.BOOKMARK_END,
    "BookmarkStart": NodeType.BOOKMARK_START,
    "Cell": NodeType.CELL,
    "Document": NodeType.DOCUMENT,
    "FieldEnd": NodeType.FIELD_END,
    "FieldSeparator": NodeType.FIELD_SEPARATOR,
    "FieldStart": NodeType.FIELD_START,
    "HeaderFooter": NodeType.HEADER_FOOTER,
    "Paragraph": NodeType.PARAGRAPH,
    "Row": NodeType.ROW,
    "Run": NodeType.RUN,
    "Section": NodeType.SECTION,
    "Shape": NodeType.SHAPE,
    "Table": NodeType.TABLE,
}


def _node_type_of(node: object) -> int:
    """Return the :class:`NodeType` of an LDM node, or ``-1`` if unmapped."""
    return _NODE_TYPE_BY_CLASS.get(type(node).__name__, -1)


def _walk_children(node: object):
    """Yield a node's direct child *nodes*, in declaration order.

    Walks declared model fields, not a hand-written attribute list, so every
    container is covered. Derived properties (``Body.paragraphs``) are skipped
    deliberately: they are views over ``children`` and would double-count.
    """
    if isinstance(node, Cell):
        yield from node.children
        return
    for child in getattr(node, "_children", None) or ():
        if type(child).__name__ in _NODE_TYPE_BY_CLASS:
            yield child

    for name in getattr(type(node), "model_fields", {}):
        try:
            value = getattr(node, name)
        except AttributeError:
            continue
        items = value if isinstance(value, list) else (value,)
        for item in items:
            if item is not None and type(item).__name__ in _NODE_TYPE_BY_CLASS:
                yield item


class NodeCastMixin:
    """``as_*()`` casts mirroring ``aspose.words.Node``; identity here."""

    def _cast(self, expected: str):
        if type(self).__name__ != expected:
            raise ValueError(
                f"cannot cast {type(self).__name__} to {expected}"
            )
        return self

    def as_body(self):
        return self._cast("Body")

    def as_cell(self):
        return self._cast("Cell")

    def as_paragraph(self):
        return self._cast("Paragraph")

    def as_row(self):
        return self._cast("Row")

    def as_run(self):
        return self._cast("Run")

    def as_section(self):
        return self._cast("Section")

    def as_shape(self):
        return self._cast("Shape")

    def as_table(self):
        return self._cast("Table")


def _collect_child_nodes(node: object, node_type: int, deep: bool) -> list:
    """Shared implementation of ``get_child_nodes``."""
    found = []
    for child in _walk_children(node):
        if node_type == NodeType.ANY or _node_type_of(child) == node_type:
            found.append(child)
        if deep:
            found.extend(_collect_child_nodes(child, node_type, deep))
    return found


# ─────────────────────────────────────────────
# Primitives
# ─────────────────────────────────────────────


class Border(BaseModel):
    line_style: int = 0
    line_width: float = 0.0
    is_visible: bool = True
    color: str = ""
    distance_from_text: float = 0.0
    shadow: bool = False


class Shading(BaseModel):
    background_pattern_color: str = ""
    foreground_pattern_color: str = ""
    # REMOVED: theme_color, theme_shade, theme_tint, theme_fill,
    #          theme_fill_shade, theme_fill_tint, foreground_tint_and_shade,
    #          background_tint_and_shade, texture


class ConditionalStyleMask(BaseModel):
    """Parsed ``<w:cnfStyle>`` 12-bit bitmask for banded-table regions.

    Each flag says which table region the paragraph / row / cell belongs
    to.  Field names match Aspose's ``ConditionalStyleType`` enum and
    ``ConditionalStyleCollection`` property names.  OOXML stores the
    mask as a 12-char ``"1"``/``"0"`` string.
    """

    first_row: bool = False
    last_row: bool = False
    first_column: bool = False
    last_column: bool = False
    odd_column_banding: bool = False
    even_column_banding: bool = False
    odd_row_banding: bool = False
    even_row_banding: bool = False
    top_right_cell: bool = False
    top_left_cell: bool = False
    bottom_right_cell: bool = False
    bottom_left_cell: bool = False

    def to_val(self) -> str:
        """Serialise back to the 12-char OOXML ``w:val`` string."""
        bits = (
            self.first_row, self.last_row, self.first_column, self.last_column,
            self.odd_column_banding, self.even_column_banding,
            self.odd_row_banding, self.even_row_banding,
            self.top_right_cell, self.top_left_cell,
            self.bottom_right_cell, self.bottom_left_cell,
        )
        return "".join("1" if b else "0" for b in bits)

    @classmethod
    def from_val(cls, val: str) -> "ConditionalStyleMask":
        """Parse the 12-char ``w:val`` attribute."""
        if not val:
            return cls()
        padded = val.ljust(12, "0")
        return cls(
            first_row=padded[0] == "1",
            last_row=padded[1] == "1",
            first_column=padded[2] == "1",
            last_column=padded[3] == "1",
            odd_column_banding=padded[4] == "1",
            even_column_banding=padded[5] == "1",
            odd_row_banding=padded[6] == "1",
            even_row_banding=padded[7] == "1",
            top_right_cell=padded[8] == "1",
            top_left_cell=padded[9] == "1",
            bottom_right_cell=padded[10] == "1",
            bottom_left_cell=padded[11] == "1",
        )

    def __bool__(self) -> bool:
        return any((
            self.first_row, self.last_row, self.first_column, self.last_column,
            self.odd_column_banding, self.even_column_banding,
            self.odd_row_banding, self.even_row_banding,
            self.top_right_cell, self.top_left_cell,
            self.bottom_right_cell, self.bottom_left_cell,
        ))


# ─────────────────────────────────────────────
# Tab stops
# ─────────────────────────────────────────────


class TabStop(BaseModel):
    position: float = 0.0
    alignment: int = 0  # TabAlignment: 0=Left,1=Center,2=Right,3=Decimal,4=Bar,5=List,6=Clear
    leader: int = 0  # TabLeader: 0=None,1=Dot,2=Dash,3=Line,4=Heavy,5=MiddleDot
    is_clear: bool = False


class TabStopCollection(BaseModel):
    tab_stops: list[TabStop] = Field(default_factory=list)

    def clear(self) -> None:
        self.tab_stops.clear()

    def add(self, position: float, alignment: int = 0, leader: int = 0) -> None:
        self.tab_stops.append(TabStop(position=position, alignment=alignment, leader=leader))
        self.tab_stops.sort(key=lambda t: t.position)

    def remove_by_position(self, position: float) -> None:
        self.tab_stops = [t for t in self.tab_stops if abs(t.position - position) > 0.01]

    def before(self, pos: float) -> TabStop | None:
        result = None
        for t in self.tab_stops:
            if t.position < pos and not t.is_clear:
                result = t
        return result

    def after(self, pos: float) -> TabStop | None:
        for t in self.tab_stops:
            if t.position > pos and not t.is_clear:
                return t
        return None

    def __len__(self) -> int:
        return len(self.tab_stops)

    def __iter__(self):
        return iter(self.tab_stops)

    def __bool__(self) -> bool:
        return bool(self.tab_stops)


# ─────────────────────────────────────────────
# Font  (was 30+ fields → 15)
# ─────────────────────────────────────────────


FONT_BOOLEAN_FIELDS = (
    "bold", "italic", "hidden", "all_caps", "small_caps", "strike_through", "outline",
    "shadow", "emboss", "engrave", "no_proofing", "bold_bi", "italic_bi",
)


class SourceColor(BaseModel):
    """Immutable source declarations, cleared by direct color editing."""

    model_config = ConfigDict(frozen=True)
    value: str | None = None
    theme_color: str | None = None
    theme_tint: str | None = None
    theme_shade: str | None = None


class SourceThemePart(BaseModel):
    """Immutable reachable part, including its original OPC content type."""

    model_config = ConfigDict(frozen=True)
    name: str
    content_type: str
    data: bytes

    @field_serializer('data', when_used='json')
    def _encode_bytes(self, value: bytes) -> dict[str, str]:
        return {'encoding': 'base64', 'data': b64encode(value).decode('ascii')}

    @field_validator('data', mode='before')
    @classmethod
    def _decode_bytes(cls, value: Any) -> Any:
        if isinstance(value, dict) and value.get('encoding') == 'base64' and isinstance(value.get('data'), str):
            return b64decode(value['data'], validate=True)
        return value


class SourceTheme(BaseModel):
    """Original theme declarations and reachable package resources."""

    model_config = ConfigDict(frozen=True)
    data: bytes
    relationships: bytes | None = None
    part_name: str = 'word/theme/theme1.xml'
    parts: tuple[SourceThemePart, ...] = ()

    @field_serializer('data', 'relationships', when_used='json')
    def _encode_bytes(self, value: bytes | None) -> dict[str, str] | None:
        return None if value is None else {'encoding': 'base64', 'data': b64encode(value).decode('ascii')}

    @field_validator('data', 'relationships', mode='before')
    @classmethod
    def _decode_bytes(cls, value: Any) -> Any:
        if isinstance(value, dict) and value.get('encoding') == 'base64' and isinstance(value.get('data'), str):
            return b64decode(value['data'], validate=True)
        return value


class Font(BaseModel):
    name: str = ""
    size: float = 0.0
    size_explicit: bool | None = None
    bold: bool = False
    italic: bool = False
    bold_explicit: bool | None = None
    italic_explicit: bool | None = None
    underline: int = 0
    color: str = ""
    color_explicit: bool | None = None
    color_rendering: str | None = None
    source_color: SourceColor | None = None
    strike_through: bool = False
    superscript: bool = False
    subscript: bool = False
    highlight_color: str = ""
    all_caps: bool = False
    small_caps: bool = False
    hidden: bool = False
    hidden_explicit: bool | None = None
    hidden_rendering: bool | None = None
    all_caps_explicit: bool | None = None
    small_caps_explicit: bool | None = None
    strike_through_explicit: bool | None = None
    outline_explicit: bool | None = None
    shadow_explicit: bool | None = None
    emboss_explicit: bool | None = None
    engrave_explicit: bool | None = None
    no_proofing_explicit: bool | None = None
    bold_bi_explicit: bool | None = None
    italic_bi_explicit: bool | None = None
    style_name: str = ""
    style_identifier: int = 0
    shading: Shading = Field(default_factory=Shading)
    emboss: bool = False
    engrave: bool = False
    outline: bool = False
    shadow: bool = False
    # 0=None, 1=OverSolidCircle (dot), 2=OverComma, 3=OverWhiteCircle,
    # 4=UnderSolidCircle (underDot).
    emphasis_mark: int = 0
    # 0=None, 1=LasVegasLights, 2=BlinkingBackground, 3=SparkleText,
    # 4=MarchingBlackAnts, 5=MarchingRedAnts, 6=Shimmer.
    text_effect: int = 0
    # Minimum font size in points to apply kerning; 0 disables.
    kerning: float = 0.0
    bold_bi: bool = False
    italic_bi: bool = False
    no_proofing: bool = False
    name_bi: str = ""
    name_far_east: str = ""
    name_ascii: str = ""
    locale_id: int = 0
    locale_id_bi: int = 0
    locale_id_far_east: int = 0
    # REMOVED: size_bi, double_strike_through, underline_color, scaling,
    #          spacing, position, complex_script, border

    @property
    def render_hidden(self) -> bool:
        """Layout visibility can differ from the resolved property getter."""
        return self.hidden if self.hidden_rendering is None else self.hidden_rendering

    @property
    def render_color(self) -> str:
        """Layout color can differ from the original RGB/automatic getter."""
        return self.color if self.color_rendering is None else self.color_rendering

    def __setattr__(self, name: str, value: Any) -> None:
        if name == 'source_color' and value is not None and not isinstance(value, SourceColor):
            raise TypeError('source_color must be SourceColor or None')
        super().__setattr__(name, value)
        if name in (*FONT_BOOLEAN_FIELDS, 'color'):
            super().__setattr__(name + '_explicit', True)
        if name == 'hidden':
            super().__setattr__('hidden_rendering', None)
        if name == 'color':
            super().__setattr__('source_color', None)
        if name in ('color', 'source_color'):
            super().__setattr__('color_rendering', None)

    def model_copy(self, *, update: Mapping[str, Any] | None = None, deep: bool = False) -> Font:
        if update is not None and 'hidden' in update and 'hidden_rendering' not in update:
            update = dict(update, hidden_rendering=None)
        if update is not None and 'color' in update:
            update = dict(update)
            update.setdefault('color_rendering', None)
            update.setdefault('source_color', None)
        if update is not None:
            update = dict(update)
            for field in (*FONT_BOOLEAN_FIELDS, 'color'):
                if field in update and field + '_explicit' not in update:
                    update[field + '_explicit'] = True
        return super().model_copy(update=update, deep=deep)

    @model_validator(mode="wrap")
    @classmethod
    def _restore_fields_set(cls, data: Any, handler) -> Font:
        explicit = None
        if isinstance(data, dict) and "_fields_set" in data:
            data = data.copy()
            explicit = data.pop("_fields_set")
            if (not isinstance(explicit, list) or
                    any(not isinstance(name, str) or name not in cls.model_fields or name not in data
                        for name in explicit)):
                raise ValueError("Invalid font field-origin metadata")
        instance = handler(data)
        if explicit is not None:
            instance.__pydantic_fields_set__ = set(explicit)
        return instance

    @model_serializer(mode="wrap")
    def _emit_fields_set(self, handler) -> dict[str, Any]:
        data = handler(self)
        # Full JSON includes defaults; keep sparse declarations distinguishable.
        data["_fields_set"] = sorted(name for name in self.model_fields_set if name in data)
        return data


# ─────────────────────────────────────────────
# Paragraph format  (was 25+ fields → 18)
# ─────────────────────────────────────────────


class ParagraphFormat(BaseModel):
    style_name: str = ""
    style_explicit: bool | None = None

    alignment: int = 0  # 0=Left, 1=Center, 2=Right, 3=Justify
    left_indent: float = 0.0
    right_indent: float = 0.0
    first_line_indent: float = 0.0
    character_unit_left_indent: float | None = None
    character_unit_right_indent: float | None = None
    character_unit_first_line_indent: float | None = None
    space_before: float = 0.0
    space_after: float = 0.0
    space_before_auto: bool = False
    space_after_auto: bool = False
    line_spacing: float = 0.0
    line_spacing_rule: int = 0  # 0=AtLeast, 1=Exactly, 2=Multiple
    keep_with_next: bool = False
    page_break_before: bool = False
    no_space_between_paragraphs_of_same_style: bool = False
    outline_level: int = 9
    is_heading: bool = False
    is_list_item: bool = False
    shading: Shading = Field(default_factory=Shading)
    borders: list[Border] = Field(default_factory=list)
    paragraph_break_font: Optional[Font] = None
    style_identifier: int = 0
    keep_together: bool = False
    widow_control: bool = True
    suppress_auto_hyphens: bool = False
    suppress_line_numbers: bool = False
    snap_to_grid: bool = True
    add_space_between_far_east_and_alpha: bool = True
    add_space_between_far_east_and_digit: bool = True
    auto_adjust_right_indent: bool = True
    # 0=Auto, 1=Top, 2=Center, 3=Baseline, 4=Bottom.
    baseline_alignment: int = 0
    conditional_style: ConditionalStyleMask = Field(default_factory=ConditionalStyleMask)
    tab_stops: TabStopCollection = Field(default_factory=TabStopCollection)
    frame_format: Optional["FrameFormat"] = None
    lines_to_drop: int = 0
    # 0=None, 1=Normal, 2=Margin.
    drop_cap_position: int = 0

    # REMOVED: bidi, line_unit_*,
    #          far_east_line_break_control, word_wrap, hanging_punctuation,
    #          mirror_indents

    def __setattr__(self, name: str, value: Any) -> None:
        super().__setattr__(name, value)
        if name in ('style_name', 'style_identifier'):
            super().__setattr__('style_explicit', None)

    def model_copy(self, *, update: Mapping[str, Any] | None = None, deep: bool = False) -> ParagraphFormat:
        if update is not None and 'style_explicit' not in update and any(
                name in update for name in ('style_name', 'style_identifier')):
            update = dict(update, style_explicit=None)
        return super().model_copy(update=update, deep=deep)

    @model_validator(mode="before")
    @classmethod
    def _migrate_frame_overflow(cls, data: Any) -> Any:
        """Push legacy ``_frame_lock_anchor``/``_frame_wrap_type`` keys
        (from when these lived as PrivateAttrs on ParagraphFormat) into
        the ``frame_format`` dict so FrameFormat picks them up.
        """
        if not isinstance(data, dict):
            return data
        la = data.pop("_frame_lock_anchor", None)
        wt = data.pop("_frame_wrap_type", None)
        if la is not None or wt is not None:
            ff = data.get("frame_format")
            if ff is None:
                ff = {}
                data["frame_format"] = ff
            if isinstance(ff, dict):
                if la is not None and "anchor_locked" not in ff:
                    ff["anchor_locked"] = bool(la)
                if wt is not None and "wrap_type" not in ff:
                    ff["wrap_type"] = int(wt)
        return data


class FrameFormat(BaseModel):
    """Floating text-frame definition.  Numeric dimensions are in points."""

    width: float = 0.0
    height: float = 0.0
    # 0=AtLeast, 1=Exactly, 2=Auto.
    height_rule: int = 0
    horizontal_position: float = 0.0
    vertical_position: float = 0.0
    # 0=None, 1=Left, 2=Center, 3=Right, 4=Inside, 5=Outside.
    horizontal_alignment: int = 0
    # -1=Inline, 0=None, 1=Top, 2=Center, 3=Bottom, 4=Inside, 5=Outside.
    vertical_alignment: int = 0
    # 0=Margin, 1=Page, 3=Character.
    relative_horizontal_position: int = 0
    # 0=Margin, 1=Page, 2=Paragraph.
    relative_vertical_position: int = 0
    horizontal_distance_from_text: float = 0.0
    vertical_distance_from_text: float = 0.0
    # ShapeBase-equivalent fields also present on <w:framePr>:
    # 0=Inline, 1=TopBottom, 2=Square, 3=None, 4=Tight, 5=Through.
    wrap_type: int = 0
    anchor_locked: bool = False

    @model_validator(mode="before")
    @classmethod
    def _normalize_legacy_keys(cls, data: Any) -> Any:
        """Absorb legacy key names from earlier FrameFormat schemas."""
        if isinstance(data, dict):
            for old in ("lock_anchor", "_lock_anchor"):
                if old in data and "anchor_locked" not in data:
                    data["anchor_locked"] = data.pop(old)
                else:
                    data.pop(old, None)
            if "_wrap_type" in data and "wrap_type" not in data:
                data["wrap_type"] = data.pop("_wrap_type")
            else:
                data.pop("_wrap_type", None)
        return data

    @property
    def is_frame(self) -> bool:
        """True when any frame property has a non-default value."""
        return (
            self.width != 0.0
            or self.height != 0.0
            or self.horizontal_position != 0.0
            or self.vertical_position != 0.0
            or self.horizontal_alignment != 0
            or self.vertical_alignment != 0
        )


class ListFormat(BaseModel):
    is_list_item: bool = False
    list_level_number: int = 0
    list_id: int = 0


class ListLabel(BaseModel):
    """Snapshot of a list-item's rendered bullet/number label."""

    # Rendered label text (e.g. ``"1."`` / ``"-"`` / ``"a)"``).
    label_string: str = ""
    # Integer counter behind the label.
    label_value: int = 0
    font: Font = Field(default_factory=Font)


# ─────────────────────────────────────────────
# Image data
# ─────────────────────────────────────────────


class ImageData(BaseModel):
    source_full_name: str = ""
    image_type: int = _IT.NO_IMAGE
    image_bytes: bytes = b""
    crop_left: float = 0
    crop_top: float = 0
    crop_right: float = 0
    crop_bottom: float = 0

    @field_serializer("image_bytes", when_used="json")
    def _encode_image(self, value: bytes) -> dict[str, str]:
        # A tagged object keeps old UTF-8 strings distinguishable from binary data.
        return {"encoding": "base64", "data": b64encode(value).decode("ascii")}

    @field_validator("image_bytes", mode="before")
    @classmethod
    def _decode_image(cls, value: Any) -> Any:
        if isinstance(value, dict) and value.get("encoding") == "base64":
            data = value.get("data")
            if isinstance(data, str):
                return b64decode(data, validate=True)
        return value

    @property
    def content_type(self) -> str:
        """MIME content-type string derived from :attr:`image_type`."""
        return _IMAGE_TYPE_TO_MIME.get(self.image_type, "")

    @classmethod
    def from_mime(cls, mime: str) -> int:
        """Map a MIME content-type string to an ``ImageType`` constant."""
        return _MIME_TO_IMAGE_TYPE.get(mime.lower(), _IT.UNKNOWN)


# ─────────────────────────────────────────────
# Inline nodes
# ─────────────────────────────────────────────


class Run(BaseModel, NodeCastMixin):
    type: str = Field(default="Run", alias="_type")
    text: str = ""
    is_hyperlink: bool = False
    font: Font = Field(default_factory=Font)

    model_config = {"populate_by_name": True}


class Shape(BaseModel, NodeCastMixin):
    type: str = Field(default="Shape", alias="_type")
    shape_type: int | None = None
    name: str = ""
    # Aspose.Words: ShapeBase.AlternativeText.
    alternative_text: str = ""
    width: float | None = None
    height: float | None = None
    is_inline: bool | None = None
    has_image: bool | None = None
    image_data: Optional[ImageData] = None
    text_box: dict[str, Any] | None = None  # textbox paragraph content
    fill_color: str = ""
    stroke: Optional[Border] = None
    # Vertical anchor of the text-box content inside the shape's
    # bounding box: 0=Top (default), 1=Center, 2=Bottom.
    # Aspose.Words: TextBox.vertical_anchor (TextBoxAnchor enum).
    text_box_anchor: int = 0
    # WrapType (see drawing.WrapType):
    # 0=Inline, 1=TopBottom, 2=Square, 3=None, 4=Tight, 5=Through.
    wrap_type: int = 0  # WrapType.INLINE
    relative_horizontal_position: int = 0
    relative_vertical_position: int = 0
    left: float = 0.0
    top: float = 0.0
    horizontal_alignment: int = 0
    vertical_alignment: int = 0
    behind_text: bool = False
    allow_overlap: bool = True
    is_layout_in_cell: bool = True
    anchor_locked: bool = False
    # REMOVED: wrap_side, rotation, z_order

    # Runtime-only flag (not part of the JSON schema / not serialised)
    # set by the reader when the shape carries absolute page
    # coordinates extracted from an anchored group.  The PDF writer
    # uses it to branch into the absolute-positioning code path.
    _is_positioned: bool = PrivateAttr(default=False)
    _page_left_mm: float = PrivateAttr(default=0.0)
    _page_top_mm: float = PrivateAttr(default=0.0)

    model_config = {"populate_by_name": True}

    @model_validator(mode="before")
    @classmethod
    def _absorb_legacy_position(cls, data: Any) -> Any:
        """Accept legacy ``horizontal_position``/``vertical_position`` keys."""
        if isinstance(data, dict):
            if "horizontal_position" in data and "left" not in data:
                data["left"] = data.pop("horizontal_position")
            else:
                data.pop("horizontal_position", None)
            if "vertical_position" in data and "top" not in data:
                data["top"] = data.pop("vertical_position")
            else:
                data.pop("vertical_position", None)
        return data


class FieldStart(BaseModel):
    type: str = Field(default="FieldStart", alias="_type")
    field_type: int | None = None

    model_config = {"populate_by_name": True}


class FieldSeparator(BaseModel):
    type: str = Field(default="FieldSeparator", alias="_type")
    field_type: int | None = None

    model_config = {"populate_by_name": True}


class FieldEnd(BaseModel):
    type: str = Field(default="FieldEnd", alias="_type")
    field_type: int | None = None
    has_separator: bool = False

    model_config = {"populate_by_name": True}


class BookmarkStart(BaseModel):
    """Marks the beginning of a Word bookmark (``<w:bookmarkStart>``).

    The reader emits one of these for every named bookmark so the PDF
    writer can register the anchor's page + Y position, turning
    ``#name`` run-text references (hyperlinks, TOC entries) into real
    clickable internal links.
    """

    type: str = Field(default="BookmarkStart", alias="_type")
    name: str = ""

    model_config = {"populate_by_name": True}


class BookmarkEnd(BaseModel):
    type: str = Field(default="BookmarkEnd", alias="_type")
    name: str = ""

    model_config = {"populate_by_name": True}


class NoteReference(BaseModel):
    """An inline anchor; note bodies remain separate source stories."""

    type: Literal["NoteReference"] = Field(default="NoteReference", alias="_type")
    kind: Literal["footnote", "endnote"]
    identifier: str
    hidden: bool = False

    model_config = {"populate_by_name": True}

ChildNode = Union[Run, Shape, FieldStart, FieldSeparator, FieldEnd, BookmarkStart, BookmarkEnd, NoteReference]

_CHILD_NODE_CLASSES: tuple[type, ...] = (
    Run, Shape, FieldStart, FieldSeparator, FieldEnd, BookmarkStart, BookmarkEnd, NoteReference,
)
_CHILD_NODE_TYPE_TAGS = {"Run", "Shape", "FieldStart", "FieldSeparator", "FieldEnd", "BookmarkStart", "BookmarkEnd", "NoteReference"}


def _coerce_child_node(item: Any) -> ChildNode | None:
    """Materialise a dict-or-instance into a typed child node.

    Returns ``None`` for anything that isn't a known child kind (e.g.
    stale ``CommentNode`` payloads) — the caller drops these on the floor.
    """
    if isinstance(item, _CHILD_NODE_CLASSES):
        return item
    if isinstance(item, dict):
        t = item.get("_type", "")
        if t == "Run":
            return Run.model_validate(item)
        if t == "Shape":
            return Shape.model_validate(item)
        if t == "FieldStart":
            return FieldStart.model_validate(item)
        if t == "FieldSeparator":
            return FieldSeparator.model_validate(item)
        if t == "FieldEnd":
            return FieldEnd.model_validate(item)
        if t == "BookmarkStart":
            return BookmarkStart.model_validate(item)
        if t == "BookmarkEnd":
            return BookmarkEnd.model_validate(item)
        if t == "NoteReference":
            return NoteReference.model_validate(item)
    return None


# ─────────────────────────────────────────────
# Paragraph
# ─────────────────────────────────────────────


class SourceLocation(BaseModel):
    """Original parsed XML element indexes, independent of model edits and XML prefixes."""

    part_name: str
    child_path: list[int]


class Paragraph(BaseModel, NodeCastMixin):
    """A paragraph whose children — ``Run``, ``BookmarkStart`` / ``End``,
    ``FieldStart`` / ``Separator`` / ``End`` and inline ``Shape`` —
    sit in a single ordered collection in document order.

    The collection itself is intentionally not part of the public attribute
    surface: it lives on the private ``_children`` slot.  Read access goes
    through the typed view :attr:`runs` (the analogue of Aspose.Words'
    ``Paragraph.Runs``); writers inside the library mutate ``_children``
    directly.  In JSON the children are emitted under the ``children`` key.
    """

    type: Literal["Paragraph"] = Field(default="Paragraph", alias="_type")
    paragraph_format: ParagraphFormat = Field(default_factory=ParagraphFormat)
    list_format: ListFormat | None = None
    list_label: Optional[ListLabel] = None

    note_references: list[NoteReference] = Field(default_factory=list)
    source_location: SourceLocation | None = None
    _children: list[ChildNode] = PrivateAttr(default_factory=list)

    model_config = {"populate_by_name": True}

    @model_validator(mode="wrap")
    @classmethod
    def _absorb_children(cls, data: Any, handler) -> "Paragraph":
        """Pull the ``children`` array out of the input dict and install
        it on the private slot.  ``_text`` is accepted but ignored (text
        is computed from runs).
        """
        raw: Any = None
        if isinstance(data, dict):
            raw = data.pop("children", None)
            data.pop("_text", None)
        instance: "Paragraph" = handler(data)
        if raw:
            kept: list[ChildNode] = []
            for entry in raw:
                node = _coerce_child_node(entry)
                if node is not None:
                    kept.append(node)
            instance._children = kept
        return instance

    @model_serializer(mode="wrap")
    def _emit_children(self, handler, info) -> dict[str, Any]:
        """Emit ``children`` and ``_text`` in the serialised form."""
        data = handler(self)
        data["children"] = [c.model_dump(mode=info.mode, by_alias=True) for c in self._children]
        data["_text"] = self.text
        return data

    @property
    def text(self) -> str:
        """Concatenated text of all ``Run`` children."""
        return "".join(r.text for r in self._children if isinstance(r, Run))

    def get_text(self) -> str:
        """Text plus the paragraph mark, as ``aspose.words`` returns it."""
        return self.text + "\r"

    def get_child_nodes(self, node_type: int = NodeType.ANY, deep: bool = False) -> list:
        """Child nodes, optionally filtered by :class:`NodeType`."""
        return _collect_child_nodes(self, node_type, deep)

    @property
    def runs(self) -> list[Run]:
        """Typed view over the paragraph's ``Run`` children, in document
        order.  Direct counterpart of Aspose.Words' ``Paragraph.Runs``.
        """
        return [c for c in self._children if isinstance(c, Run)]

    @property
    def paragraph_break_font(self) -> Optional[Font]:
        """Aspose.Words exposes this on ``Paragraph``; storage lives on
        ``ParagraphFormat`` for style-chain merging.
        """
        return self.paragraph_format.paragraph_break_font

    @paragraph_break_font.setter
    def paragraph_break_font(self, value: Optional[Font]) -> None:
        self.paragraph_format.paragraph_break_font = value

    @property
    def frame_format(self) -> Optional["FrameFormat"]:
        """Aspose.Words exposes this on ``Paragraph``; storage lives on
        ``ParagraphFormat`` for style-chain merging.
        """
        return self.paragraph_format.frame_format


# ─────────────────────────────────────────────
# Table → Row → Cell
# ─────────────────────────────────────────────


class PreferredWidth(BaseModel):
    """Preferred width of a table / cell.

    Mirrors Aspose.Words' ``PreferredWidth`` with ``PreferredWidthType``.
    ``type`` values: ``PreferredWidthType.AUTO`` (0),
    ``PERCENT`` (1), ``POINTS`` (2).
    """

    type: int = _PWT.AUTO
    value: float = 0.0

    @property
    def is_auto(self) -> bool:
        return self.type == _PWT.AUTO

    @classmethod
    def auto(cls) -> "PreferredWidth":
        return cls(type=_PWT.AUTO, value=0.0)

    @classmethod
    def from_percent(cls, percent: float) -> "PreferredWidth":
        return cls(type=_PWT.PERCENT, value=percent)

    @classmethod
    def from_points(cls, points: float) -> "PreferredWidth":
        return cls(type=_PWT.POINTS, value=points)


class CellFormat(BaseModel):
    width: float = 0.0
    preferred_width: PreferredWidth = Field(default_factory=PreferredWidth)
    vertical_alignment: int = 0
    vertical_merge: int = 0  # 0=None, 1=First, 2=Previous
    horizontal_merge: int = 0  # 0=None, 1=First, 2=Previous
    grid_span: int = Field(default=1, ge=1, le=MAX_TABLE_COLUMNS)
    top_padding: Optional[float] = Field(default=None, ge=0, allow_inf_nan=False)
    bottom_padding: Optional[float] = Field(default=None, ge=0, allow_inf_nan=False)
    left_padding: Optional[float] = Field(default=None, ge=0, allow_inf_nan=False)
    right_padding: Optional[float] = Field(default=None, ge=0, allow_inf_nan=False)
    shading: Shading = Field(default_factory=Shading)
    borders: list[Border] = Field(default_factory=list)
    orientation: int = 0
    wrap_text: bool = True
    conditional_style: ConditionalStyleMask = Field(default_factory=ConditionalStyleMask)
    # REMOVED: fit_text


class Cell(BaseModel, NodeCastMixin):
    type: str = Field(default="Cell", alias="_type")
    cell_format: CellFormat = Field(default_factory=CellFormat)
    paragraphs: list[Paragraph] = Field(default_factory=list)
    tables: list[Table] = Field(default_factory=list)
    content_order: list[Literal["paragraph", "table"]] = Field(default_factory=list)

    model_config = {"populate_by_name": True}

    @model_validator(mode="before")
    @classmethod
    def _absorb_children(cls, data):
        if isinstance(data, dict) and "children" in data:
            data = dict(data)
            children = [_parse_body_child(child) for child in data.pop("children")]
            if any(not isinstance(child, (Paragraph, Table)) for child in children):
                raise ValueError("Cell children must be paragraphs or tables")
            data["paragraphs"] = [child for child in children if isinstance(child, Paragraph)]
            data["tables"] = [child for child in children if isinstance(child, Table)]
            data["content_order"] = ["paragraph" if isinstance(child, Paragraph) else "table"
                                     for child in children]
        return data

    @property
    def children(self) -> list[Paragraph | Table]:
        """Ordered content; legacy list constructors retain paragraph-then-table order."""
        paragraphs, tables = iter(self.paragraphs), iter(self.tables)
        children = []
        for kind in self.content_order:
            child = next(paragraphs if kind == "paragraph" else tables, None)
            if child is not None:
                children.append(child)
        children.extend(paragraphs)
        children.extend(tables)
        return children


class RowFormat(BaseModel):
    height: float = 0.0
    height_rule: int = 0
    heading_format: bool = False
    allow_break_across_pages: bool = True  # row split control
    conditional_style: ConditionalStyleMask = Field(default_factory=ConditionalStyleMask)
    borders: list[Border] = Field(default_factory=list)


class Row(BaseModel, NodeCastMixin):
    type: str = Field(default="Row", alias="_type")
    row_format: RowFormat = Field(default_factory=RowFormat)
    cells: list[Cell] = Field(default_factory=list)
    # Per-row table-width override (OOXML <w:tblPrEx><w:tblW/>).
    # Analog of Table.preferred_width, applied at row level.
    preferred_width: PreferredWidth = Field(default_factory=PreferredWidth)

    model_config = {"populate_by_name": True}

    @model_validator(mode="before")
    @classmethod
    def _migrate_tblPrEx(cls, data: Any) -> Any:
        """Accept legacy ``_tblPrEx_width`` and old
        ``row_format.preferred_width`` keys.
        """
        if not isinstance(data, dict):
            return data
        pw = data.pop("_tblPrEx_width", None)
        rf = data.get("row_format")
        if isinstance(rf, dict) and pw is None:
            pw = rf.pop("preferred_width", None)
        if pw is not None and "preferred_width" not in data:
            data["preferred_width"] = pw
        return data


def iter_grid_cells(row):
    """Yield (cell, column, span), combining legacy horizontal-merge continuations."""
    column, index = 0, 0
    while index < len(row.cells):
        cell = row.cells[index]
        span = cell.cell_format.grid_span
        index += 1
        if cell.cell_format.horizontal_merge == 1:
            children = list(cell.children)
            while index < len(row.cells) and row.cells[index].cell_format.horizontal_merge == 2:
                following = row.cells[index]
                span += following.cell_format.grid_span
                children.extend(following.children)
                index += 1
            cell = cell.model_copy(update={
                "paragraphs": [child for child in children if isinstance(child, Paragraph)],
                "tables": [child for child in children if isinstance(child, Table)],
                "content_order": ["paragraph" if isinstance(child, Paragraph) else "table"
                                  for child in children],
            })
        if column + span > MAX_TABLE_COLUMNS:
            raise ValueError("Table grid exceeds the column limit")
        yield cell, column, span
        column += span


class Table(BaseModel, NodeCastMixin):
    type: Literal["Table"] = Field(default="Table", alias="_type")
    source_location: SourceLocation | None = None
    alignment: int = 0
    preferred_width: PreferredWidth = Field(default_factory=PreferredWidth)
    left_indent: float = 0.0
    bidi: bool = False
    left_padding: Optional[float] = Field(default=None, ge=0, allow_inf_nan=False)
    right_padding: Optional[float] = Field(default=None, ge=0, allow_inf_nan=False)
    top_padding: Optional[float] = Field(default=None, ge=0, allow_inf_nan=False)
    bottom_padding: Optional[float] = Field(default=None, ge=0, allow_inf_nan=False)
    style_name: str = ""
    text_wrapping: int = 0  # 0=None, 1=Default/Around
    title: str = ""
    description: str = ""
    rows: list[Row] = Field(default_factory=list)
    # REMOVED: style_identifier, allow_auto_fit,
    #          allow_cell_spacing, cell_spacing

    _tblp_pr_attrs: dict[str, str] = PrivateAttr(default_factory=dict)

    model_config = {"populate_by_name": True}

    @property
    def first_row(self) -> Row | None:
        return self.rows[0] if self.rows else None

    @property
    def last_row(self) -> Row | None:
        return self.rows[-1] if self.rows else None


Cell.model_rebuild()


# ─────────────────────────────────────────────
# Body children
# ─────────────────────────────────────────────


class UnknownNode(BaseModel):
    type: str = Field(default="", alias="_type")
    model_config = {"extra": "allow", "populate_by_name": True}


def _parse_body_child(v: Any) -> Paragraph | Table | UnknownNode:
    if isinstance(v, (Paragraph, Table, UnknownNode)):
        return v
    if isinstance(v, dict):
        t = v.get("_type", "")
        if t == "Paragraph":
            return Paragraph.model_validate(v)
        if t == "Table":
            return Table.model_validate(v)
        # Infer type from structure when _type is absent
        if not t:
            if "rows" in v:
                return Table.model_validate(v)
            if "runs" in v or "paragraph_format" in v or "_text" in v:
                return Paragraph.model_validate(v)
        return UnknownNode.model_validate(v)
    raise ValueError(f"Cannot parse body child: {v!r}")


BodyChild = Annotated[
    Union[Paragraph, Table, UnknownNode],
    BeforeValidator(_parse_body_child),
]


# ─────────────────────────────────────────────
# Text columns
# ─────────────────────────────────────────────


class TextColumn(BaseModel):
    width: float = 0.0
    space_after: float = 0.0


class TextColumns(BaseModel):
    count: int = 1
    evenly_spaced: bool = True
    spacing: float = 36.0
    line_between: bool = False
    columns: list[TextColumn] = Field(default_factory=list)


# ─────────────────────────────────────────────
# Section structure
# ─────────────────────────────────────────────


class PageSetup(BaseModel):
    paper_size: int = 0
    # ``model.enums.Orientation``; unrelated to ``CellFormat.orientation``,
    # which is a text-direction enum.
    orientation: int = 1  # 1=Portrait, 2=Landscape
    top_margin: float = 70.85
    bottom_margin: float = 70.85
    left_margin: float = 70.85
    right_margin: float = 70.85
    header_distance: float = 35.4
    footer_distance: float = 35.4
    gutter: float = 0.0  # Aspose.Words: PageSetup.Gutter — extra binding margin
    page_width: float = 0.0
    page_height: float = 0.0
    page_number_style: int = 0  # Roman, Arabic, letters...
    page_starting_number: int = 1  # section starts at page N
    restart_page_numbering: bool = False  # reset counter at section
    section_start: int = 2  # 0=Continuous,1=NewColumn,2=NextPage,3=EvenPage,4=OddPage
    different_first_page_header_footer: bool = False
    odd_and_even_pages_header_footer: bool = False  # enable odd/even headers
    text_columns: Optional[TextColumns] = None
    # REMOVED: multiple_pages, sheets_per_booklet,
    #          vertical_alignment, line_number_*, bidi, borders


class HeaderFooter(BaseModel):
    type: str = Field(default="", alias="_type")
    header_footer_type: int | None = None
    children: list[BodyChild] = Field(default_factory=list)

    model_config = {"populate_by_name": True}

    @property
    def paragraphs(self) -> list[Paragraph]:
        return [c for c in self.children if isinstance(c, Paragraph)]

    @property
    def tables(self) -> list[Table]:
        return [c for c in self.children if isinstance(c, Table)]


class Body(BaseModel, NodeCastMixin):
    type: str = Field(default="Body", alias="_type")
    children: list[BodyChild] = Field(default_factory=list)

    model_config = {"populate_by_name": True}

    @property
    def paragraphs(self) -> list[Paragraph]:
        """Direct ``Paragraph`` children."""
        return [c for c in self.children if isinstance(c, Paragraph)]

    @property
    def tables(self) -> list[Table]:
        """Direct ``Table`` children."""
        return [c for c in self.children if isinstance(c, Table)]

    def get_child_nodes(self, node_type: int = NodeType.ANY, deep: bool = False) -> list:
        """Child nodes, optionally filtered by :class:`NodeType`."""
        return _collect_child_nodes(self, node_type, deep)


class Section(BaseModel, NodeCastMixin):
    type: str = Field(default="Section", alias="_type")
    page_setup: PageSetup = Field(default_factory=PageSetup)
    body: Body = Field(default_factory=Body)
    headers_footers: list[HeaderFooter] = Field(default_factory=list)

    model_config = {"populate_by_name": True}


# ─────────────────────────────────────────────
# Styles
# ─────────────────────────────────────────────


class TableStyleFormat(BaseModel):
    """Table-level properties stored on table styles (``w:tblPr`` inside ``w:style``)."""

    borders: list[Border] = Field(default_factory=list)
    bidi: Optional[bool] = None
    left_padding: Optional[float] = Field(default=None, ge=0, allow_inf_nan=False)
    right_padding: Optional[float] = Field(default=None, ge=0, allow_inf_nan=False)
    top_padding: Optional[float] = Field(default=None, ge=0, allow_inf_nan=False)
    bottom_padding: Optional[float] = Field(default=None, ge=0, allow_inf_nan=False)


class TableStyleProperty(BaseModel):
    """Conditional formatting for a table region (``w:tblStylePr``)."""

    type: str = ""
    font: Optional[Font] = None
    shading: Optional[Shading] = None
    borders: list[Border] = Field(default_factory=list)
    paragraph_format: Optional["ParagraphFormat"] = None


class Style(BaseModel):
    name: str = ""
    is_default: bool | None = None  # None keeps legacy built-in Normal inference.
    type: int = 0  # 1=paragraph, 2=character, 3=table
    is_heading: bool = False
    base_style_name: str = ""
    next_paragraph_style_name: str = ""
    paragraph_format: ParagraphFormat | None = None
    font: Font | None = None
    table_style_format: TableStyleFormat | None = None
    table_style_properties: list[TableStyleProperty] = Field(default_factory=list)
    style_identifier: int = 0
    built_in: bool = False
    priority: int = 99
    semi_hidden: bool = False
    unhide_when_used: bool = False
    locked: bool = False
    # REMOVED: is_quick_style, linked_style_name, aliases, automatically_update


# ─────────────────────────────────────────────
# Lists
# ─────────────────────────────────────────────


class ListLevel(BaseModel):
    number_format: str = ""
    number_style: int = 0
    start_at: int = 1
    alignment: int = 0
    number_position: float = 0.0
    text_position: float = 0.0
    character_unit_left_indent: float | None = None
    character_unit_right_indent: float | None = None
    character_unit_first_line_indent: float | None = None
    # ``<w:lvl><w:rPr>`` — formatting of the bullet / number glyph itself
    # (font name, size, italic, color).  Distinct from the paragraph
    # mark font carried on each individual list-item paragraph.
    font: Optional[Font] = None
    # Restart this level's counter when a higher level resets past N;
    # -1 disables restart entirely.
    restart_after_level: int = -1
    # 0=Tab (default), 1=Space, 2=Nothing.
    trailing_character: int = 0


class ListLevelOverride(BaseModel):
    """One ``<w:lvlOverride>`` inside a concrete ``<w:num>``."""

    ilvl: int = 0
    start_at: Optional[int] = None
    list_level: Optional[ListLevel] = None


class DocList(BaseModel):
    list_id: int = 0
    is_multi_level: bool = False
    list_levels: list[ListLevel] = Field(default_factory=list)
    # ``<w:lvlOverride>`` entries on the concrete ``<w:num>`` element.
    overrides: list[ListLevelOverride] = Field(default_factory=list)
    # REMOVED: is_list_style_definition, is_list_style_reference,
    #          is_restart_at_each_section

    @model_validator(mode="before")
    @classmethod
    def _migrate_levels_key(cls, data: Any) -> Any:
        """Accept legacy ``levels`` key (renamed to ``list_levels``)."""
        if isinstance(data, dict):
            if "levels" in data and "list_levels" not in data:
                data["list_levels"] = data.pop("levels")
        return data

    @model_validator(mode="after")
    def _validate_overrides(self) -> "DocList":
        """Aspose normalization: if any override carries ``start_at``,
        every override must have one (default 0 for those that didn't).
        """
        if not self.overrides:
            return self
        if any(ov.start_at is not None for ov in self.overrides):
            for ov in self.overrides:
                if ov.start_at is None:
                    ov.start_at = 0
        return self


# ─────────────────────────────────────────────
# Root Document
# ─────────────────────────────────────────────

# REMOVED entirely (17 classes):
#   CompatibilityOptions, EndnoteOptions, FootnoteOptions,
#   HyphenationOptions, ViewOptions, WriteProtection, Watermark,
#   Theme, ThemeColors, MailMergeSettings, FontInfo,
#   BuiltInProperties, TextColumns


class SourceStory(BaseModel):
    """Extracted source content, not a promise that conversion writers can render it."""

    kind: Literal["footnote", "endnote", "header", "footer"]
    part_name: str
    identifier: str | None = None
    references: list[dict[str, Any]] = Field(default_factory=list)
    children: list[BodyChild] = Field(default_factory=list)


class Document(BaseModel):
    type: str = Field(default="Document", alias="_type")
    default_tab_stop: float = 36.0
    do_not_expand_shift_return: bool = False
    compatibility_mode: int = Field(default=15, ge=0)
    page_color: str = ""
    page_count: int = 0
    source_theme: SourceTheme | None = None
    doc_defaults_font: Optional[Font] = None
    doc_defaults_rpr_present: bool | None = None

    styles: list[Style] = Field(default_factory=list)
    lists: list[DocList] = Field(default_factory=list)
    sections: list[Section] = Field(default_factory=list)
    source_stories: list[SourceStory] = Field(default_factory=list)

    model_config = {"populate_by_name": True}

    def get_child_nodes(self, node_type: int = NodeType.ANY, deep: bool = False) -> list:
        """Child nodes, optionally filtered by :class:`NodeType`."""
        return _collect_child_nodes(self, node_type, deep)

    @model_validator(mode="wrap")
    @classmethod
    def _absorb_legacy_hf(cls, data: Any, handler) -> "Document":
        """Absorb legacy ``header_paragraphs`` / ``footer_paragraphs``
        from old serialized LDMs into per-section HeaderFooter objects.
        """
        hdr_raw = ftr_raw = None
        if isinstance(data, dict):
            hdr_raw = data.pop("header_paragraphs", None)
            ftr_raw = data.pop("footer_paragraphs", None)
        instance: "Document" = handler(data)
        if (hdr_raw or ftr_raw) and instance.sections:
            sec = instance.sections[0]
            existing_types = {hf.header_footer_type for hf in sec.headers_footers}
            if hdr_raw and 0 not in existing_types:
                paras = [Paragraph.model_validate(p) if isinstance(p, dict) else p for p in hdr_raw]
                sec.headers_footers.append(HeaderFooter(header_footer_type=0, children=paras))
            if ftr_raw and 1 not in existing_types:
                paras = [Paragraph.model_validate(p) if isinstance(p, dict) else p for p in ftr_raw]
                sec.headers_footers.append(HeaderFooter(header_footer_type=1, children=paras))
        return instance

    @model_serializer(mode="wrap")
    def _emit_legacy_hf(self, handler, info) -> dict[str, Any]:
        """Emit ``header_paragraphs`` / ``footer_paragraphs`` in the
        serialized form for backward compatibility.
        """
        data = handler(self)
        data["header_paragraphs"] = [p.model_dump(mode=info.mode, by_alias=True) for p in self.header_paragraphs]
        data["footer_paragraphs"] = [p.model_dump(mode=info.mode, by_alias=True) for p in self.footer_paragraphs]
        return data

    @property
    def header_paragraphs(self) -> list[Paragraph]:
        """Flat list of header paragraphs collected from all sections."""
        result: list[Paragraph] = []
        for sec in self.sections:
            for hf in sec.headers_footers:
                if hf.header_footer_type == 0:
                    result.extend(hf.paragraphs)
        return result

    @property
    def footer_paragraphs(self) -> list[Paragraph]:
        """Flat list of footer paragraphs collected from all sections."""
        result: list[Paragraph] = []
        for sec in self.sections:
            for hf in sec.headers_footers:
                if hf.header_footer_type == 1:
                    result.extend(hf.paragraphs)
        return result

    # REMOVED root fields: node_type, attached_template,
    #   automatically_update_styles, compliance, custom_node_id,
    #   grammar_checked, has_macros, has_revisions, justification_mode,
    #   original_load_format, protection_type, punctuation_kerning,
    #   remove_personal_information, revisions_view, shade_form_data,
    #   show_grammatical_errors, show_spelling_errors, spelling_checked,
    #   track_revisions, versions_count,
    #   include_textboxes_footnotes_endnotes_in_stat,
    #   compatibility_options, endnote_options, footnote_options,
    #   hyphenation_options, view_options, write_protection, watermark,
    #   theme, mail_merge_settings, font_infos, built_in_properties

    # ── convenience helpers ──

    @property
    def all_paragraphs(self) -> list[Paragraph]:
        """Flat list of direct body paragraphs only (excludes table cells)."""
        result: list[Paragraph] = []
        for sec in self.sections:
            for child in sec.body.children:
                if isinstance(child, Paragraph):
                    result.append(child)
        return result

    @property
    def tables(self) -> list[Table]:
        """Flat list of all top-level tables in the body."""
        return [
            child
            for sec in self.sections
            for child in sec.body.children
            if isinstance(child, Table)
        ]

    @property
    def all_tables(self) -> list[Table]:
        """Alias for tables (backwards compatibility)."""
        return self.tables

    @property
    def text(self) -> str:
        """Body text, including nested table cells in content order; no source stories."""
        def paragraphs(children):
            for child in children:
                if isinstance(child, Paragraph):
                    yield child
                elif isinstance(child, Table):
                    for row in child.rows:
                        for cell in row.cells:
                            yield from paragraphs(cell.children)
        from aspose.words_foss.pdf_writer.text import source_plain_text as plain_text

        return "\n".join(text for section in self.sections
                         for p in paragraphs(section.body.children) if (text := plain_text(p)))

    def get_list(self, list_id: int) -> DocList | None:
        """Look up a list definition by ID."""
        for dl in self.lists:
            if dl.list_id == list_id:
                return dl
        return None

    def resolve_list_level(self, list_id: int, level_num: int) -> ListLevel | None:
        """Resolve the effective ``ListLevel`` for a list-id + level pair,
        applying any ``<w:lvlOverride>`` formatting from the concrete
        ``<w:num>`` element.
        """
        dl = self.get_list(list_id)
        if dl is None or not dl.list_levels:
            return None
        idx = min(level_num, len(dl.list_levels) - 1)
        base = dl.list_levels[idx]
        for ov in dl.overrides:
            if ov.ilvl != level_num:
                continue
            if ov.list_level is not None:
                return ov.list_level
            if ov.start_at is not None:
                return base.model_copy(update={"start_at": ov.start_at})
        return base

    def find_style(self, name: str) -> Style | None:
        """Look up a style by its exact name."""
        for s in self.styles:
            if s.name == name:
                return s
        return None

    def headings(self, max_level: int = 9) -> list[Paragraph]:
        """All heading paragraphs up to a given outline level."""
        return [
            p
            for p in self.all_paragraphs
            if p.paragraph_format.is_heading and p.paragraph_format.outline_level < max_level
        ]
