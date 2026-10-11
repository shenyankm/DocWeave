"""Editable document and section note configuration."""
from enum import IntEnum

from aspose.words_foss.dom.notes import FootnoteType
from aspose.words_foss.model.enums import NumberStyle

__all__ = ["EndnoteOptions", "EndnotePosition", "FootnoteNumberingRule", "FootnoteOptions", "FootnotePosition", "FootnoteType"]

class FootnoteNumberingRule(IntEnum):
    CONTINUOUS = 0
    RESTART_SECTION = 1
    RESTART_PAGE = 2


class FootnotePosition(IntEnum):
    BOTTOM_OF_PAGE = 1
    BENEATH_TEXT = 2


class EndnotePosition(IntEnum):
    END_OF_SECTION = 0
    END_OF_DOCUMENT = 3


_NUMBER_TOKENS = {0: 'decimal', 1: 'upperRoman', 2: 'lowerRoman', 3: 'upperLetter', 4: 'lowerLetter', 5: 'ordinal', 6: 'cardinalText', 7: 'ordinalText', 8: 'hex', 9: 'chicago', 10: 'ideographDigital', 11: 'japaneseCounting', 12: 'aiueo', 13: 'iroha', 14: 'decimalFullWidth', 15: 'decimalHalfWidth', 16: 'japaneseLegal', 17: 'japaneseDigitalTenThousand', 18: 'decimalEnclosedCircle', 19: 'decimalFullWidth2', 20: 'aiueoFullWidth', 21: 'irohaFullWidth', 22: 'decimalZero', 23: 'bullet', 24: 'ganada', 25: 'chosung', 26: 'decimalEnclosedFullstop', 27: 'decimalEnclosedParen', 28: 'decimalEnclosedCircleChinese', 29: 'ideographEnclosedCircle', 30: 'ideographTraditional', 31: 'ideographZodiac', 32: 'ideographZodiacTraditional', 33: 'taiwaneseCounting', 34: 'ideographLegalTraditional', 35: 'taiwaneseCountingThousand', 36: 'taiwaneseDigital', 37: 'chineseCounting', 38: 'chineseLegalSimplified', 39: 'chineseCountingThousand', 41: 'koreanDigital', 42: 'koreanCounting', 43: 'koreanLegal', 44: 'koreanDigital2', 45: 'hebrew1', 46: 'arabicAlpha', 47: 'hebrew2', 48: 'arabicAbjad', 49: 'hindiVowels', 50: 'hindiConsonants', 51: 'hindiNumbers', 52: 'hindiCounting', 53: 'thaiLetters', 54: 'thaiNumbers', 55: 'thaiCounting', 56: 'vietnameseCounting', 57: 'numberInDash', 58: 'russianLower', 59: 'russianUpper', 255: 'none', 65280: 'custom'}
_POSITION_TOKENS = {0: 'sectEnd', 1: 'pageBottom', 2: 'beneathText', 3: 'docEnd'}
_RESTART_TOKENS = {0: 'continuous', 1: 'eachSect', 2: 'eachPage'}
_FIELDS = {'number_style': 'numFmt', 'position': 'pos', 'restart_rule': 'numRestart', 'start_number': 'numStart'}


def _property(properties, kind):
    found = [p for p in properties if p.kind == kind]
    if len(found) > 1:
        raise ValueError('Duplicate note properties')
    return found[0] if found else None


def _get(properties, defaults, kind, field):
    prop = _property(properties, kind)
    if prop is not None:
        if field == 'columns' and prop.columns is not None:
            return prop.columns
        if field == 'number_style' and prop.number_style_override is not None:
            return NumberStyle(prop.number_style_override)
        found = [p for p in prop.children if p.kind == _FIELDS.get(field)]
        if found:
            value = found[-1].value
            if field == 'start_number':
                from aspose.words_foss._notes import note_id
                return int(note_id(value))
            if field == 'number_style':
                return NumberStyle(next((n for n, token in _NUMBER_TOKENS.items() if token == value), 0))
            if field == 'restart_rule':
                return FootnoteNumberingRule(next((n for n, token in _RESTART_TOKENS.items() if token == value), 0))
            if field == 'position':
                cls = FootnotePosition if kind == 'footnote' else EndnotePosition
                return cls(next((n for n, token in _POSITION_TOKENS.items() if token == value and n in ({1, 2} if kind == 'footnote' else {0, 3})), 3 if kind == 'endnote' else 1))
    if defaults is not None:
        return _get(defaults, None, kind, field)
    if field == 'number_style':
        return NumberStyle.ARABIC if kind == 'footnote' else NumberStyle.LOWERCASE_ROMAN
    if field == 'position':
        return FootnotePosition.BOTTOM_OF_PAGE if kind == 'footnote' else EndnotePosition.END_OF_DOCUMENT
    if field == 'restart_rule':
        return FootnoteNumberingRule.CONTINUOUS
    return 0 if field == 'columns' else 1


def _coerce(kind, field, value):
    cls = {'number_style': NumberStyle, 'restart_rule': FootnoteNumberingRule,
           'position': FootnotePosition if kind == 'footnote' else EndnotePosition}.get(field)
    if cls is not None:
        if type(value) is not cls:
            raise TypeError(f"can't build {cls.__name__} value from '{type(value).__name__}'")
        if value not in cls:
            raise ValueError('Invalid note option enum value')
    else:
        if type(value) is not int:
            raise TypeError('Note option requires an integer')
        if not -(1 << 31) <= value < (1 << 31):
            raise OverflowError('Note option integer is outside signed 32-bit range')
    return value


_PROPERTY_ORDER = {name: rank for rank, name in enumerate(('pos', 'numFmt', 'numStart', 'numRestart', 'footnote', 'endnote'))}


def _set(properties, kind, field, value):
    from aspose.words_foss import light_document_model as ldm
    prop = _property(properties, kind)
    if prop is None:
        prop = ldm.NoteProperties(kind=kind)
        properties.append(prop)
    if field == 'columns':
        prop.columns = value
        return
    if field == 'number_style':
        prop.number_style_override = int(value) if value == NumberStyle.SIMP_CHIN_NUM4 else None
        token = _NUMBER_TOKENS.get(int(value))
    elif field == 'position':
        token = _POSITION_TOKENS[int(value)]
    elif field == 'restart_rule':
        token = _RESTART_TOKENS[int(value)]
    else:
        token = str(value)
    tag = _FIELDS[field]
    children = [p for p in prop.children if p.kind != tag]
    if token is not None:
        index = next((i for i, child in enumerate(children)
                      if _PROPERTY_ORDER[child.kind] > _PROPERTY_ORDER[tag]), len(children))
        children.insert(index, ldm.NoteProperty(kind=tag, value=token))
    prop.children = children


def bind_note_options(document):
    targets = (document.note_properties, *(section.page_setup.note_properties for section in document.sections))
    for section in document.sections:
        section.page_setup._note_defaults = document.note_properties
        section.page_setup._note_targets = targets


class _NoteOptions:
    __slots__ = ('_dom', '_kind', '_owner')

    def __init__(self, *args, **kwargs):
        raise TypeError('type has no available constructor')

    @classmethod
    def _for_owner(cls, owner, kind, *, dom=False):
        result = object.__new__(cls)
        result._owner, result._kind, result._dom = owner, kind, dom
        return result

    def _get(self, field):
        if self._dom:
            return self._owner.get(self._kind, field)
        return _get(self._owner.note_properties, getattr(self._owner, '_note_defaults', None), self._kind, field)

    def _set(self, field, value):
        value = _coerce(self._kind, field, value)
        if self._dom:
            self._owner.set(self._kind, field, value)
            return
        # Native endnote position setters apply globally, even through PageSetup.
        targets = getattr(self._owner, '_note_targets', None) if self._kind == 'endnote' and field == 'position' else None
        if targets is None:
            targets = [self._owner.note_properties] + [s.page_setup.note_properties for s in getattr(self._owner, 'sections', [])]
        for properties in targets:
            _property(properties, self._kind)
        for properties in targets:
            _set(properties, self._kind, field, value)
        if hasattr(self._owner, 'sections'):
            bind_note_options(self._owner)

    def get_type(self):
        raise NotImplementedError('CLR reflection is not available in the Python implementation')

    number_style = property(lambda self: self._get('number_style'), lambda self, value: self._set('number_style', value))
    position = property(lambda self: self._get('position'), lambda self, value: self._set('position', value))
    restart_rule = property(lambda self: self._get('restart_rule'), lambda self, value: self._set('restart_rule', value))
    start_number = property(lambda self: self._get('start_number'), lambda self, value: self._set('start_number', value))


class FootnoteOptions(_NoteOptions):
    __slots__ = ()
    columns = property(lambda self: self._get('columns'), lambda self, value: self._set('columns', value))


class EndnoteOptions(_NoteOptions):
    __slots__ = ()


FOOTNOTE_COLUMNS_NS = 'http://schemas.microsoft.com/office/word/2012/wordml'


def read_columns(root, properties):
    from aspose.words_foss._notes import note_id
    from aspose.words_foss.docx_reader.constants import W_NS
    nodes = root.findall('{' + FOOTNOTE_COLUMNS_NS + '}footnoteColumns')
    if not nodes:
        return
    prop = _property(properties, 'footnote')
    if prop is None:
        from aspose.words_foss import light_document_model as ldm
        prop = ldm.NoteProperties(kind='footnote')
        properties.append(prop)
    prop.columns = max(0, int(note_id(nodes[-1].get(W_NS + 'val', ''))))


def render_columns(properties, compatibility_mode):
    prop = _property(properties, 'footnote')
    if compatibility_mode is None or compatibility_mode < 15 or prop is None or not prop.columns or prop.columns < 0:
        return None
    from aspose.words_foss.docx_writer.xml_utils import el
    return el('w15:footnoteColumns', {'xmlns:w15': FOOTNOTE_COLUMNS_NS, 'w:val': str(prop.columns)})
