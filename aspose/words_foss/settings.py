"""Document compatibility preferences and calibrated Word-version profiles."""
from enum import IntEnum
from xml.etree import ElementTree as ET
from defusedxml.ElementTree import fromstring

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WORD_URI = 'http://schemas.microsoft.com/office/word'

class MsWordVersion(IntEnum):
    WORD2000 = 0
    WORD2002 = 1
    WORD2003 = 2
    WORD2007 = 3
    WORD2010 = 4
    WORD2013 = 5
    WORD2016 = 6
    WORD2019 = 7

_COMPAT_TAGS = {'adjust_line_height_in_table': 'adjustLineHeightInTable',
 'align_tables_row_by_row': 'alignTablesRowByRow',
 'allow_space_of_same_style_in_table': 'allowSpaceOfSameStyleInTable',
 'apply_breaking_rules': 'applyBreakingRules',
 'auto_space_like_word95': 'autoSpaceLikeWord95',
 'autofit_to_first_fixed_width_cell': 'autofitToFirstFixedWidthCell',
 'balance_single_byte_double_byte_width': 'balanceSingleByteDoubleByteWidth',
 'cached_col_balance': 'cachedColBalance',
 'conv_mail_merge_esc': 'convMailMergeEsc',
 'display_hangul_fixed_width': 'displayHangulFixedWidth',
 'do_not_autofit_constrained_tables': 'doNotAutofitConstrainedTables',
 'do_not_break_constrained_forced_table': 'doNotBreakConstrainedForcedTable',
 'do_not_break_wrapped_tables': 'doNotBreakWrappedTables',
 'do_not_expand_shift_return': 'doNotExpandShiftReturn',
 'do_not_leave_backslash_alone': 'doNotLeaveBackslashAlone',
 'do_not_snap_to_grid_in_cell': 'doNotSnapToGridInCell',
 'do_not_suppress_indentation': 'doNotSuppressIndentation',
 'do_not_suppress_paragraph_borders': 'doNotSuppressParagraphBorders',
 'do_not_use_east_asian_break_rules': 'doNotUseEastAsianBreakRules',
 'do_not_use_html_paragraph_auto_spacing': 'doNotUseHTMLParagraphAutoSpacing',
 'do_not_use_indent_as_numbering_tab_stop': 'doNotUseIndentAsNumberingTabStop',
 'do_not_vert_align_cell_with_sp': 'doNotVertAlignCellWithSp',
 'do_not_vert_align_in_txbx': 'doNotVertAlignInTxbx',
 'do_not_wrap_text_with_punct': 'doNotWrapTextWithPunct',
 'footnote_layout_like_ww8': 'footnoteLayoutLikeWW8',
 'forget_last_tab_alignment': 'forgetLastTabAlignment',
 'grow_autofit': 'growAutofit',
 'layout_raw_table_width': 'layoutRawTableWidth',
 'layout_table_rows_apart': 'layoutTableRowsApart',
 'line_wrap_like_word6': 'lineWrapLikeWord6',
 'mw_small_caps': 'mwSmallCaps',
 'no_column_balance': 'noColumnBalance',
 'no_extra_line_spacing': 'noExtraLineSpacing',
 'no_leading': 'noLeading',
 'no_space_raise_lower': 'noSpaceRaiseLower',
 'no_tab_hang_ind': 'noTabHangInd',
 'print_body_text_before_header': 'printBodyTextBeforeHeader',
 'print_col_black': 'printColBlack',
 'select_fld_with_first_or_last_char': 'selectFldWithFirstOrLastChar',
 'shape_layout_like_ww8': 'shapeLayoutLikeWW8',
 'show_breaks_in_frames': 'showBreaksInFrames',
 'space_for_ul': 'spaceForUL',
 'spacing_in_whole_points': 'spacingInWholePoints',
 'split_pg_break_and_para_mark': 'splitPgBreakAndParaMark',
 'sub_font_by_size': 'subFontBySize',
 'suppress_bottom_spacing': 'suppressBottomSpacing',
 'suppress_sp_bf_after_pg_brk': 'suppressSpBfAfterPgBrk',
 'suppress_spacing_at_top_of_page': 'suppressSpacingAtTopOfPage',
 'suppress_top_spacing': 'suppressTopSpacing',
 'suppress_top_spacing_wp': 'suppressTopSpacingWP',
 'swap_borders_facing_pgs': 'swapBordersFacingPages',
 'truncate_font_heights_like_wp6': 'truncateFontHeightsLikeWP6',
 'ul_trail_space': 'ulTrailSpace',
 'underline_tab_in_num_list': 'underlineTabInNumList',
 'use_alt_kinsoku_line_break_rules': 'useAltKinsokuLineBreakRules',
 'use_ansi_kerning_pairs': 'useAnsiKerningPairs',
 'use_fe_layout': 'useFELayout',
 'use_normal_style_for_list': 'useNormalStyleForList',
 'use_printer_metrics': 'usePrinterMetrics',
 'use_single_borderfor_contiguous_cells': 'useSingleBorderforContiguousCells',
 'use_word2002_table_style_rules': 'useWord2002TableStyleRules',
 'use_word97_line_break_rules': 'useWord97LineBreakRules',
 'wp_justification': 'wpJustification',
 'wp_space_width': 'wpSpaceWidth',
 'wrap_trail_spaces': 'wrapTrailSpaces'}
_MODERN_FLAGS = (
    'disable_open_type_font_formatting_features',
    'override_table_style_font_size_and_justification',
    'swap_inside_and_outside_for_mirror_indents_and_relative_positioning',
    'use_word2010_table_style_rules',
)
_FLAG_NAMES = frozenset(_COMPAT_TAGS) | frozenset(_MODERN_FLAGS) | {'ui_compat_97_to_2003', 'transparent_metafiles'}


def _source_compat(document):
    source = document.source_settings
    if source is not None:
        root = fromstring(source.data, forbid_dtd=True) if source.data is not None else ET.Element(W+'compat')
        if root.tag != W+'compat':
            raise ValueError('Source compatibility XML must contain w:compat')
        return root
    from aspose.words_foss.docx_writer.blank_template import settings_compat
    return fromstring('<root xmlns:w="'+W[1:-1]+'">'+settings_compat()+'</root>', forbid_dtd=True)[0]


def _source_flags(document):
    from aspose.words_foss.docx_reader.utils import parse_onoff
    root = _source_compat(document)
    result = {}
    for name,tag in _COMPAT_TAGS.items():
        declarations = root.findall(W+tag)
        result[name] = parse_onoff(declarations[-1]) if declarations else False
    ui = document.source_settings.ui_compat_97_to_2003 if document.source_settings is not None else None
    result['ui_compat_97_to_2003'] = ui is True
    result['transparent_metafiles'] = False
    # Native legacy declarations select the older defaults even with w:val="0".
    modern_default = ui is None and not any(root.find(W+tag) is not None for tag in _COMPAT_TAGS.values())
    result.update(dict.fromkeys(_MODERN_FLAGS, modern_default))
    return result


class CompatibilityOptions:
    """A live view of a document's compatibility preferences."""

    __slots__ = ('_owner',)

    def __init__(self, *args, **kwargs):
        raise TypeError('type has no available constructor')

    @classmethod
    def _for_document(cls, owner):
        result = object.__new__(cls)
        result._owner = owner
        return result

    def _get(self, name):
        document = self._owner.light_document_model
        if name == 'do_not_expand_shift_return':
            return document.do_not_expand_shift_return
        state = document.compatibility_options_state
        return state.flags[name] if state is not None else _source_flags(document)[name]

    def _state(self):
        from aspose.words_foss.light_document_model import CompatibilityOptionsState
        doc = self._owner.light_document_model
        if doc.compatibility_options_state is None:
            flags = _source_flags(doc)
            flags['do_not_expand_shift_return'] = doc.do_not_expand_shift_return
            doc.compatibility_options_state = CompatibilityOptionsState(flags=flags)
        return doc.compatibility_options_state

    def _set(self, name, value):
        if type(value) is not bool:
            raise TypeError('Compatibility option requires a bool')
        state = self._state()
        state.flags[name] = value
        state.edited_flags.add(name)
        if name == 'do_not_expand_shift_return':
            self._owner.light_document_model.do_not_expand_shift_return = value

    def optimize_for(self, version):
        """Apply the measured compatibility profile for a Microsoft Word version."""
        if not isinstance(version, MsWordVersion):
            raise TypeError('optimize_for requires MsWordVersion')
        state = self._state()
        previous = dict(state.flags)
        enabled = frozenset(_PROFILE_TRUE[int(version)])
        implicit_flags = _source_flags(self._owner.light_document_model)
        for name in _FLAG_NAMES:
            if name == 'ul_trail_space':
                continue
            # The four implicit newer defaults survive modern optimization; explicit edits do not.
            if name in _MODERN_FLAGS and int(version) >= 3 and name not in state.edited_flags:
                state.flags[name] = implicit_flags[name]
                continue
            state.flags[name] = name in enabled
        state.optimized_for = int(version)
        state.font_size_override = int(version) >= 4
        state.edited_flags.update(name for name in _FLAG_NAMES
                                 if name not in _MODERN_FLAGS and previous[name] != state.flags[name])
        doc = self._owner.light_document_model
        doc.do_not_expand_shift_return = state.flags['do_not_expand_shift_return']
        doc.compatibility_mode = (12, 12, 11, 12, 14, 15, 15, 15)[int(version)]


def _flag_property(name):
    return property(lambda self:self._get(name), lambda self,value:self._set(name,value))


for _name in _FLAG_NAMES:
    setattr(CompatibilityOptions, _name, _flag_property(_name))
_PROFILE_TRUE = (('allow_space_of_same_style_in_table',
  'autofit_to_first_fixed_width_cell',
  'cached_col_balance',
  'display_hangul_fixed_width',
  'do_not_autofit_constrained_tables',
  'do_not_break_constrained_forced_table',
  'do_not_break_wrapped_tables',
  'do_not_snap_to_grid_in_cell',
  'do_not_suppress_indentation',
  'do_not_use_east_asian_break_rules',
  'do_not_use_indent_as_numbering_tab_stop',
  'do_not_vert_align_cell_with_sp',
  'do_not_vert_align_in_txbx',
  'do_not_wrap_text_with_punct',
  'grow_autofit',
  'select_fld_with_first_or_last_char',
  'split_pg_break_and_para_mark',
  'underline_tab_in_num_list',
  'use_alt_kinsoku_line_break_rules',
  'use_ansi_kerning_pairs',
  'use_normal_style_for_list',
  'use_word2002_table_style_rules'),
 ('allow_space_of_same_style_in_table',
  'autofit_to_first_fixed_width_cell',
  'cached_col_balance',
  'display_hangul_fixed_width',
  'do_not_autofit_constrained_tables',
  'do_not_break_constrained_forced_table',
  'do_not_suppress_indentation',
  'do_not_use_indent_as_numbering_tab_stop',
  'do_not_vert_align_cell_with_sp',
  'do_not_vert_align_in_txbx',
  'grow_autofit',
  'split_pg_break_and_para_mark',
  'use_alt_kinsoku_line_break_rules',
  'use_ansi_kerning_pairs',
  'use_normal_style_for_list',
  'use_word2002_table_style_rules'),
 ('allow_space_of_same_style_in_table',
  'autofit_to_first_fixed_width_cell',
  'cached_col_balance',
  'display_hangul_fixed_width',
  'do_not_autofit_constrained_tables',
  'do_not_break_constrained_forced_table',
  'do_not_suppress_indentation',
  'do_not_use_indent_as_numbering_tab_stop',
  'do_not_vert_align_cell_with_sp',
  'do_not_vert_align_in_txbx',
  'split_pg_break_and_para_mark',
  'use_alt_kinsoku_line_break_rules',
  'use_ansi_kerning_pairs',
  'use_normal_style_for_list'),
 (),
 (),
 (),
 (),
 ())
