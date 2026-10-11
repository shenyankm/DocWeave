"""Strict public type contracts independently calibrated against 26.9."""
from enum import IntEnum
import pytest
import dockiln as dk

FLAGS = (
    'adjust_line_height_in_table',
    'align_tables_row_by_row',
    'allow_space_of_same_style_in_table',
    'apply_breaking_rules',
    'auto_space_like_word95',
    'autofit_to_first_fixed_width_cell',
    'balance_single_byte_double_byte_width',
    'cached_col_balance',
    'conv_mail_merge_esc',
    'disable_open_type_font_formatting_features',
    'display_hangul_fixed_width',
    'do_not_autofit_constrained_tables',
    'do_not_break_constrained_forced_table',
    'do_not_break_wrapped_tables',
    'do_not_expand_shift_return',
    'do_not_leave_backslash_alone',
    'do_not_snap_to_grid_in_cell',
    'do_not_suppress_indentation',
    'do_not_suppress_paragraph_borders',
    'do_not_use_east_asian_break_rules',
    'do_not_use_html_paragraph_auto_spacing',
    'do_not_use_indent_as_numbering_tab_stop',
    'do_not_vert_align_cell_with_sp',
    'do_not_vert_align_in_txbx',
    'do_not_wrap_text_with_punct',
    'footnote_layout_like_ww8',
    'forget_last_tab_alignment',
    'grow_autofit',
    'layout_raw_table_width',
    'layout_table_rows_apart',
    'line_wrap_like_word6',
    'mw_small_caps',
    'no_column_balance',
    'no_extra_line_spacing',
    'no_leading',
    'no_space_raise_lower',
    'no_tab_hang_ind',
    'override_table_style_font_size_and_justification',
    'print_body_text_before_header',
    'print_col_black',
    'select_fld_with_first_or_last_char',
    'shape_layout_like_ww8',
    'show_breaks_in_frames',
    'space_for_ul',
    'spacing_in_whole_points',
    'split_pg_break_and_para_mark',
    'sub_font_by_size',
    'suppress_bottom_spacing',
    'suppress_sp_bf_after_pg_brk',
    'suppress_spacing_at_top_of_page',
    'suppress_top_spacing',
    'suppress_top_spacing_wp',
    'swap_borders_facing_pgs',
    'swap_inside_and_outside_for_mirror_indents_and_relative_positioning',
    'transparent_metafiles',
    'truncate_font_heights_like_wp6',
    'ui_compat_97_to_2003',
    'ul_trail_space',
    'underline_tab_in_num_list',
    'use_alt_kinsoku_line_break_rules',
    'use_ansi_kerning_pairs',
    'use_fe_layout',
    'use_normal_style_for_list',
    'use_printer_metrics',
    'use_single_borderfor_contiguous_cells',
    'use_word2002_table_style_rules',
    'use_word2010_table_style_rules',
    'use_word97_line_break_rules',
    'wp_justification',
    'wp_space_width',
    'wrap_trail_spaces',
)


@pytest.mark.parametrize('name', FLAGS)
@pytest.mark.parametrize('value', [0, 1, None, 'false', 1.0, object()])
def test_all_flags_reject_non_bool_without_creating_or_mutating_state(name, value):
    doc = dk.Document()
    held = doc.compatibility_options
    before = doc.light_document_model.model_dump_json()
    with pytest.raises(TypeError):
        setattr(held, name, value)
    assert doc.light_document_model.model_dump_json() == before
    assert doc.compatibility_options is held


@pytest.mark.parametrize('name', FLAGS)
@pytest.mark.parametrize('value', [False, True])
def test_all_flags_accept_only_actual_bools(name, value):
    doc = dk.Document()
    held = doc.compatibility_options
    setattr(held, name, value)
    assert getattr(held, name) is value
    assert doc.compatibility_options is held


class OtherVersion(IntEnum):
    ZERO = 0


@pytest.mark.parametrize('value', [0, 1, -1, 8, None, 'WORD2019', 1.0, object(),
                                 False, True, OtherVersion.ZERO, [], {}, b'WORD2019'])
def test_optimization_rejects_non_word_version_enum_without_mutation(value):
    doc = dk.Document()
    held = doc.compatibility_options
    before = doc.light_document_model.model_dump_json()
    with pytest.raises(TypeError):
        held.optimize_for(value)
    assert doc.light_document_model.model_dump_json() == before
    assert doc.compatibility_options is held


@pytest.mark.parametrize('name', ['compatibility_mode', 'owned_undeclared'])
def test_undeclared_attributes_reject_read_and_write(name):
    options = dk.Document().compatibility_options
    with pytest.raises(AttributeError):
        getattr(options, name)
    with pytest.raises(AttributeError):
        setattr(options, name, True)


@pytest.mark.parametrize('args', [(), (dk.settings.MsWordVersion.WORD2019, 0)])
def test_optimization_arity_errors_preserve_model(args):
    doc = dk.Document()
    held = doc.compatibility_options
    before = doc.light_document_model.model_dump_json()
    with pytest.raises(TypeError):
        held.optimize_for(*args)
    assert doc.light_document_model.model_dump_json() == before
