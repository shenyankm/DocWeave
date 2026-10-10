"""Row getter inheritance and setter-local state are distinct contracts."""
import pytest

from aspose.words_foss.dom import DocxDocument
from aspose.words_foss.model.enums import HeightRule
from .test_dom_row_format import row_document, values


@pytest.mark.parametrize('default', [False, True])
@pytest.mark.parametrize('local_mode,local_height,local_rule', [('unset', 0, 2), ('source-zero', 0, 2), ('exact', 1, 1), ('source-auto', 2, 0)])
@pytest.mark.parametrize('operation', ['height', 'zero', 'warmzero', 'exact', 'atLeast', 'auto'])
def test_setters_use_direct_values_while_getters_inherit(tmp_path, default, local_mode, local_height, local_rule, operation):
    if local_mode == 'unset':
        direct = ''
        initial = (2000, 1, True)
    elif local_mode == 'source-zero':
        direct = '<w:trHeight w:val="0" w:hRule="exact"/>'
        initial = (0, 2, True)
    else:
        direct = f'<w:trHeight w:val="{local_height * 20}" w:hRule="{"exact" if local_rule == 1 else "auto"}"/>'
        initial = (local_height, local_rule, True)
    source, _ = row_document(tmp_path, direct,
        '<w:trPr><w:trHeight w:val="40000" w:hRule="exact"/></w:trPr>', default=default)
    document = DocxDocument(source)
    held = document.body.tables[0].rows[0].row_format
    assert values(held) == initial
    if operation == 'height':
        held.height = 2000
        expected = (2000, 0 if local_rule == 2 else local_rule, True)
    elif operation == 'zero':
        held.height = 0
        expected = (0, 2, True)
    elif operation == 'warmzero':
        held.height = 0
        held.height_rule = HeightRule.EXACTLY
        held.height = 2000
        expected = (2000, 1, True)
    else:
        rule = {'exact': HeightRule.EXACTLY, 'atLeast': HeightRule.AT_LEAST, 'auto': HeightRule.AUTO}[operation]
        held.height_rule = rule
        expected = (0 if rule == HeightRule.AUTO else local_height, int(rule), True)
    assert values(held) == expected
    for flat in [False, True]:
        output = tmp_path / ('result.xml' if flat else 'result.docx')
        for _ in range(2):
            (document.save_flat_opc if flat else document.save)(output)
            expected = (min(expected[0], 1584), expected[1], True)
            assert values(held) == expected
            cold = document.body.tables[0].rows[0].row_format
            assert held._node._element is cold._node._element
            assert values(DocxDocument(output).body.tables[0].rows[0].row_format) == (
                expected[0], 2 if expected[0] == 0 else expected[1], True)


@pytest.mark.parametrize('kind', ['header', 'footer'])
@pytest.mark.parametrize('flat', [False, True])
def test_inherited_story_height_edit_uses_local_rule(tmp_path, kind, flat):
    from .test_dom_story_row_height import story_package, row_format
    source = tmp_path / 'input.docx'
    name, _ = story_package(source, kind, custom=True, inherited=True)
    document = DocxDocument(source)
    held = row_format(document, name)
    assert values(held) == (2000, 1, True)
    held.height = 2000
    assert values(held) == (2000, 0, True)
    output = tmp_path / 'result'
    for _ in range(2):
        (document.save_flat_opc if flat else document.save)(output)
        assert values(held) == (1584, 0, True)
        assert values(row_format(DocxDocument(output), name)) == (1584, 0, True)
