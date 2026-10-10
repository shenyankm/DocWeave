"""NodeType's fixed 26.9 enum surface, separately from node implementation coverage."""
import json
from io import BytesIO

import pytest

from aspose.words_foss import DocxDocument, NodeType
from aspose.words_foss import light_document_model as ldm
from .test_docx_dom import package

NAMES = '''ANY DOCUMENT SECTION BODY HEADER_FOOTER TABLE ROW CELL PARAGRAPH
BOOKMARK_START BOOKMARK_END EDITABLE_RANGE_START EDITABLE_RANGE_END
MOVE_FROM_RANGE_START MOVE_FROM_RANGE_END MOVE_TO_RANGE_START MOVE_TO_RANGE_END
GROUP_SHAPE SHAPE COMMENT FOOTNOTE RUN FIELD_START FIELD_SEPARATOR FIELD_END
FORM_FIELD SPECIAL_CHAR SMART_TAG STRUCTURED_DOCUMENT_TAG
STRUCTURED_DOCUMENT_TAG_RANGE_START STRUCTURED_DOCUMENT_TAG_RANGE_END
GLOSSARY_DOCUMENT BUILDING_BLOCK COMMENT_RANGE_START COMMENT_RANGE_END
OFFICE_MATH SUB_DOCUMENT SYSTEM NULL'''.split()


@pytest.mark.parametrize('value,name', list(enumerate(NAMES)))
def test_fixed_members_are_constructible_integer_enums(value, name):
    member = NodeType(value)
    assert member is getattr(NodeType, name)
    assert member.name == name and member.value == value
    assert isinstance(member, int) and int(member) == value
    assert json.loads(json.dumps(member)) == value


def test_enum_iteration_and_native_coercion():
    assert [member.name for member in NodeType] == NAMES
    assert NodeType(True) is NodeType.DOCUMENT
    assert NodeType(8.0) is NodeType.PARAGRAPH


@pytest.mark.parametrize('value', [-1, 39, 999, '8', None])
def test_invalid_constructor_raises_value_error(value):
    with pytest.raises(ValueError):
        NodeType(value)


def test_enum_and_integer_filters_agree_across_models_and_cold_package(tmp_path):
    run = ldm.Run(text='Owned')
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[ldm.Paragraph(children=[run])]))])
    assert model.get_child_nodes(NodeType(21), True) == model.get_child_nodes(21, True) == [run]
    restored = ldm.Document.model_validate_json(model.model_dump_json())
    assert restored.get_child_nodes(NodeType.RUN, True)[0].text == 'Owned'
    path = tmp_path / 'owned.docx'
    package(path, body='<w:p><w:r><w:t>Owned</w:t></w:r></w:p><w:sectPr/>')
    document = DocxDocument(path)
    assert document.get_child_nodes(NodeType.RUN, True) == document.get_child_nodes(21, True)
    cold = DocxDocument(BytesIO(document.to_bytes()))
    assert cold.get_child_nodes(NodeType(8), True)[0].node_type is NodeType.PARAGRAPH
