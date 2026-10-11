"""Explicit model inputs must retain content or fail without changing callers."""
from copy import deepcopy

import pytest
from pydantic import ValidationError
from dockiln import light_document_model as ldm


@pytest.mark.parametrize('child', [{'_type': 'OwnedUnknown', 'text': 'OWNED SECRET'},
                                  {'text': 'OWNED SECRET'}, None, False, 0, 'OWNED SECRET'])
def test_invalid_inline_child_is_rejected_without_input_mutation(child):
    data = {'children': [{'_type': 'Run', 'text': 'KEEP'}, child], '_text': 'OWNED SECRET'}
    before = deepcopy(data)
    with pytest.raises(ValidationError):
        ldm.Paragraph.model_validate(data)
    assert data == before


@pytest.mark.parametrize('children', [None, False, 0, {}, '', 'PRIVATE'])
def test_invalid_children_container_is_rejected(children):
    data = {'children': children}
    before = deepcopy(data)
    with pytest.raises(ValidationError):
        ldm.Paragraph.model_validate(data)
    assert data == before


@pytest.mark.parametrize('runs', [[], [{'_type': 'Run', 'text': 'OWNED CONTENT'}]])
def test_runs_constructor_does_not_silently_discard_content(runs):
    data = {'runs': runs}
    with pytest.raises(ValidationError, match='children'):
        ldm.Paragraph.model_validate(data)
    assert data == {'runs': runs}


@pytest.mark.parametrize('model,data', [
    (ldm.Paragraph, {'children': [{'_type': 'Run', 'text': 'KEEP'}], '_text': 'computed'}),
    (ldm.ParagraphFormat, {'_frame_lock_anchor': True, '_frame_wrap_type': 2, 'frame_format': {'width': 30}}),
    (ldm.FrameFormat, {'_frame_lock_anchor': True, '_wrap_type': 2}),
    (ldm.Shape, {'horizontal_position': 12, 'vertical_position': 24}),
    (ldm.Row, {'_tblPrEx_width': {'value': 36}, 'row_format': {'preferred_width': {'value': 40}}}),
    (ldm.Row, {'row_format': {'preferred_width': {'value': 40}}}),
    (ldm.DocList, {'levels': [{'number_format': '%1.'}]}),
    (ldm.Document, {'header_paragraphs': [{'children': [{'_type': 'Run', 'text': 'HEADER'}]}],
                    'sections': [{'body': {'children': []}}]}),
])
def test_legacy_validators_do_not_modify_reusable_input(model, data):
    before = deepcopy(data)
    first = model.model_validate(data)
    second = model.model_validate(data)
    assert data == before
    assert first.model_dump_json() == second.model_dump_json()


@pytest.mark.parametrize('by_alias', [False, True])
def test_mixed_children_round_trip_preserves_order_text_and_deepcopy(by_alias):
    children = [ldm.BookmarkStart(name='owned'), ldm.Run(text='A'), ldm.FieldStart(field_type=88),
                ldm.Run(text=' CODE'), ldm.FieldSeparator(), ldm.Run(text='B'), ldm.FieldEnd(),
                ldm.BookmarkEnd(name='owned'), ldm.Shape(width=1, height=1),
                ldm.NoteReference(kind='footnote', identifier='1'),
                ldm.CommentMarker(kind='reference', identifier='0')]
    paragraph = ldm.Paragraph(children=children)
    assert paragraph._children == children
    data = paragraph.model_dump(by_alias=by_alias)
    original = deepcopy(data)
    restored = ldm.Paragraph.model_validate(data)
    assert data == original
    cold = ldm.Paragraph.model_validate_json(paragraph.model_dump_json(by_alias=by_alias))
    copied = deepcopy(paragraph)
    assert [type(c) for c in restored._children] == [type(c) for c in children]
    assert paragraph.text == restored.text == cold.text == copied.text == 'A CODEB'
    copied._children[1].text = 'COPIED'
    assert paragraph.text == 'A CODEB'


def test_body_unknown_node_is_kept_for_existing_writer_diagnostics():
    document = ldm.Document.model_validate({'sections': [{'body': {'children': [
        {'_type': 'OwnedForeignBlock', 'owned_payload': 'KEEP'}]}}]})
    node = document.sections[0].body.children[0]
    assert isinstance(node, ldm.UnknownNode)
    assert node.model_dump()['owned_payload'] == 'KEEP'


@pytest.mark.parametrize('fmt', ['DOCX', 'FLAT_OPC'])
@pytest.mark.parametrize('via_json', [False, True])
def test_valid_public_input_survives_repeated_save_and_reload(tmp_path, fmt, via_json):
    import dockiln as dk
    paragraph = {'_type': 'Paragraph', 'children': [
        {'_type': 'BookmarkStart', 'name': 'owned'},
        {'_type': 'Run', 'text': 'OWNED A '},
        {'_type': 'Run', 'text': 'Ω B', 'font': {'bold': True}},
        {'_type': 'BookmarkEnd', 'name': 'owned'}]}
    data = {'sections': [{'body': {'children': [paragraph]}}]}
    before = deepcopy(data)
    model = ldm.Document.model_validate(data)
    if via_json:
        model = ldm.Document.model_validate_json(model.model_dump_json())
    assert data == before
    doc = dk.Document()
    doc._document = model
    for save in (1, 2):
        path = tmp_path / f'{fmt}-{save}.bin'
        doc.save(path, getattr(dk.SaveFormat, fmt))
        cold = dk.Document(path)
        p = cold.light_document_model.sections[0].body.paragraphs[0]
        assert p.text == 'OWNED A Ω B'
        assert [type(child) for child in p._children] == [ldm.BookmarkStart, ldm.Run,
                                                        ldm.Run, ldm.BookmarkEnd]
        assert p._children[2].font.bold is True
