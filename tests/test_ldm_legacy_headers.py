"""Legacy flat header/footer content must migrate without loss or duplication."""
from copy import deepcopy

import pytest
from pydantic import ValidationError

import dockiln as dk
from dockiln import light_document_model as ldm


def paragraph(text):
    return ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(bold=True))])


@pytest.mark.parametrize('by_alias', [False, True])
@pytest.mark.parametrize('typed', [False, True])
def test_legacy_content_without_section_is_migrated_and_input_is_reusable(by_alias, typed):
    header, footer = paragraph('OWNED HEADER Ω'), paragraph('OWNED FOOTER Ω')
    data = {'header_paragraphs': [header], 'footer_paragraphs': [footer]}
    if not typed:
        data = {key: [p.model_dump(by_alias=by_alias) for p in value]
                for key, value in data.items()}
    before = deepcopy(data)
    doc = ldm.Document.model_validate(data)
    # A first document adopts fresh typed nodes; runtime bindings are not declarations.
    declaration = lambda value: value.model_dump() if isinstance(value, ldm.Paragraph) else value
    assert {k: [declaration(p) for p in v] for k, v in data.items()} == {
        k: [declaration(p) for p in v] for k, v in before.items()}
    assert len(doc.sections) == 1
    assert [p.text for p in doc.header_paragraphs] == ['OWNED HEADER Ω']
    assert [p.text for p in doc.footer_paragraphs] == ['OWNED FOOTER Ω']
    for _ in range(2):
        doc = ldm.Document.model_validate_json(doc.model_dump_json(by_alias=by_alias))
        assert len(doc.sections[0].headers_footers) == 2
        assert doc.header_paragraphs[0]._children[0].font.bold is True


@pytest.mark.parametrize('kind', [0, 1])
@pytest.mark.parametrize('by_alias', [False, True])
def test_flat_mirror_of_later_sections_does_not_create_first_section_copy(kind, by_alias):
    doc = ldm.Document(sections=[ldm.Section(), ldm.Section(headers_footers=[
        ldm.HeaderFooter(header_footer_type=kind, children=[paragraph('SECOND')])]),
        ldm.Section(headers_footers=[ldm.HeaderFooter(header_footer_type=kind,
                                                    children=[paragraph('THIRD')])])])
    for _ in range(2):
        doc = ldm.Document.model_validate_json(doc.model_dump_json(by_alias=by_alias))
        assert doc.sections[0].headers_footers == []
        assert [len(s.headers_footers) for s in doc.sections] == [0, 1, 1]


@pytest.mark.parametrize('kind,key', [(0, 'header_paragraphs'), (1, 'footer_paragraphs')])
@pytest.mark.parametrize('first_has_content', [False, True])
def test_conflicting_flat_content_is_explicitly_rejected_without_mutation(kind, key, first_has_content):
    data = {'sections': [{'headers_footers': [
        {'header_footer_type': kind, 'children': [paragraph('SECTION').model_dump(by_alias=True)]}]}],
        key: [paragraph('OWNED PRIVATE LEGACY').model_dump(by_alias=True)]}
    if not first_has_content:
        data['sections'].insert(0, {})
    before = deepcopy(data)
    with pytest.raises(ValidationError, match='conflicts') as exc:
        ldm.Document.model_validate(data)
    assert 'OWNED PRIVATE LEGACY' not in str(exc.value)
    assert data == before


@pytest.mark.parametrize('raw', [False, 0, '', {}, [None], [{'children': [None]}]])
def test_malformed_legacy_content_is_rejected(raw):
    with pytest.raises(ValidationError):
        ldm.Document(header_paragraphs=raw)


def test_empty_legacy_lists_do_not_add_section():
    assert ldm.Document(header_paragraphs=[], footer_paragraphs=[]).sections == []


@pytest.mark.parametrize('fmt', ['DOCX', 'FLAT_OPC'])
@pytest.mark.parametrize('via_json', [False, True])
def test_legacy_content_survives_two_public_saves(tmp_path, fmt, via_json):
    model = ldm.Document(header_paragraphs=[paragraph('OWNED HEADER Ω')],
                         footer_paragraphs=[paragraph('OWNED FOOTER Ω')])
    if via_json:
        model = ldm.Document.model_validate_json(model.model_dump_json())
    doc = dk.Document()
    doc._document = model
    for step in (1, 2):
        path = tmp_path / f'{fmt}-{step}.bin'
        doc.save(path, getattr(dk.SaveFormat, fmt))
        doc = dk.Document(path)
        cold = doc.light_document_model
        assert [p.text for p in cold.header_paragraphs] == ['OWNED HEADER Ω']
        assert [p.text for p in cold.footer_paragraphs] == ['OWNED FOOTER Ω']
        assert cold.header_paragraphs[0]._children[0].font.bold is True
        assert cold.footer_paragraphs[0]._children[0].font.bold is True


@pytest.mark.parametrize('fmt', [dk.SaveFormat.DOCX, dk.SaveFormat.FLAT_OPC])
def test_loaded_bound_header_fonts_keep_legacy_mirrors_equal(tmp_path, fmt):
    model = ldm.Document(header_paragraphs=[paragraph('OWNED HEADER Ω')],
                         footer_paragraphs=[paragraph('OWNED FOOTER Ω')])
    document = dk.Document()
    document._document = model
    path = tmp_path / 'owned'
    document.save(path, fmt)
    current = dk.Document(path).light_document_model
    held = current.header_paragraphs[0]._children[0].font
    for _ in range(2):
        data = current.model_dump()
        current = ldm.Document.model_validate(data)
        current = ldm.Document.model_validate_json(current.model_dump_json())
        font = current.header_paragraphs[0]._children[0].font
        assert font.model_dump() == held.model_dump()
        assert font.size == held.size
        assert len(current.sections[0].headers_footers) == 2
