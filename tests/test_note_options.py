"""Public note options, association, validation and package-preserving edits."""
from copy import deepcopy
from io import BytesIO
from xml.etree import ElementTree as E
from zipfile import ZipFile

import pytest

import dockiln as dk
from dockiln import light_document_model as ldm
from dockiln._flat_opc import encode
from tests.test_conversion_notes import W, owned_notes, package


def document(dom=False):
    return dk.DocxDocument(BytesIO(owned_notes())) if dom else dk.Document()


@pytest.mark.parametrize('dom', [False, True])
@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
def test_defaults_and_constructor(kind, dom):
    doc = document(dom)
    if dom:
        # Source properties from the owned-note fixture remain live.
        assert doc.footnote_options.start_number == 7
    else:
        opt = getattr(doc, kind + '_options')
        assert opt.start_number == 1
        assert opt.number_style == (dk.NumberStyle.ARABIC if kind == 'footnote' else dk.NumberStyle.LOWERCASE_ROMAN)
        assert opt.position == (dk.notes.FootnotePosition.BOTTOM_OF_PAGE if kind == 'footnote' else dk.notes.EndnotePosition.END_OF_DOCUMENT)
        assert opt.restart_rule == dk.notes.FootnoteNumberingRule.CONTINUOUS
    cls = dk.notes.FootnoteOptions if kind == 'footnote' else dk.notes.EndnoteOptions
    with pytest.raises(TypeError):
        cls()
    with pytest.raises(TypeError):
        cls(doc)
    with pytest.raises(NotImplementedError, match='CLR'):
        getattr(doc, kind + '_options').get_type()


@pytest.mark.parametrize('field,value,error', [
    ('start_number', True, TypeError), ('start_number', 1.5, TypeError),
    ('start_number', None, TypeError), ('start_number', '1', TypeError),
    ('start_number', 2**31, OverflowError), ('start_number', -2**31-1, OverflowError),
    ('number_style', 0, TypeError), ('position', 1, TypeError),
    ('restart_rule', 0, TypeError), ('columns', False, TypeError),
    ('position', dk.notes.EndnotePosition.END_OF_DOCUMENT, TypeError),
])
@pytest.mark.parametrize('dom', [False, True])
def test_invalid_edit_is_atomic(field, value, error, dom):
    doc = document(dom)
    opt = doc.footnote_options
    before = getattr(opt, field)
    original = ({n: doc._package.payload(n) for n in doc.part_names} if dom else doc.light_document_model.model_dump())
    with pytest.raises(error):
        setattr(opt, field, value)
    assert getattr(opt, field) == before
    if dom:
        assert {n: doc._package.payload(n) for n in doc.part_names} == original
    else:
        assert doc.light_document_model.model_dump() == original


@pytest.mark.parametrize('fmt', [dk.SaveFormat.DOCX, dk.SaveFormat.FLAT_OPC])
@pytest.mark.parametrize('style', list(dk.NumberStyle))
def test_all_number_styles_survive_or_follow_observed_native_projection(style, fmt):
    doc = dk.Document()
    for kind in ['footnote', 'endnote']:
        getattr(doc, kind + '_options').number_style = style
    for _ in range(2):
        raw = doc.to_bytes(fmt)
        assert doc.footnote_options.number_style == style
        assert doc.endnote_options.number_style == style
        cold = dk.Document(BytesIO(raw))
        assert cold.footnote_options.number_style == (dk.NumberStyle.ARABIC if style == dk.NumberStyle.SIMP_CHIN_NUM4 else style)
        assert cold.endnote_options.number_style == (dk.NumberStyle.LOWERCASE_ROMAN if style == dk.NumberStyle.SIMP_CHIN_NUM4 else style)


@pytest.mark.parametrize('dom', [False, True])
@pytest.mark.parametrize('flat', [False, True])
def test_document_setter_replaces_explicit_section_override_and_two_saves(tmp_path, dom, flat):
    raw = owned_notes()
    with ZipFile(BytesIO(raw)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    root = E.fromstring(parts['word/document.xml'])
    body = root.find(W + 'body')
    E.SubElement(body[0], W + 'pPr').append(deepcopy(body[-1]))
    parts['word/document.xml'] = E.tostring(root)
    raw = package(parts)
    doc = dk.DocxDocument(BytesIO(encode(raw) if flat else raw)) if dom else dk.Document(BytesIO(raw))
    section = doc.sections[1].page_setup.footnote_options
    section.start_number = 8
    assert doc.sections[0].page_setup.footnote_options.start_number == 3
    assert section.start_number == 8
    doc.footnote_options.start_number = 4
    assert [s.page_setup.footnote_options.start_number for s in doc.sections] == [4, 4]
    section.start_number = 9
    assert doc.footnote_options.start_number == 4
    doc.footnote_options.position = dk.notes.FootnotePosition.BENEATH_TEXT
    doc.footnote_options.restart_rule = dk.notes.FootnoteNumberingRule.RESTART_PAGE
    if not dom:
        doc._document = ldm.Document.model_validate_json(doc.light_document_model.model_dump_json())
        assert doc.sections[1].page_setup.footnote_options.start_number == 9
        copied = doc.light_document_model.model_copy(deep=True)
        copied.footnote_options.start_number = 12
        assert doc.footnote_options.start_number == 4
    for _ in range(2):
        target = tmp_path / ('options.xml' if flat else 'options.docx')
        if dom:
            (doc.save_flat_opc if flat else doc.save)(target)
        else:
            doc.save(target, dk.SaveFormat.FLAT_OPC if flat else dk.SaveFormat.DOCX)
        cold = dk.Document(target)
        assert cold.footnote_options.start_number == 4
        assert [s.page_setup.footnote_options.start_number for s in cold.sections] == [4, 9]
        assert cold.footnote_options.restart_rule == dk.notes.FootnoteNumberingRule.RESTART_PAGE
        assert cold.footnote_options.position == dk.notes.FootnotePosition.BENEATH_TEXT


@pytest.mark.parametrize('dom', [False, True])
def test_columns_are_editable_warm_and_not_claimed_to_be_persisted(tmp_path, dom):
    doc = document(dom)
    doc.footnote_options.columns = -2147483648
    doc.first_section.page_setup.footnote_options.columns = 3
    assert doc.footnote_options.columns == -2147483648
    assert doc.first_section.page_setup.footnote_options.columns == 3
    target = tmp_path / 'options.docx'
    doc.save(target)
    assert doc.first_section.page_setup.footnote_options.columns == 3
    assert dk.Document(target).footnote_options.columns == 0
    assert dk.DocxDocument(target).first_section.page_setup.footnote_options.columns == 0


@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
@pytest.mark.parametrize('value,expected', [('bad', 0), ('1e0', 1), ('١', 0), ('2147483648', -2147483648)])
def test_raw_start_number_recovery(kind, value, expected):
    raw = owned_notes(kind)
    with ZipFile(BytesIO(raw)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    root = E.fromstring(parts['word/settings.xml'])
    root.find(W + kind + 'Pr').find(W + 'numStart').set(W + 'val', value)
    parts['word/settings.xml'] = E.tostring(root)
    for cls in [dk.Document, dk.DocxDocument]:
        doc = cls(BytesIO(package(parts)))
        assert getattr(doc, kind + '_options').start_number == expected


def test_reusing_bound_page_setup_does_not_retarget_original_document():
    original = ldm.Document(sections=[ldm.Section()])
    original.footnote_options.start_number = 4
    # Remove the section override to exercise the associated document default.
    original.sections[0].page_setup.note_properties.clear()
    other = ldm.Document(sections=[original.sections[0]])
    other.footnote_options.start_number = 8
    assert original.sections[0].page_setup.footnote_options.start_number == 4
    assert other.sections[0].page_setup.footnote_options.start_number == 8


def test_section_reads_related_document_defaults_after_reader_population():
    raw = owned_notes()
    with ZipFile(BytesIO(raw)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    root = E.fromstring(parts['word/document.xml'])
    sect = root.find(W + 'body').find(W + 'sectPr')
    sect.remove(sect.find(W + 'footnotePr'))
    parts['word/document.xml'] = E.tostring(root)
    for cls in [dk.Document, dk.DocxDocument]:
        doc = cls(BytesIO(package(parts)))
        assert doc.first_section.page_setup.footnote_options.start_number == 7


@pytest.mark.parametrize('dom', [False, True])
def test_section_endnote_position_setter_updates_document_and_all_sections(dom):
    doc = document(dom)
    if not dom:
        doc.light_document_model.sections.append(ldm.Section())
        from dockiln.notes import bind_note_options
        bind_note_options(doc.light_document_model)
    doc.sections[-1].page_setup.endnote_options.position = dk.notes.EndnotePosition.END_OF_SECTION
    assert doc.endnote_options.position == dk.notes.EndnotePosition.END_OF_SECTION
    assert all(s.page_setup.endnote_options.position == dk.notes.EndnotePosition.END_OF_SECTION for s in doc.sections)


@pytest.mark.parametrize('version,expected', [
    (dk.settings.MsWordVersion.WORD2007, 0),
    (dk.settings.MsWordVersion.WORD2010, 0),
    (dk.settings.MsWordVersion.WORD2013, 3),
    (dk.settings.MsWordVersion.WORD2016, 3),
])
@pytest.mark.parametrize('dom', [False, True])
@pytest.mark.parametrize('flat', [False, True])
def test_columns_persist_in_modern_section_properties(tmp_path, version, expected, dom, flat):
    d = dk.Document(BytesIO(owned_notes()))
    d.compatibility_options.optimize_for(version)
    source = d.to_bytes(dk.SaveFormat.DOCX)
    doc = dk.DocxDocument(BytesIO(source)) if dom else dk.Document(BytesIO(source))
    doc.footnote_options.columns = 3
    assert doc.footnote_options.columns == 3
    for _ in range(2):
        target = tmp_path / ('columns.xml' if flat else 'columns.docx')
        if dom:
            (doc.save_flat_opc if flat else doc.save)(target)
        else:
            doc.save(target, dk.SaveFormat.FLAT_OPC if flat else dk.SaveFormat.DOCX)
        assert doc.footnote_options.columns == 3
        for cls in [dk.Document, dk.DocxDocument]:
            cold = cls(target)
            assert cold.footnote_options.columns == 0
            assert cold.first_section.page_setup.footnote_options.columns == expected


def test_dom_option_edits_preserve_unrelated_xml_parts_and_existing_nodes(tmp_path):
    with ZipFile(BytesIO(owned_notes())) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    settings = E.fromstring(parts['word/settings.xml'])
    E.SubElement(settings, '{urn:owned}keep', {'token': 'unchanged'})
    props = settings.find(W + 'footnotePr')
    E.SubElement(props, '{urn:owned}option', {'token': 'unchanged'})
    parts['word/settings.xml'] = E.tostring(settings)
    doc = dk.DocxDocument(BytesIO(package(parts)))
    paragraph = doc.body.child_nodes[0]
    element = paragraph._element
    doc.footnote_options.start_number = 9
    assert doc.body.child_nodes[0] is paragraph
    assert paragraph._element is element
    raw = doc._package.payload('word/settings.xml')
    current = E.fromstring(raw)
    assert current.find('{urn:owned}keep').get('token') == 'unchanged'
    assert current.find(W + 'footnotePr').find('{urn:owned}option').get('token') == 'unchanged'
    for name, payload in parts.items():
        if name not in {'word/settings.xml', 'word/document.xml'}:
            assert doc._package.payload(name) == payload


def test_endnote_imported_bad_position_recovers_document_and_section_defaults():
    with ZipFile(BytesIO(owned_notes('endnote'))) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    settings = E.fromstring(parts['word/settings.xml'])
    settings.find(W + 'endnotePr').find(W + 'pos').set(W + 'val', 'bad')
    parts['word/settings.xml'] = E.tostring(settings)
    root = E.fromstring(parts['word/document.xml'])
    props = root.find(W + 'body').find(W + 'sectPr').find(W + 'endnotePr')
    E.SubElement(props, W + 'pos', {W + 'val': 'bad'})
    parts['word/document.xml'] = E.tostring(root)
    for cls in [dk.Document, dk.DocxDocument]:
        doc = cls(BytesIO(package(parts)))
        assert doc.endnote_options.position == dk.notes.EndnotePosition.END_OF_DOCUMENT
        assert doc.first_section.page_setup.endnote_options.position == dk.notes.EndnotePosition.END_OF_DOCUMENT


@pytest.mark.parametrize('extension', ['txt', 'md'])
def test_text_reader_binds_section_options_and_json_preserves_value_equality(tmp_path, extension):
    source = tmp_path / ('input.' + extension)
    source.write_text('OWNED_NOTE_OPTIONS')
    doc = dk.Document(source)
    model = doc.light_document_model
    page = doc.first_section.page_setup
    assert page._note_defaults is model.note_properties
    page.endnote_options.position = dk.notes.EndnotePosition.END_OF_SECTION
    assert doc.endnote_options.position == dk.notes.EndnotePosition.END_OF_SECTION
    restored = ldm.Document.model_validate_json(model.model_dump_json())
    assert restored.sections[0].page_setup == page
    assert restored.model_copy(deep=True).sections[0].page_setup == page
    # Inherited document defaults do not change the section's declared values.
    other = ldm.Document(sections=[ldm.Section(page_setup=ldm.PageSetup.model_validate(page.model_dump()))])
    other.footnote_options.start_number = 4
    assert other.sections[0].page_setup != page
    same = ldm.PageSetup.model_validate(page.model_dump())
    assert same == page
    assert same != object()


@pytest.mark.parametrize('dom', [False, True])
@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
def test_option_setters_preserve_schema_order_before_separator_identifiers(dom, kind):
    doc = (dk.DocxDocument if dom else dk.Document)(BytesIO(owned_notes(kind)))
    options = getattr(doc, kind + '_options')
    options.restart_rule = dk.notes.FootnoteNumberingRule.RESTART_SECTION
    options.start_number = 8
    options.number_style = dk.NumberStyle.UPPERCASE_ROMAN
    options.position = (dk.notes.FootnotePosition.BENEATH_TEXT if kind == 'footnote'
                        else dk.notes.EndnotePosition.END_OF_SECTION)
    if dom:
        settings = E.fromstring(doc._package.payload('word/settings.xml'))
        nodes = list(settings.find(W + kind + 'Pr'))
        names = [node.tag.removeprefix(W) for node in nodes]
    else:
        props = next(p for p in doc.light_document_model.note_properties if p.kind == kind)
        names = [p.kind for p in props.children]
    assert names[:4] == ['pos', 'numFmt', 'numStart', 'numRestart']
    assert names[4:] == [kind, kind]


@pytest.mark.parametrize('value', ['bad', 'pageBottom', 'beneathText', ''])
@pytest.mark.parametrize('dom', [False, True])
def test_invalid_endnote_position_recovers_valid_enum_and_can_be_set_back(value, dom):
    with ZipFile(BytesIO(owned_notes('endnote'))) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    settings = E.fromstring(parts['word/settings.xml'])
    settings.find(W + 'endnotePr').find(W + 'pos').set(W + 'val', value)
    parts['word/settings.xml'] = E.tostring(settings)
    doc = (dk.DocxDocument if dom else dk.Document)(BytesIO(package(parts)))
    options = doc.endnote_options
    assert options.position is dk.notes.EndnotePosition.END_OF_DOCUMENT
    options.position = options.position
    assert options.position is dk.notes.EndnotePosition.END_OF_DOCUMENT


@pytest.mark.parametrize('json_roundtrip', [False, True])
def test_shallow_copy_option_defaults_keep_original_owner_context(json_roundtrip):
    def defaults(number):
        return [ldm.NoteProperties(kind='footnote', children=[
            ldm.NoteProperty(kind='numStart', value=str(number))])]
    original = ldm.Document(sections=[ldm.Section()], note_properties=defaults(5))
    if json_roundtrip:
        original = ldm.Document.model_validate_json(original.model_dump_json())
    copied = original.model_copy(update={'note_properties': defaults(9)})
    assert original.footnote_options.start_number == 5
    assert original.sections[0].page_setup.footnote_options.start_number == 5
    assert copied.footnote_options.start_number == 9
    assert copied.sections[0].page_setup.footnote_options.start_number == 9
    assert copied.sections[0].body is original.sections[0].body
    assert copied.sections[0].page_setup.note_properties is original.sections[0].page_setup.note_properties
    deep = copied.model_copy(deep=True)
    deep.footnote_options.start_number = 11
    assert original.sections[0].page_setup.footnote_options.start_number == 5
    assert copied.sections[0].page_setup.footnote_options.start_number == 9
    assert deep.sections[0].page_setup.footnote_options.start_number == 11


@pytest.mark.parametrize('dom', [False, True])
@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
def test_default_numbering_rule_alias_is_accessible_and_roundtrips(tmp_path, dom, kind):
    rule = dk.notes.FootnoteNumberingRule
    assert rule.DEFAULT is rule.CONTINUOUS
    assert int(rule.DEFAULT) == 0
    assert rule.__members__['DEFAULT'] is rule.CONTINUOUS
    doc = document(dom)
    options = getattr(doc, kind + '_options')
    options.restart_rule = rule.RESTART_SECTION
    options.restart_rule = rule.DEFAULT
    assert options.restart_rule is rule.DEFAULT
    if not dom:
        doc._document = ldm.Document.model_validate_json(doc.light_document_model.model_dump_json())
        assert getattr(doc, kind + '_options').restart_rule is rule.DEFAULT
    for format, extension in [(dk.SaveFormat.DOCX, 'docx'), (dk.SaveFormat.FLAT_OPC, 'xml')]:
        target = tmp_path / ('default.' + extension)
        if dom:
            (doc.save if extension == 'docx' else doc.save_flat_opc)(target)
        else:
            doc.save(target, format)
        cold = dk.Document(target)
        assert getattr(cold, kind + '_options').restart_rule is rule.DEFAULT


def test_page_setup_deepcopy_detaches_owner_and_keeps_declarations():
    doc = ldm.Document(sections=[ldm.Section()])
    doc.footnote_options.start_number = 5
    page = doc.sections[0].page_setup
    page.footnote_options.start_number = 3
    copied = deepcopy(page)
    assert copied._note_defaults is None
    assert copied._note_targets is None
    assert copied.model_dump() == page.model_dump()
    assert copied.footnote_options.start_number == 3
    assert page._note_defaults is doc.note_properties
    assert page._note_targets[0] is doc.note_properties
    copied.footnote_options.start_number = 9
    assert page.footnote_options.start_number == 3
    assert doc.footnote_options.start_number == 5


def test_first_detached_section_adoption_keeps_typed_caller_identity():
    original = ldm.Document(sections=[ldm.Section()])
    original.footnote_options.start_number = 5
    original.sections[0].page_setup.note_properties.clear()
    typed = deepcopy({'note_properties': original.note_properties, 'sections': original.sections})
    owned_section = typed['sections'][0]
    owned_page = owned_section.page_setup
    assert owned_page._note_defaults is None
    adopted = ldm.Document.model_validate(typed)
    assert adopted.sections[0] is owned_section
    assert adopted.sections[0].page_setup is owned_page
    assert owned_page._note_defaults is adopted.note_properties
    assert owned_page.footnote_options.start_number == 5
    adopted.footnote_options.start_number = 9
    assert original.footnote_options.start_number == 5
    assert original.sections[0].page_setup.footnote_options.start_number == 5


def test_direct_bound_section_adoption_still_isolates_original_owner():
    original = ldm.Document(sections=[ldm.Section()])
    original.footnote_options.start_number = 5
    original.sections[0].page_setup.note_properties.clear()
    adopted = ldm.Document.model_validate({'sections': original.sections,
        'note_properties': [ldm.NoteProperties(kind='footnote', children=[
            ldm.NoteProperty(kind='numStart', value='9')])]})
    assert adopted.sections[0] is not original.sections[0]
    assert adopted.sections[0].page_setup._note_defaults is adopted.note_properties
    assert adopted.sections[0].page_setup.footnote_options.start_number == 9
    assert original.sections[0].page_setup._note_defaults is original.note_properties
    assert original.sections[0].page_setup.footnote_options.start_number == 5


def test_document_deep_model_copy_rebinds_note_owner_views():
    original = ldm.Document(sections=[ldm.Section(), ldm.Section()])
    original.endnote_options.start_number = 5
    for section in original.sections:
        section.page_setup.note_properties.clear()
    copied = original.model_copy(deep=True)
    for section in copied.sections:
        assert section.page_setup._note_defaults is copied.note_properties
        assert section.page_setup._note_targets[0] is copied.note_properties
        assert section.page_setup.endnote_options.start_number == 5
    copied.sections[0].page_setup.endnote_options.position = dk.notes.EndnotePosition.END_OF_SECTION
    assert copied.endnote_options.position is dk.notes.EndnotePosition.END_OF_SECTION
    assert all(s.page_setup.endnote_options.position is dk.notes.EndnotePosition.END_OF_SECTION
               for s in copied.sections)
    assert original.endnote_options.position is dk.notes.EndnotePosition.END_OF_DOCUMENT
    assert all(s.page_setup.endnote_options.position is dk.notes.EndnotePosition.END_OF_DOCUMENT
               for s in original.sections)
