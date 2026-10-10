"""Public name getters follow live styles without writing inherited declarations."""

import copy
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss.dom.nodes import W
from .test_docx_dom import package

FIELDS = ("name", "name_ascii", "name_other", "name_bi", "name_far_east")

@pytest.fixture
def font_document(tmp_path):
    path = tmp_path / 'input.docx'
    fonts = '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:eastAsia="Courier New"/>'
    styles = (f'<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr>{fonts}'
              '</w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              f'<w:style w:type="paragraph" w:styleId="NameP"><w:name w:val="NameP"/>'
              f'<w:basedOn w:val="Normal"/><w:rPr>{fonts}</w:rPr></w:style>'
              f'<w:style w:type="character" w:styleId="NameC"><w:name w:val="NameC"/>'
              f'<w:rPr>{fonts}</w:rPr></w:style></w:styles>')
    package(path, body='<w:p><w:pPr><w:pStyle w:val="NameP"/></w:pPr>'
                      '<w:r><w:rPr/><w:t>IMPORT</w:t></w:r></w:p><w:sectPr/>',
            extras={'word/styles.xml': styles.encode(), '[Content_Types].xml': (
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="bin" ContentType="application/octet-stream"/>'
                '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                '</Types>').encode()})
    return path


NS = '{' + W + '}'


@pytest.fixture
def inherited_document(font_document):
    with ZipFile(font_document) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    styles = ET.fromstring(parts['word/styles.xml'])
    char = next(style for style in styles.findall(NS + 'style') if style.get(NS + 'styleId') == 'NameC')
    char.remove(char.find(NS + 'rPr'))
    body = ET.fromstring(parts['word/document.xml'])
    ET.SubElement(body.find('.//' + NS + 'rPr'), NS + 'rStyle', {NS + 'val': 'NameC'})
    parts['word/styles.xml'] = ET.tostring(styles)
    parts['word/document.xml'] = ET.tostring(body)
    with ZipFile(font_document, 'w') as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    return font_document


def handles(document):
    run = next(node for node in document.get_child_nodes(aw.NodeType.RUN, True) if node.text == 'IMPORT')
    styles = {style.name: style for style in document.styles}
    return run, styles['NameP'], styles['NameC']


def names(font):
    return {field: getattr(font, field) for field in FIELDS}


def test_missing_declarations_have_native_default_without_creating_local_names(inherited_document, tmp_path):
    with ZipFile(inherited_document) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    for name in ('word/styles.xml', 'word/document.xml'):
        root = ET.fromstring(parts[name])
        for parent in root.iter():
            for child in list(parent):
                if child.tag == NS + 'rFonts':
                    parent.remove(child)
        parts[name] = ET.tostring(root)
    with ZipFile(inherited_document, 'w') as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    document = aw.Document(inherited_document)
    fonts = [item.font for item in handles(document)]
    assert all(names(font) == dict.fromkeys(FIELDS, 'Times New Roman') for font in fonts)
    assert all(font.source_font_names is None and font.font_names_explicit is False for font in fonts)
    assert aw.light_document_model.Font().name == ''
    before = document.light_document_model.model_dump(exclude={"source_font_table"})
    path = tmp_path / 'defaults.docx'
    document.save(path)
    assert document.light_document_model.model_dump(exclude={"source_font_table"}) == before
    with ZipFile(path) as archive:
        for name in ('word/styles.xml', 'word/document.xml'):
            assert not list(ET.fromstring(archive.read(name)).iter(NS + 'rFonts'))


@pytest.mark.parametrize('field', FIELDS)
@pytest.mark.parametrize('local', [False, True])
def test_held_font_follows_parent_edit_with_independent_local_channel(inherited_document, tmp_path, field, local):
    document = aw.Document(inherited_document)
    run, paragraph, character = handles(document)
    font = run.font
    if local:
        font.name_ascii = 'Courier Local'
    setattr(paragraph.font, field, 'Verdana')
    expected = dict.fromkeys(FIELDS, 'Courier New')
    changed = FIELDS if field == 'name' else ('name', 'name_ascii') if field == 'name_ascii' else (field,)
    expected.update(dict.fromkeys(changed, 'Verdana'))
    if local:
        expected.update(name='Courier Local', name_ascii='Courier Local')
    assert names(font) == expected
    assert names(character.font) == dict.fromkeys(FIELDS, 'Courier New')
    assert character.font.source_font_names is None
    before = document.light_document_model.model_dump(exclude={"source_font_table"})
    for format in (aw.SaveFormat.DOCX, aw.SaveFormat.FLAT_OPC):
        path = tmp_path / f'output-{format}'
        document.save(path, format)
        assert document.light_document_model.model_dump(exclude={"source_font_table"}) == before
        assert names(handles(aw.Document(path))[0].font) == expected
    assert font.source_font_names is None if not local else font.source_font_names.attributes == {'ascii': 'Courier Local'}


@pytest.mark.parametrize('field', FIELDS)
def test_missing_character_font_is_editable_and_updates_held_run(inherited_document, field):
    run, _, character = handles(aw.Document(inherited_document))
    font = run.font
    setattr(character.font, field, 'Verdana')
    assert getattr(font, field) == 'Verdana'
    assert font.source_font_names is None


def test_font_copy_and_deep_document_copy_have_independent_owners(inherited_document):
    document = aw.Document(inherited_document)
    run, paragraph, _ = handles(document)
    font = run.font
    detached = font.model_copy(deep=True)
    python_copy = copy.deepcopy(font)
    copied_document = document.light_document_model.model_copy(deep=True)
    copied_run = next(node for node in copied_document.get_child_nodes(aw.NodeType.RUN, True) if node.text == 'IMPORT')
    copied_style = next(style for style in copied_document.styles if style.name == 'NameP')
    copied_style.font.name = 'Copied Parent'
    assert copied_run.font.name == 'Copied Parent'
    assert font.name == 'Courier New'
    paragraph.font.name = 'Original Parent'
    assert font.name == 'Original Parent'
    assert detached.name == python_copy.name == 'Courier New'
    run.font = detached.model_copy(update={'name_bi': 'Local Bidi'})
    assert run.font.name == 'Original Parent'
    assert run.font.name_bi == 'Local Bidi'
    assert run.font.source_font_names.attributes == {'cs': 'Local Bidi'}
