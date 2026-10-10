"""Native 26.9 UTF-16 name input boundaries, without external benchmark fixtures."""

from io import BytesIO
from xml.etree import ElementTree as ET

import pytest

import aspose.words_foss as aw
from aspose.words_foss.dom.nodes import W
from .test_docx_dom import package, payloads

FIELDS = ('name', 'name_ascii', 'name_other', 'name_bi', 'name_far_east')
VALUES = [
    ('\ud800', '\ufffd'), ('\udc00', '\ufffd'),
    ('\ud83d\ude00', '\U0001f600'), ('\udc00\ud800', '\ufffd\ufffd'),
    ('A\ud800B\udc00C', 'A\ufffdB\ufffdC'),
    ('A\ud83d\ude00B\ud800C', 'A\U0001f600B\ufffdC'),
    ('\ufffe', '\ufffe'), ('\uffff', '\uffff'), ('\U0010ffff', '\U0010ffff'),
]


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


def dom_font(document, target, *, effective=False):
    if target == 'run':
        run = document.body.paragraphs[0].runs[0]
        return run.effective_font if effective else run.font
    return document.styles.get_by_id('NameP' if target == 'paragraph' else 'NameC').font


def assert_saved_names(raw, target, field, normalized):
    document = aw.DocxDocument(BytesIO(raw))
    expected = 'Courier New' if normalized in {'\ufffe', '\uffff'} else normalized
    assert getattr(dom_font(document, target, effective=True), field) == expected
    for name, value in payloads(document.to_bytes()).items():
        if name.endswith(('.xml', '.rels')):
            ET.fromstring(value)


@pytest.mark.parametrize('value,normalized', VALUES)
@pytest.mark.parametrize('field', FIELDS)
@pytest.mark.parametrize('target', ['run', 'paragraph', 'character'])
def test_dom_name_utf16_input_and_xml_save_boundary(font_document, target, field, value, normalized):
    document = aw.DocxDocument(font_document)
    font = dom_font(document, target)
    setattr(font, field, value)
    assert getattr(font, field) == normalized
    live = font._node._element.toxml()
    for raw in (document.to_bytes(), document.to_flat_opc()):
        assert_saved_names(raw, target, field, normalized)
        assert font._node._element.toxml() == live
        assert getattr(font, field) == normalized




@pytest.mark.parametrize('field', FIELDS)
@pytest.mark.parametrize('value,error', [(None, RuntimeError), ('', RuntimeError), (42, TypeError)])
@pytest.mark.parametrize('target', ['run', 'paragraph', 'character'])
def test_invalid_name_keeps_package_unchanged(font_document, target, field, value, error):
    document = aw.DocxDocument(font_document)
    before = document.to_bytes()
    with pytest.raises(error):
        setattr(dom_font(document, target), field, value)
    assert document.to_bytes() == before


@pytest.mark.parametrize('field,channel', list(zip(FIELDS[1:], ('ascii', 'hAnsi', 'cs', 'eastAsia'))))
def test_channel_edit_preserves_other_declarations_and_hint(font_document, field, channel):
    document = aw.DocxDocument(font_document)
    run = document.body.paragraphs[0].runs[0]
    properties = run._element.getElementsByTagNameNS(W, 'rPr')[0]
    declaration = document._package.tree('word/document.xml').createElementNS(W, 'w:rFonts')
    for name in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
        declaration.setAttributeNS(W, 'w:' + name, 'Original-' + name)
    declaration.setAttributeNS(W, 'w:hint', 'eastAsia')
    properties.appendChild(declaration)
    setattr(run.font, field, 'Updated')
    assert declaration.getAttributeNS(W, 'hint') == 'eastAsia'
    for name in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
        assert declaration.getAttributeNS(W, name) == ('Updated' if name == channel else 'Original-' + name)
    cold = aw.DocxDocument(BytesIO(document.to_bytes()))
    assert getattr(cold.body.paragraphs[0].runs[0].font, field) == 'Updated'


def test_whole_name_replaces_four_theme_channels(font_document):
    document = aw.DocxDocument(font_document)
    run = document.body.paragraphs[0].runs[0]
    font = run.font
    properties = run._element.getElementsByTagNameNS(W, 'rPr')[0]
    declaration = document._package.tree('word/document.xml').createElementNS(W, 'w:rFonts')
    for name in ('asciiTheme', 'hAnsiTheme', 'cstheme', 'eastAsiaTheme'):
        declaration.setAttributeNS(W, 'w:' + name, 'minorHAnsi')
    properties.appendChild(declaration)
    font.name = 'Whole'
    assert all(not declaration.hasAttributeNS(W, name) for name in ('asciiTheme', 'hAnsiTheme', 'cstheme', 'eastAsiaTheme'))
    assert all(getattr(font, field) == 'Whole' for field in FIELDS)
    font.name_ascii = 'Latin'
    assert font.name == font.name_ascii == 'Latin'
    assert font.name_other == font.name_bi == font.name_far_east == 'Whole'
