"""Owned fixtures for native paragraph-style font-name source-order rules."""
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest
import aspose.words_foss as aw

NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
FIELDS = ('name', 'name_ascii', 'name_other', 'name_bi', 'name_far_east')


@pytest.fixture
def font_document(tmp_path):
    path = tmp_path / 'input.docx'
    fonts = '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:eastAsia="Courier New"/>'
    rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
    parts = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
        '_rels/.rels': f'<Relationships xmlns="{rel}"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        'word/_rels/document.xml.rels': f'<Relationships xmlns="{rel}"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        'word/document.xml': f'<w:document xmlns:w="{NS[1:-1]}"><w:body><w:p><w:pPr><w:pStyle w:val="NameP"/></w:pPr><w:r><w:rPr/><w:t>IMPORT</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
        'word/styles.xml': f'<w:styles xmlns:w="{NS[1:-1]}"><w:docDefaults><w:rPrDefault><w:rPr>{fonts}</w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="NameP"><w:name w:val="NameP"/><w:basedOn w:val="Normal"/><w:rPr>{fonts}</w:rPr></w:style><w:style w:type="character" w:styleId="NameC"><w:name w:val="NameC"/><w:rPr>{fonts}</w:rPr></w:style></w:styles>',
    }
    with ZipFile(path, 'w') as archive:
        for key, value in parts.items():
            archive.writestr(key, value)
    return path


def names(font):
    return {field: getattr(font, field) for field in FIELDS}


def handles(document):
    run = next(node for node in document.get_child_nodes(aw.NodeType.RUN, True) if node.text == 'IMPORT')
    styles = {style.name: style for style in document.styles}
    return run, styles['NameP'], styles['NameC']


def style_conflict(path, mark_last):
    with ZipFile(path) as archive:
        parts = {key: archive.read(key) for key in archive.namelist()}
    tree = ET.fromstring(parts['word/styles.xml'])
    style = next(item for item in tree if item.get(NS + 'styleId') == 'NameP')
    body = style.find(NS + 'rPr')
    for attr in body.find(NS + 'rFonts').attrib:
        body.find(NS + 'rFonts').set(attr, 'Georgia Body')
    mark = ET.Element(NS + 'pPr')
    properties = ET.SubElement(mark, NS + 'rPr')
    ET.SubElement(properties, NS + 'b')
    ET.SubElement(properties, NS + 'rFonts', {NS + key: 'Courier Mark' for key in ('ascii', 'hAnsi', 'cs', 'eastAsia')})
    style.insert(len(style) if mark_last else list(style).index(body), mark)
    parts['word/styles.xml'] = ET.tostring(tree)
    with ZipFile(path, 'w') as archive:
        for key, value in parts.items():
            archive.writestr(key, value)



@pytest.mark.parametrize('mark_last', [False, True])
def test_style_order_and_single_channel_edit(font_document, tmp_path, mark_last):
    style_conflict(font_document, mark_last)
    winner = 'Courier Mark' if mark_last else 'Georgia Body'
    document = aw.DocxDocument(font_document)
    style = document.styles.get_by_id('NameP')
    run = document.body.paragraphs[0].runs[0]
    before = document.to_bytes()
    assert names(style.font) == names(run.effective_font) == dict.fromkeys(FIELDS, winner)
    assert document.to_bytes() == before
    style.font.name_ascii = 'Verdana'
    expected = dict.fromkeys(FIELDS, winner)
    expected.update(name='Verdana', name_ascii='Verdana')
    assert names(style.font) == names(run.effective_font) == expected
    for fmt in (aw.SaveFormat.DOCX, aw.SaveFormat.FLAT_OPC):
        path = tmp_path / str(fmt)
        path.write_bytes(document.to_bytes() if fmt == aw.SaveFormat.DOCX else document.to_flat_opc())
        cold = aw.DocxDocument(path)
        assert names(cold.styles.get_by_id('NameP').font) == expected
        assert names(cold.body.paragraphs[0].runs[0].effective_font) == expected
    tree = style._element
    assert tree.getElementsByTagNameNS(NS[1:-1], 'b')
    mark = next(node for node in tree.childNodes if getattr(node, 'localName', None) == 'pPr')
    assert not mark.getElementsByTagNameNS(NS[1:-1], 'rFonts')


@pytest.mark.parametrize('invalid', [None, '', 1])
def test_dom_style_invalid_edit_keeps_conflict_declarations(font_document, invalid):
    style_conflict(font_document, True)
    document = aw.DocxDocument(font_document)
    before = document.to_bytes()
    style = document.styles.get_by_id('NameP')
    with pytest.raises((RuntimeError, TypeError)):
        style.font.name_ascii = invalid
    assert document.to_bytes() == before
    assert style.font.name_ascii == 'Courier Mark'



def test_dom_style_merge_preserves_decorations_and_removes_losing_theme(font_document):
    style_conflict(font_document, False)
    document = aw.DocxDocument(font_document)
    style = document.styles.get_by_id('NameP')
    element = style._element
    groups = list(element.getElementsByTagNameNS(NS[1:-1], 'rFonts'))
    nested = groups[0]
    nested.removeAttributeNS(NS[1:-1], 'ascii')
    nested.setAttributeNS(NS[1:-1], 'w:asciiTheme', 'minorHAnsi')
    nested.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:owned', 'urn:owned')
    nested.setAttributeNS('urn:owned', 'owned:keep', 'metadata')
    assert style.font.name_ascii == 'Georgia Body'
    style.font.name_ascii = 'Verdana'
    assert style.font.name_ascii == 'Verdana'
    assert nested.getAttributeNS('urn:owned', 'keep') == 'metadata'
    assert not nested.hasAttributeNS(NS[1:-1], 'asciiTheme')
    assert names(document.body.paragraphs[0].runs[0].effective_font)['name'] == 'Verdana'



def test_dom_duplicate_mark_names_rejects_before_edit(font_document):
    style_conflict(font_document, True)
    document = aw.DocxDocument(font_document)
    style = document.styles.get_by_id('NameP')
    nested = style._element.getElementsByTagNameNS(NS[1:-1], 'pPr')[0].getElementsByTagNameNS(NS[1:-1], 'rPr')[0]
    nested.appendChild(nested.getElementsByTagNameNS(NS[1:-1], 'rFonts')[0].cloneNode(True))
    before = style._element.toxml()
    with pytest.raises(ValueError, match='Duplicate'):
        style.font.name = 'Verdana'
    assert style._element.toxml() == before
