"""Owned linked header/footer row projections, with opaque parts preserved."""
from io import BytesIO
from zipfile import ZipFile
from xml.etree.ElementTree import fromstring, tostring

import pytest

import aspose.words_foss as aw
from aspose.words_foss import _io
from aspose.words_foss.dom import DocxDocument
from .test_docx_dom import package, payloads

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def story_package(path, kind, *, custom=False, properties=None, inherited=False):
    name = f'owned/{kind}-rows.part' if custom else f'word/{kind}1.xml'
    tag = 'hdr' if kind == 'header' else 'ftr'
    properties = '<w:trHeight w:val="40000" w:hRule="exact"/>' if properties is None else properties
    table = ('<w:tbl><w:tblPr><w:tblStyle w:val="OwnedTable"/></w:tblPr>'
             '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr>'
             + ('' if inherited else properties) + '</w:trPr><w:tc><w:p><w:r>'
             '<w:t>Owned story</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>')
    extras = {name: f'<w:{tag} xmlns:w="{W}">{table}</w:{tag}>'.encode(),
              'customXml/opaque.xml': f'<w:hdr xmlns:w="{W}">{table}</w:hdr>'.encode()}
    parts = package(path, body=f'<w:p/><w:sectPr><w:{kind}Reference xmlns:r="{R}" '
                    f'w:type="default" r:id="owned-story"/></w:sectPr>', extras=extras)
    parts['word/document.xml'] = parts['word/document.xml'].replace(b'<?retain instruction?>', b'')
    target = '../' + name if custom else name.removeprefix('word/')
    parts['word/_rels/document.xml.rels'] = parts['word/_rels/document.xml.rels'].replace(
        b'</Relationships>', f'<Relationship Id="owned-story" Type="{R}/{kind}" '
        f'Target="{target}"/></Relationships>'.encode())
    parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(b'</Types>',
        (f'<Override PartName="/{name}" ContentType="application/vnd.openxmlformats-'
         f'officedocument.wordprocessingml.{kind}+xml"/>'
         '<Default Extension="bin" ContentType="application/octet-stream"/></Types>').encode())
    parts['word/styles.xml'] = (f'<w:styles xmlns:w="{W}"><w:style w:type="table" '
        'w:styleId="OwnedTable"><w:name w:val="OwnedTable"/><w:trPr>'
        '<w:trHeight w:val="40000" w:hRule="exact"/></w:trPr></w:style></w:styles>').encode()
    with ZipFile(path, 'w') as archive:
        for part, data in parts.items():
            archive.writestr(part, data)
    return name, parts


def row_format(document, name):
    return document.story(name).get_child_nodes(aw.NodeType.ROW, True)[0].row_format


@pytest.mark.parametrize('kind', ['header', 'footer'])
@pytest.mark.parametrize('custom', [False, True])
@pytest.mark.parametrize('flat', [False, True])
@pytest.mark.parametrize('accessed', [False, True])
def test_projection_includes_unvisited_linked_stories(tmp_path, kind, custom, flat, accessed):
    source = tmp_path / 'input.docx'
    name, original = story_package(source, kind, custom=custom)
    document = DocxDocument(source)
    held = row_format(document, name) if accessed else None
    if held:
        assert held.height == 2000
    exported = document.to_flat_opc() if flat else document.to_bytes()
    assert row_format(DocxDocument(BytesIO(exported)), name).height == 1584
    assert held is None or held.height == 2000
    assert name not in document._package._dirty
    out = tmp_path / ('result.xml' if flat else 'result.docx')
    (document.save_flat_opc if flat else document.save)(out)
    assert held is None or held.height == 1584
    assert row_format(document, name).height == 1584
    cold = DocxDocument(out)
    assert row_format(cold, name).height == 1584
    assert row_format(cold, name).height_rule is aw.HeightRule.EXACTLY
    for opaque in ['customXml/opaque.xml', 'word/styles.xml', 'word/media/image.bin']:
        actual = cold._package.payload(opaque)
        if flat and opaque.endswith('.xml'):
            assert tostring(fromstring(actual)) == tostring(fromstring(original[opaque]))
        else:
            assert actual == original[opaque]
    (document.save_flat_opc if flat else document.save)(out)
    assert row_format(DocxDocument(out), name).height == 1584


@pytest.mark.parametrize('kind', ['header', 'footer'])
@pytest.mark.parametrize('flat', [False, True])
def test_inherited_height_is_not_clamped_or_made_direct(tmp_path, kind, flat):
    source = tmp_path / 'input.docx'
    name, original = story_package(source, kind, custom=True, inherited=True)
    document = DocxDocument(source)
    held = row_format(document, name)
    assert held.height == 2000
    out = tmp_path / 'result'
    (document.save_flat_opc if flat else document.save)(out)
    assert held.height == row_format(DocxDocument(out), name).height == 2000
    actual = DocxDocument(out)._package.payload(name)
    if flat:
        assert tostring(fromstring(actual)) == tostring(fromstring(original[name]))
    else:
        assert actual == original[name]


@pytest.mark.parametrize('kind', ['header', 'footer'])
@pytest.mark.parametrize('flat', [False, True])
@pytest.mark.parametrize('failure', ['io', 'budget', 'publish'])
def test_failure_preserves_held_rows_package_and_destination(tmp_path, monkeypatch, kind, flat, failure):
    source = tmp_path / 'input.docx'
    name, _ = story_package(source, kind, custom=True, properties='')
    document = DocxDocument(source)
    held = row_format(document, name)
    held.height = 2000
    before = document._package.tree(name).toxml()
    dirty = set(document._package._dirty)
    out = tmp_path / 'result'
    out.write_bytes(b'original target')
    if failure == 'budget':
        monkeypatch.setattr(_io, 'MAX_PART_BYTES', 1)
        error = ValueError
    elif failure == 'publish':
        def fail_publish(*args, **kwargs):
            raise OSError('owned publish failure')
        monkeypatch.setattr(_io.os, 'replace', fail_publish)
        error = OSError
    else:
        from pathlib import Path
        def fail_write(*args, **kwargs):
            raise OSError('owned failure')
        monkeypatch.setattr(Path, 'write_bytes', fail_write)
        error = OSError
    with pytest.raises(error):
        (document.save_flat_opc if flat else document.save)(out)
    assert out.read_bytes() == b'original target'
    assert held.height == 2000
    assert document._package.tree(name).toxml() == before
    assert document._package._dirty == dirty


@pytest.mark.parametrize('kind', ['header', 'footer'])
def test_invalid_duplicate_in_unvisited_story_is_not_hidden(tmp_path, kind):
    source = tmp_path / 'input.docx'
    story_package(source, kind, custom=True, properties='<w:trHeight w:val="1PT"/>'
                  '<w:trHeight w:val="40000" w:hRule="exact"/>')
    original = source.read_bytes()
    with pytest.raises(RuntimeError, match='Invalid row-height'):
        DocxDocument(source)
    assert source.read_bytes() == original


def test_opaque_unlinked_word_root_is_not_projected(tmp_path):
    source = tmp_path / 'input.docx'
    _, original = story_package(source, 'footer', custom=True)
    document = DocxDocument(source)
    assert payloads(document.to_bytes())['customXml/opaque.xml'] == original['customXml/opaque.xml']


@pytest.mark.parametrize('kind', ['header', 'footer'])
@pytest.mark.parametrize('flat', [False, True])
def test_last_valid_duplicate_controls_story_projection(tmp_path, kind, flat):
    source = tmp_path / 'input.docx'
    name, _ = story_package(source, kind, custom=True, properties=
        '<w:trHeight w:val="20" w:hRule="atLeast"/>'
        '<w:trHeight w:val="40000" w:hRule="exact"/>')
    document = DocxDocument(source)
    held = row_format(document, name)
    assert held.height == 2000
    output = tmp_path / 'result'
    (document.save_flat_opc if flat else document.save)(output)
    assert held.height == row_format(DocxDocument(output), name).height == 1584
    assert 'w:val="20"' in document.part_xml(name)


def test_invalid_duplicate_added_after_load_fails_before_output(tmp_path):
    source = tmp_path / 'input.docx'
    name, _ = story_package(source, 'header', custom=True)
    document = DocxDocument(source)
    tree = document._package.tree(name)
    properties = tree.getElementsByTagNameNS(W, 'trPr')[0]
    invalid = tree.createElementNS(W, 'w:trHeight')
    invalid.setAttributeNS(W, 'w:val', '1PT')
    properties.insertBefore(invalid, properties.firstChild)
    before = tree.toxml()
    output = tmp_path / 'result'
    output.write_bytes(b'original')
    for save in [document.save, document.save_flat_opc, lambda _: document.to_bytes(),
                 lambda _: document.to_flat_opc()]:
        with pytest.raises(RuntimeError, match='Invalid row-height'):
            save(output)
        assert output.read_bytes() == b'original'
        assert tree.toxml() == before
