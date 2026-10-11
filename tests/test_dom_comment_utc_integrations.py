"""Owned UTC metadata interactions with duplicate entries and theme publication."""
from copy import deepcopy
from datetime import datetime, timezone
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as E

import pytest

from dockiln import NodeType
from dockiln.dom import DocxDocument
from dockiln._flat_opc import decode
from .test_dom_comment_utc import modern, W, W14, CID, CEX, REL, CT


def fixture(tmp_path, *, missing_theme=False):
    with ZipFile(BytesIO(modern(tmp_path).to_bytes())) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    comments = E.fromstring(parts['word/comments.xml'])
    first = comments[0]
    paragraph = first.find('{' + W + '}p')
    run = E.SubElement(paragraph, '{' + W + '}r')
    E.SubElement(run, '{' + W + '}t').text = 'Owned first'
    duplicate = deepcopy(first)
    duplicate.find('.//{' + W + '}t').text = 'Owned conflicting duplicate'
    duplicate.find('{' + W + '}p').set('{' + W14 + '}paraId', '34567890')
    comments.append(duplicate)
    if missing_theme:
        properties = E.Element('{' + W + '}rPr')
        E.SubElement(properties, '{' + W + '}rFonts', {'{' + W + '}asciiTheme': 'minorHAnsi'})
        run.insert(0, properties)
        for name in ['word/commentsIds.xml', 'word/commentsExtensible.xml']:
            parts.pop(name)
        relationships = E.fromstring(parts['word/_rels/document.xml.rels'])
        for item in list(relationships):
            if item.get('Type', '').endswith(('/commentsIds', '/commentsExtensible', '/theme')):
                relationships.remove(item)
        parts['word/_rels/document.xml.rels'] = E.tostring(relationships)
        types = E.fromstring(parts['[Content_Types].xml'])
        for item in list(types):
            if item.get('PartName') in ['/word/commentsIds.xml', '/word/commentsExtensible.xml']:
                types.remove(item)
        parts['[Content_Types].xml'] = E.tostring(types)
    else:
        ids = E.fromstring(parts['word/commentsIds.xml'])
        E.SubElement(ids, '{' + CID + '}commentId', {
            '{' + CID + '}paraId': '34567890', '{' + CID + '}durableId': '45678901'})
        extended = E.fromstring(parts['word/commentsExtensible.xml'])
        E.SubElement(extended, '{' + CEX + '}commentExtensible', {
            '{' + CEX + '}durableId': '45678901', '{' + CEX + '}dateUtc': '2044-01-02T03:04:05Z'})
        parts['word/commentsIds.xml'] = E.tostring(ids)
        parts['word/commentsExtensible.xml'] = E.tostring(extended)
    parts['word/comments.xml'] = E.tostring(comments)
    stream = BytesIO()
    with ZipFile(stream, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    doc = DocxDocument(BytesIO(stream.getvalue()))
    donor = source = None
    if missing_theme:
        donor = modern(tmp_path)
        source = donor.get_child_nodes(NodeType.COMMENT, True)[0]
        doc.body.paragraphs[0].append_child(doc.import_node(source, True))
    return doc, donor, source


def snapshot(doc, held):
    return (dict(doc._package._payloads), {name: tree.toxml() for name, tree in doc._package._trees.items()},
            set(doc._package._dirty), tuple(item.filename for item in doc._package._entries),
            tuple(doc._note_graph.duplicate_entries), doc._package._injected_default_theme,
            tuple((node, node._element, node.parent_node, node.id, node.author, node.initial,
                   node.date_time.isoformat(), str(node.date_time.tzinfo),
                   node.date_time_utc.isoformat(), str(node.date_time_utc.tzinfo),
                   node.first_paragraph.text) for node in doc.get_child_nodes(NodeType.COMMENT, True)))


@pytest.mark.parametrize('missing_theme', [False, True])
@pytest.mark.parametrize('flat', [False, True])
def test_duplicate_content_utc_and_theme_publish_once_after_atomic_failures(tmp_path, monkeypatch, missing_theme, flat):
    doc, donor, source = fixture(tmp_path, missing_theme=missing_theme)
    held = doc.get_child_nodes(NodeType.COMMENT, True)[0]
    assert held.first_paragraph.text == 'Owned first'
    if not missing_theme:
        assert held.date_time_utc == datetime(2030, 1, 2, 3, 4, 5, tzinfo=timezone.utc)
    held.date_time = datetime(2060, 1, 2, 3, 4, 5, tzinfo=timezone.utc)
    for name in doc.part_names:
        if name.endswith(('.xml', '.rels')):
            doc._package.tree(name)
    before = snapshot(doc, held)
    donor_before = snapshot(donor, source) if donor is not None else None
    output = tmp_path / 'published'
    output.write_bytes(b'previous-owned-output')
    save = doc.save_flat_opc if flat else doc.save
    with monkeypatch.context() as context:
        def fail(*args):
            raise OSError('owned publication failure')
        context.setattr('os.replace', fail)
        with pytest.raises(OSError, match='owned publication failure'):
            save(output)
    assert snapshot(doc, held) == before and output.read_bytes() == b'previous-owned-output'
    if donor is not None:
        assert snapshot(donor, source) == donor_before
    with monkeypatch.context() as context:
        context.setattr('dockiln._io.MAX_PART_BYTES', 1)
        with pytest.raises(ValueError, match='safety limit'):
            save(output)
    assert snapshot(doc, held) == before and output.read_bytes() == b'previous-owned-output'
    if donor is not None:
        assert snapshot(donor, source) == donor_before
    for _ in range(2):
        save(output)
        assert doc.get_child_nodes(NodeType.COMMENT, True)[0] is held
        if donor is not None:
            assert snapshot(donor, source) == donor_before
            assert donor.get_child_nodes(NodeType.COMMENT, True)[0] is source
        cold = DocxDocument(output)
        comments = cold.get_child_nodes(NodeType.COMMENT, True)
        assert len(comments) == (2 if missing_theme else 1) and comments[0].first_paragraph.text == 'Owned first'
        assert comments[0].date_time_utc == datetime(2060, 1, 2, 3, 4, 5, tzinfo=timezone.utc)
        if missing_theme:
            assert comments[1].date_time_utc == datetime(2030, 1, 2, 3, 4, 5, tzinfo=timezone.utc)
            assert source.date_time_utc == datetime(2030, 1, 2, 3, 4, 5, tzinfo=timezone.utc)
        data = decode(output.read_bytes()) if flat else output.read_bytes()
        with ZipFile(BytesIO(data)) as archive:
            assert len(archive.namelist()) == len(set(archive.namelist()))
            assert len(E.fromstring(archive.read('word/comments.xml'))) == (2 if missing_theme else 1)
            assert len(E.fromstring(archive.read('word/commentsIds.xml'))) == (2 if missing_theme else 1)
            assert len(E.fromstring(archive.read('word/commentsExtensible.xml'))) == (2 if missing_theme else 1)
            if missing_theme:
                assert 'word/theme/theme1.xml' in archive.namelist()
                relationships = E.fromstring(archive.read('word/_rels/document.xml.rels'))
                for suffix in ['commentsIds', 'commentsExtensible', 'theme']:
                    assert sum(item.get('Type', '').endswith('/' + suffix) for item in relationships) == 1


@pytest.mark.parametrize('flat', [False, True])
@pytest.mark.parametrize('extension', ['xml', 'bin'])
def test_new_header_override_theme_reference_is_projected_without_source_mutation(tmp_path, flat, extension):
    from dockiln._flat_opc import encode
    doc, _, _ = fixture(tmp_path)
    package = doc._package
    # The reference only exists in a newly added part, never in the source package.
    name = 'word/owned/header.' + extension
    header = ('<w:hdr xmlns:w="' + W + '"><w:p><w:r><w:rPr>'
              '<w:rFonts w:asciiTheme="minorHAnsi"/></w:rPr>'
              '<w:t>Owned header</w:t></w:r></w:p></w:hdr>').encode()
    rels = E.fromstring(package.payload('word/_rels/document.xml.rels'))
    E.SubElement(rels, '{' + REL + '}Relationship', Id='owned-header',
                 Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
                 Target='owned/header.xml')
    types = E.fromstring(package.payload('[Content_Types].xml'))
    E.SubElement(types, '{' + CT + '}Override', PartName='/' + name,
                 ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
    overrides = {name: header, 'word/_rels/document.xml.rels': E.tostring(rels),
                 '[Content_Types].xml': E.tostring(types)}
    doc.get_child_nodes(NodeType.COMMENT, True)
    before = snapshot(doc, None)
    raw = package.to_bytes(overrides)
    if flat:
        raw = decode(encode(raw))
    assert snapshot(doc, None) == before
    with ZipFile(BytesIO(raw)) as archive:
        assert len(archive.namelist()) == len(set(archive.namelist()))
        assert E.tostring(E.fromstring(archive.read(name))) == E.tostring(E.fromstring(header))
        if not flat:
            assert archive.read(name) == header
        assert 'word/theme/theme1.xml' in archive.namelist()
        relationships = E.fromstring(archive.read('word/_rels/document.xml.rels'))
        assert sum(item.get('Type', '').endswith('/theme') for item in relationships) == 1
        assert sum(item.get('Id') == 'owned-header' for item in relationships) == 1
        content_types = E.fromstring(archive.read('[Content_Types].xml'))
        for part in [name, 'word/theme/theme1.xml']:
            assert sum(item.get('PartName') == '/' + part for item in content_types) == 1


@pytest.mark.parametrize('flat', [False, True])
def test_binary_content_type_with_xml_name_remains_opaque(tmp_path, flat):
    from dockiln._flat_opc import encode
    doc, _, _ = fixture(tmp_path)
    package = doc._package
    name, payload = 'word/opaque.xml', b'Owned non-XML bytes\x00\xff'
    types = E.fromstring(package.payload('[Content_Types].xml'))
    E.SubElement(types, '{' + CT + '}Override', PartName='/' + name,
                 ContentType='application/octet-stream')
    package.set_parts({name: payload, '[Content_Types].xml': E.tostring(types)})
    before = dict(package._payloads)
    raw = package.to_bytes()
    if flat:
        raw = decode(encode(raw))
    with ZipFile(BytesIO(raw)) as archive:
        assert archive.read(name) == payload
    assert package._payloads == before


@pytest.mark.parametrize('extension', ['xml', 'bin'])
def test_declared_xml_with_invalid_bytes_is_not_hidden(tmp_path, extension):
    doc, _, _ = fixture(tmp_path)
    package = doc._package
    name = 'word/owned-invalid.' + extension
    types = E.fromstring(package.payload('[Content_Types].xml'))
    E.SubElement(types, '{' + CT + '}Override', PartName='/' + name, ContentType='application/xml')
    package.set_parts({name: b'Owned non-XML bytes', '[Content_Types].xml': E.tostring(types)})
    with pytest.raises(E.ParseError):
        package.to_bytes()
