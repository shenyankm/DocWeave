"""Owned note/comment parts and real VML/DrawingML textbox row editing."""
from io import BytesIO
from zipfile import ZipFile
from xml.etree.ElementTree import fromstring, tostring
from xml.dom import minidom

import pytest

import dockiln as dk
from dockiln.dom import Comment, DocxDocument, Footnote, StoryPart
from .test_dom_story_row_height import W, R, story_package, row_format


KINDS = ['footnotes', 'endnotes', 'comments', 'vml', 'dml']


@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
@pytest.mark.parametrize('inherited', [False, True])
@pytest.mark.parametrize('flat', [False, True])
def test_multiple_entries_keep_references_and_independent_heights(tmp_path, kind, inherited, flat):
    source = tmp_path / 'input.docx'
    name, parts = additional_story(source, kind, inherited=inherited)
    tree = minidom.parseString(parts[name])
    second = tree.documentElement.firstChild.cloneNode(True)
    second.setAttributeNS(W, 'w:id', '2')
    tree.documentElement.appendChild(second)
    parts[name] = tree.toxml(encoding='utf-8')
    body = minidom.parseString(parts['word/document.xml'])
    marker = body.getElementsByTagNameNS(W, kind.removesuffix('s') + 'Reference')[0]
    copied = marker.cloneNode(True)
    copied.setAttributeNS(W, 'w:id', '2')
    marker.parentNode.appendChild(copied)
    parts['word/document.xml'] = body.toxml(encoding='utf-8')
    with ZipFile(source, 'w') as archive:
        for part, data in parts.items():
            archive.writestr(part, data)
    document = DocxDocument(source)
    entries = document.story(name).child_nodes
    formats = [entry.get_child_nodes(dk.NodeType.ROW, True)[0].row_format for entry in entries]
    formats[0].height = 2000
    if not inherited:
        formats[1].height = 1440
    output = tmp_path / 'output'
    (document.save_flat_opc if flat else document.save)(output)
    expected = [1584, 2000 if inherited else 1440]
    assert [format.height for format in formats] == expected
    cold = DocxDocument(output)
    cold_entries = cold.story(name).child_nodes
    assert [entry._element.getAttributeNS(W, 'id') for entry in cold_entries] == ['1', '2']
    assert [entry.get_child_nodes(dk.NodeType.ROW, True)[0].row_format.height
            for entry in cold_entries] == expected
    assert tostring(fromstring(cold._package.payload('word/document.xml'))) == tostring(
        fromstring(parts['word/document.xml']))


def additional_story(path, kind, *, inherited=False, invalid=False, opaque=False):
    _, parts = story_package(path, 'header', inherited=inherited)
    header = parts.pop('word/header1.xml')
    table = header.decode().split('>', 1)[1].rsplit('</w:hdr>', 1)[0]
    if invalid:
        table = table.replace('<w:trPr>', '<w:trPr><w:trHeight w:val="1PT"/>', 1)
    parts['word/_rels/document.xml.rels'] = parts['word/_rels/document.xml.rels'].replace(
        f'<Relationship Id="owned-story" Type="{R}/header" Target="header1.xml"/>'.encode(), b'')
    parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(
        b'<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>', b'')
    prefix = f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>'
    if kind in ['footnotes', 'endnotes', 'comments']:
        entry = {'footnotes': 'footnote', 'endnotes': 'endnote', 'comments': 'comment'}[kind]
        attributes = 'w:id="1"' if kind != 'comments' else 'w:id="1" w:author="Owned" w:initials="O" w:date="2026-10-11T00:00:00Z"'
        name = 'owned/' + kind + '.part'
        parts[name] = f'<w:{kind} xmlns:w="{W}"><w:{entry} {attributes}>{table}</w:{entry}></w:{kind}>'.encode()
        parts['word/_rels/document.xml.rels'] = parts['word/_rels/document.xml.rels'].replace(
            b'</Relationships>', f'<Relationship Id="owned-story" Type="{R}/{kind}" Target="../{name}"/></Relationships>'.encode())
        parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(b'</Types>',
            f'<Override PartName="/{name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/></Types>'.encode())
        marker = f'<w:{entry}Reference w:id="1"/>' if kind != 'comments' else '<w:commentReference w:id="1"/>'
        body = f'<w:p><w:r><w:t>Owned body</w:t>{marker}</w:r></w:p>'
    else:
        name = 'word/document.xml'
        if kind == 'vml':
            box = f'<w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="OwnedBox" style="width:200pt;height:300pt" coordsize="21600,21600"><v:textbox><w:txbxContent>{table}</w:txbxContent></v:textbox></v:shape></w:pict>'
        else:
            box = f'''<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wp:inline><wp:extent cx="2540000" cy="3810000"/><wp:docPr id="1" name="OwnedBox"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2540000" cy="3810000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>{table}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>'''
        body = '<w:p><w:r>' + box + '</w:r></w:p>'
    if opaque:
        body += '<ext:opaque xmlns:ext="urn:owned:opaque">' + table + '</ext:opaque>'
    parts['word/document.xml'] = (prefix + body + '<w:sectPr/></w:body></w:document>').encode()
    with ZipFile(path, 'w') as archive:
        for part, data in parts.items():
            archive.writestr(part, data)
    return name, parts


@pytest.mark.parametrize('kind', KINDS)
@pytest.mark.parametrize('flat', [False, True])
@pytest.mark.parametrize('accessed', [False, True])
def test_projection_and_held_updates_in_typed_stories(tmp_path, kind, flat, accessed):
    source = tmp_path / 'input.docx'
    name, _ = additional_story(source, kind)
    document = DocxDocument(source)
    held = row_format(document, name) if accessed else None
    pure = document.to_flat_opc() if flat else document.to_bytes()
    assert row_format(DocxDocument(BytesIO(pure)), name).height == 1584
    assert held is None or held.height == 2000
    output = tmp_path / 'result'
    for _ in range(2):
        (document.save_flat_opc if flat else document.save)(output)
        assert held is None or held.height == 1584
        assert row_format(DocxDocument(output), name).height == 1584
        assert row_format(DocxDocument(output), name).height_rule is dk.HeightRule.EXACTLY


@pytest.mark.parametrize('kind', KINDS)
@pytest.mark.parametrize('flat', [False, True])
def test_inherited_values_and_public_edits(tmp_path, kind, flat):
    source = tmp_path / 'input.docx'
    name, parts = additional_story(source, kind, inherited=True)
    document = DocxDocument(source)
    held = row_format(document, name)
    assert held.height == 2000
    pure = document.to_flat_opc() if flat else document.to_bytes()
    assert row_format(DocxDocument(BytesIO(pure)), name).height == 2000
    held.height = 2000
    held.allow_break_across_pages = False
    assert held.height_rule is dk.HeightRule.AT_LEAST
    output = tmp_path / 'result'
    for _ in range(2):
        (document.save_flat_opc if flat else document.save)(output)
        cold = row_format(DocxDocument(output), name)
        assert held.height == cold.height == 1584
        assert cold.height_rule is dk.HeightRule.AT_LEAST
        assert held.allow_break_across_pages is cold.allow_break_across_pages is False
    if kind in ['footnotes', 'endnotes', 'comments']:
        root = document.story(name)
        assert isinstance(root, StoryPart)
        entry = root.child_nodes[0]
        assert isinstance(entry, Comment if kind == 'comments' else Footnote)
        assert entry.xml.startswith('<w:' + kind.removesuffix('s'))
        assert entry._element.getAttributeNS(W, 'id') == '1'
        assert document._package.payload('word/document.xml') == parts['word/document.xml']


@pytest.mark.parametrize('kind', KINDS)
@pytest.mark.parametrize('flat', [False, True])
def test_publish_failure_keeps_all_live_state(tmp_path, monkeypatch, kind, flat):
    from dockiln import _io
    source = tmp_path / 'input.docx'
    name, _ = additional_story(source, kind)
    document = DocxDocument(source)
    held = row_format(document, name)
    before = document._package.tree(name).toxml()
    dirty = set(document._package._dirty)
    output = tmp_path / 'result'
    output.write_bytes(b'original target')
    def fail(*args):
        raise OSError('owned failure')
    monkeypatch.setattr(_io.os, 'replace', fail)
    with pytest.raises(OSError):
        (document.save_flat_opc if flat else document.save)(output)
    assert held.height == 2000
    assert document._package.tree(name).toxml() == before
    assert document._package._dirty == dirty
    assert output.read_bytes() == b'original target'


@pytest.mark.parametrize('kind', KINDS)
def test_invalid_duplicate_in_registered_story_fails_at_load(tmp_path, kind):
    source = tmp_path / 'input.docx'
    additional_story(source, kind, invalid=True)
    with pytest.raises(RuntimeError, match='Invalid row-height'):
        DocxDocument(source)


@pytest.mark.parametrize('kind', ['vml', 'dml'])
def test_opaque_containers_are_neither_edited_nor_projected(tmp_path, kind):
    source = tmp_path / 'input.docx'
    additional_story(source, kind, opaque=True)
    document = DocxDocument(source)
    unknown = next(n for n in document.body.child_nodes if n._element.namespaceURI == "urn:owned:opaque")
    original = unknown.xml
    row = unknown.get_child_nodes(dk.NodeType.ROW, True)[0]
    with pytest.raises(NotImplementedError):
        row.row_format.height = 2
    exported = DocxDocument(BytesIO(document.to_bytes()))
    assert next(n for n in exported.body.child_nodes if n._element.namespaceURI == "urn:owned:opaque").xml == original
    assert unknown.xml == original
    assert fromstring('<root xmlns:w="' + W + '">' + next(n for n in exported.body.child_nodes if n._element.namespaceURI == "urn:owned:opaque").xml + '</root>')[0][0].find('.//{' + W + '}trHeight').get('{' + W + '}val') == '40000'


@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
def test_whole_note_part_identity_operations_remain_rejected(tmp_path, kind):
    source = tmp_path / 'input.docx'
    name, _ = additional_story(source, kind)
    document = DocxDocument(source)
    root = document.story(name)
    entry = root.child_nodes[0]
    before = document._package.tree(name).toxml()
    for node in [root]:
        with pytest.raises(NotImplementedError, match='reference updates'):
            node.remove()
        with pytest.raises(NotImplementedError, match='resource import'):
            node.clone(True)
    assert document._package.tree(name).toxml() == before
    assert not document._package._dirty


@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
def test_unlinked_note_xml_remains_opaque(tmp_path, kind):
    source = tmp_path / 'input.docx'
    name, parts = additional_story(source, kind)
    start = parts['word/_rels/document.xml.rels'].index(b'<Relationship Id="owned-story"')
    parts['word/_rels/document.xml.rels'] = parts['word/_rels/document.xml.rels'][:start] + b'</Relationships>'
    with ZipFile(source, 'w') as archive:
        for part, data in parts.items():
            archive.writestr(part, data)
    document = DocxDocument(source)
    assert name not in document._package._trees
    with pytest.raises(NotImplementedError):
        document.story(name)
    assert document._package.payload(name) == parts[name]


@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
@pytest.mark.parametrize('flat', [False, True])
def test_plain_story_table_copy_and_cross_document_import(tmp_path, kind, flat):
    source = tmp_path / 'source.docx'
    name, _ = additional_story(source, kind)
    document = DocxDocument(source)
    table = document.story(name).get_child_nodes(dk.NodeType.TABLE, True)[0]
    original = document._package.tree(name).toxml()
    copied = table.clone(True)
    assert copied.parent_node is None
    assert copied.owner_document is document
    assert copied.get_child_nodes(dk.NodeType.ROW, True)[0].row_format.height == 2000
    # Existing destination-style import is sufficient for these identical owned styles.
    target = tmp_path / 'target.docx'
    story_package(target, 'header')
    destination = DocxDocument(target)
    imported = destination.import_node(table, True)
    destination.body.append_child(imported)
    held = imported.get_child_nodes(dk.NodeType.ROW, True)[0].row_format
    assert held.height == 2000
    held.allow_break_across_pages = False
    output = tmp_path / 'output'
    for _ in range(2):
        (destination.save_flat_opc if flat else destination.save)(output)
        cold = DocxDocument(output).body.get_child_nodes(dk.NodeType.ROW, True)[0].row_format
        assert held.height == cold.height == 1584
        assert held.height_rule is cold.height_rule is dk.HeightRule.EXACTLY
        assert held.allow_break_across_pages is cold.allow_break_across_pages is False
    assert document._package.tree(name).toxml() == original


@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
def test_cross_part_insertion_and_complex_import_do_not_change_target(tmp_path, kind):
    source = tmp_path / 'source.docx'
    name, _ = additional_story(source, kind)
    document = DocxDocument(source)
    table = document.story(name).get_child_nodes(dk.NodeType.TABLE, True)[0]
    target = tmp_path / 'target.docx'
    story_package(target, 'header')
    destination = DocxDocument(target)
    before = destination.to_bytes()
    with pytest.raises(ValueError, match='Cross-document/part'):
        destination.body.append_child(table)
    assert destination.to_bytes() == before
    # A relationship-bearing payload cannot be copied without resource translation.
    paragraph = table.get_child_nodes(dk.NodeType.PARAGRAPH, True)[0]
    paragraph._element.setAttributeNS(R, 'r:owned', 'missing-owned-resource')
    with pytest.raises(NotImplementedError):
        destination.import_node(table, True)
    assert destination.to_bytes() == before
    assert table.parent_node is not None


@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
@pytest.mark.parametrize('failure', ['missing', 'type', 'root', 'mode'])
def test_typed_relationship_failures_are_explicit_and_private(tmp_path, kind, failure):
    source = tmp_path / 'owned-secret-input.docx'
    name, parts = additional_story(source, kind)
    if failure == 'missing':
        parts.pop(name)
    elif failure == 'type':
        parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(
            f'wordprocessingml.{kind}+xml'.encode(), b'wordprocessingml.document.main+xml')
    elif failure == 'root':
        parts[name] = f'<w:document xmlns:w="{W}"/>'.encode()
    else:
        parts['word/_rels/document.xml.rels'] = parts['word/_rels/document.xml.rels'].replace(
            b'Target="../owned/', b'TargetMode="owned-secret-mode" Target="../owned/')
    with ZipFile(source, 'w') as archive:
        for part, data in parts.items():
            archive.writestr(part, data)
    original = source.read_bytes()
    with pytest.raises(ValueError) as error:
        DocxDocument(source)
    assert 'owned-secret' not in str(error.value)
    assert source.read_bytes() == original


def test_unknown_drawing_graphic_role_remains_opaque(tmp_path):
    source = tmp_path / 'source.docx'
    _, parts = additional_story(source, 'dml')
    parts['word/document.xml'] = parts['word/document.xml'].replace(
        b'uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
        b'uri="urn:owned:unknown-role"')
    with ZipFile(source, 'w') as archive:
        for part, data in parts.items():
            archive.writestr(part, data)
    document = DocxDocument(source)
    row = document.body.get_child_nodes(dk.NodeType.ROW, True)[0]
    with pytest.raises(NotImplementedError):
        row.row_format.height = 2
    with ZipFile(BytesIO(document.to_bytes())) as archive:
        assert archive.read('word/document.xml') == parts['word/document.xml']


@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
@pytest.mark.parametrize('operation', ['clone', 'import'])
def test_copy_retains_setter_only_zero_rule_with_new_xml_identity(tmp_path, kind, operation):
    source = tmp_path / 'source.docx'
    name, _ = additional_story(source, kind)
    document = DocxDocument(source)
    table = document.story(name).get_child_nodes(dk.NodeType.TABLE, True)[0]
    held = table.get_child_nodes(dk.NodeType.ROW, True)[0].row_format
    held.height = 0
    held.height_rule = dk.HeightRule.EXACTLY
    if operation == 'clone':
        copied = table.clone(True)
    else:
        target = tmp_path / 'target.docx'
        story_package(target, 'header')
        destination = DocxDocument(target)
        copied = destination.import_node(table, True)
        destination.body.append_child(copied)
    clone_format = copied.get_child_nodes(dk.NodeType.ROW, True)[0].row_format
    assert clone_format.height == 0
    assert clone_format.height_rule is dk.HeightRule.EXACTLY
    clone_format.height = 3
    assert clone_format.height == 3
    assert clone_format.height_rule is dk.HeightRule.EXACTLY
    assert held.height == 0
    assert held.height_rule is dk.HeightRule.EXACTLY


@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
@pytest.mark.parametrize('flat', [False, True])
def test_related_story_does_not_validate_or_project_opaque_descendants(tmp_path, kind, flat):
    source = tmp_path / 'input.docx'
    name, parts = additional_story(source, kind)
    entry = kind.removesuffix('s')
    opaque = ('<ext:opaque xmlns:ext="urn:owned:opaque"><w:tbl><w:tr><w:trPr>'
              '<w:trHeight w:val="1PT" w:hRule="exact"/></w:trPr></w:tr></w:tbl></ext:opaque>')
    parts[name] = parts[name].replace(f'</w:{entry}>'.encode(), (opaque + f'</w:{entry}>').encode())
    with ZipFile(source, 'w') as archive:
        for part, data in parts.items():
            archive.writestr(part, data)
    document = DocxDocument(source)
    unknown = next(node for node in document.story(name).child_nodes[0].child_nodes
                   if node._element.namespaceURI == 'urn:owned:opaque')
    before = unknown.xml
    with pytest.raises(NotImplementedError):
        unknown.get_child_nodes(dk.NodeType.ROW, True)[0].row_format.height = 3
    data = document.to_flat_opc() if flat else document.to_bytes()
    cold = DocxDocument(BytesIO(data))
    cold_unknown = next(node for node in cold.story(name).child_nodes[0].child_nodes
                        if node._element.namespaceURI == 'urn:owned:opaque')
    assert cold_unknown.xml == before
    assert row_format(cold, name).height == 1584
    assert unknown.xml == before
