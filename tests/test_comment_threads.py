"""Self-generated comment thread metadata and edited save cycles."""
from copy import deepcopy
from datetime import datetime
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._comment_threads import (
    CONTENT_TYPE,
    RELATION,
    W14,
    W15,
    read_threads,
)
from aspose.words_foss._comment_threads import W as W_URI
from aspose.words_foss._flat_opc import encode
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.dom import DocxDocument
from tests.test_conversion_comments import owned_comment, package

W = '{' + W_URI + '}'
EX = '{' + W15 + '}'


def owned_thread(*, references=True, reverse=False, done='1', case='range'):
    with ZipFile(BytesIO(owned_comment(case))) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    comments = ET.fromstring(parts['word/comments.xml'])
    base = comments[0]
    for index in range(3):
        entry = base if index == 0 else deepcopy(base)
        entry.set(W + 'id', str(index))
        entry.set(W + 'author', f'OWNED {index}')
        entry[-1].set('{' + W14 + '}paraId', f'{index + 1:08X}')
        next(entry.iter(W + 't')).text = f'OWNED COMMENT {index}'
        if index:
            comments.append(entry)
    parts['word/comments.xml'] = ET.tostring(comments)
    document = ET.fromstring(parts['word/document.xml'])
    paragraph = next(p for p in document.iter(W + 'p') if any(n.tag == W + 'commentReference' for n in p.iter()))
    if references:
        for index in (1, 2):
            ET.SubElement(ET.SubElement(paragraph, W + 'r'), W + 'commentReference', {W + 'id': str(index)})
    parts['word/document.xml'] = ET.tostring(document)
    extension = ET.Element(EX + 'commentsEx')
    for index in (reversed(range(3)) if reverse else range(3)):
        attrs = {EX + 'paraId': f'{index + 1:08X}', EX + 'done': done if index == 0 else '0'}
        if index:
            attrs[EX + 'paraIdParent'] = '00000001'
        ET.SubElement(extension, EX + 'commentEx', attrs)
    parts['word/commentsExtended.xml'] = ET.tostring(extension)
    rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
    root = ET.fromstring(parts['word/_rels/document.xml.rels'])
    ET.SubElement(root, '{' + rel + '}Relationship', {'Id': 'threads', 'Type': RELATION, 'Target': 'commentsExtended.xml'})
    parts['word/_rels/document.xml.rels'] = ET.tostring(root)
    ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    root = ET.fromstring(parts['[Content_Types].xml'])
    ET.SubElement(root, '{' + ns + '}Override', {'PartName': '/word/commentsExtended.xml', 'ContentType': CONTENT_TYPE})
    parts['[Content_Types].xml'] = ET.tostring(root)
    return package(parts)


def owned_import_thread():
    with ZipFile(BytesIO(owned_thread())) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    styles = ET.fromstring(parts['word/styles.xml'])
    existing = {style.get(W + 'styleId') for style in styles.findall(W + 'style')}
    for name, kind in [('CommentText', 'paragraph'), ('CommentReference', 'character')]:
        if name not in existing:
            style = ET.SubElement(styles, W + 'style', {W + 'styleId': name, W + 'type': kind})
            ET.SubElement(style, W + 'name', {W + 'val': name})
    parts['word/styles.xml'] = ET.tostring(styles)
    return package(parts)


def model_snapshot(model):
    return [(story.identifier, story.comment_metadata.ancestor_identifier, story.comment_metadata.done,
             story.comment_metadata.author, [p.get_text() for p in story.children])
            for story in model.source_stories if story.kind == 'comment']


@pytest.mark.parametrize('flat_input', [False, True])
@pytest.mark.parametrize('via_json', [False, True])
@pytest.mark.parametrize('flat_output', [False, True])
@pytest.mark.parametrize('reverse', [False, True])
def test_thread_edit_json_and_consecutive_save(flat_input, via_json, flat_output, reverse):
    raw = owned_thread(reverse=reverse)
    model = aw.Document(BytesIO(encode(raw) if flat_input else raw)).light_document_model
    assert model_snapshot(model)[0][1:3] == (None, True)
    assert [item[1] for item in model_snapshot(model)[1:]] == ['0', '0']
    if via_json:
        model = ldm.Document.model_validate_json(model.model_dump_json())
    stories = [s for s in model.source_stories if s.kind == 'comment']
    stories[0].comment_metadata.done = False
    stories[1].comment_metadata.done = True
    stories[1].children[-1].runs[-1].text = 'OWNED EDITED REPLY'
    # A new terminal paragraph forces regeneration of the association IDs.
    stories[1].children.append(ldm.Paragraph(children=[ldm.Run(text='OWNED SECOND PARAGRAPH')]))
    assert stories[1].children[-1].text == 'OWNED SECOND PARAGRAPH'
    expected = model_snapshot(model)
    for _ in range(2):
        output = LdmDocxWriter().write_to_bytes(model)
        model = aw.Document(BytesIO(encode(output) if flat_output else output)).light_document_model
        assert model_snapshot(model) == expected


@pytest.mark.parametrize('flat', [False, True])
@pytest.mark.parametrize('reverse', [False, True])
@pytest.mark.parametrize('case', ['range', 'table'])
def test_dom_thread_edit_and_add_reply(flat, reverse, case):
    document = DocxDocument(BytesIO(owned_thread(reverse=reverse, case=case)))
    parent, first, second = document.get_child_nodes(aw.NodeType.COMMENT, True)
    held_replies = parent.replies
    assert held_replies.count == 2
    assert first.ancestor is parent and second.ancestor is parent
    assert parent.replies == (first, second)
    assert first.parent_node is parent.parent_node
    parent.done = False
    first.done = True
    first.author = 'OWNED EDITED AUTHOR'
    reply = parent.add_reply('OWNED NEW', 'ON', datetime(2026, 2, 3), 'OWNED ADDED REPLY')  # noqa: DTZ001
    assert reply.ancestor is parent and reply.done is False
    assert held_replies.count == 3
    with pytest.raises(ValueError, match='another reply'):
        reply.add_reply('NESTED', '', datetime(2026, 2, 3), 'NESTED')  # noqa: DTZ001
    for _ in range(2):
        output = document.to_flat_opc() if flat else document.to_bytes()
        document = DocxDocument(BytesIO(output))
        comments = document.get_child_nodes(aw.NodeType.COMMENT, True)
        assert len(comments) == 4
        expected_done = [False, False, True, False] if case == 'range' else [False, True, False, False]
        assert [c.done for c in comments] == expected_done
        assert [c.ancestor.id if c.ancestor else None for c in comments] == [None, 0, 0, 0]
        edited, added = (comments[2], comments[1]) if case == 'range' else (comments[1], comments[3])
        assert edited.author == 'OWNED EDITED AUTHOR'
        assert 'OWNED ADDED REPLY' in ''.join(p.text for p in added.paragraphs)


@pytest.mark.parametrize('mutation', ['self', 'nested', 'dangling', 'duplicate', 'identifier', 'unknown', 'child'])
def test_invalid_thread_graph_rejected(mutation):
    with ZipFile(BytesIO(owned_thread())) as archive:
        comments = ET.fromstring(archive.read('word/comments.xml'))
        extension = ET.fromstring(archive.read('word/commentsExtended.xml'))
    if mutation == 'self':extension[0].set(EX + 'paraIdParent', '00000001')
    elif mutation == 'nested':extension[2].set(EX + 'paraIdParent', '00000002')
    elif mutation == 'dangling':extension[2].set(EX + 'paraIdParent', '00000099')
    elif mutation == 'duplicate':extension.append(deepcopy(extension[0]))
    elif mutation == 'identifier':extension[0].set(EX + 'paraId', 'secret-invalid-value')
    elif mutation == 'child':ET.SubElement(extension[0], EX + 'unsupported').text = 'secret-invalid-value'
    else:extension[0].set(EX + 'unknown', 'secret-invalid-value')
    with pytest.raises((ValueError, NotImplementedError)) as error:
        read_threads(comments, extension)
    assert 'secret-invalid-value' not in str(error.value)


@pytest.mark.parametrize('value, expected', [(None, False), ('0', False), ('1', True), ('true', True),
                                           ('false', True), ('on', True), ('off', True), ('', True),
                                           ('00', True), (' 0', True), ('0 ', True), ('2', True),
                                           ('invalid', True), ('FALSE', True)])
def test_done_fixed_baseline_lexical_values(value, expected):
    with ZipFile(BytesIO(owned_thread())) as archive:
        comments = ET.fromstring(archive.read('word/comments.xml'))
        extension = ET.fromstring(archive.read('word/commentsExtended.xml'))
    if value is None:
        extension[0].attrib.pop(EX + 'done')
    else:
        extension[0].set(EX + 'done', value)
    assert read_threads(comments, extension)['0']['done'] is expected


def test_detached_reply_is_rejected_without_mutating_collection():
    document = DocxDocument(BytesIO(owned_comment()))
    parent = document.create_comment('OWNED NEW', 'ON', datetime(2026, 2, 3))  # noqa: DTZ001
    held_replies = parent.replies
    with pytest.raises(ValueError, match='no parent'):
        parent.add_reply('OWNED FIRST', 'OF', datetime(2026, 2, 3), 'FIRST')  # noqa: DTZ001
    assert held_replies.count == 0


def test_thread_move_clone_remove_reject_before_mutation():
    document = DocxDocument(BytesIO(owned_thread()))
    parent = document.get_child_nodes(aw.NodeType.COMMENT, True)[0]
    before = document.to_bytes()
    for action in [lambda: parent.remove(), lambda: parent.clone(True),
                   lambda: document.body.paragraphs[0].append_child(parent)]:
        with pytest.raises(NotImplementedError, match='comment threads'):
            action()
        assert document.to_bytes() == before


@pytest.mark.parametrize('failure', ['dangling', 'cycle', 'done'])
def test_invalid_edited_thread_save_is_atomic(tmp_path, failure):
    model = aw.Document(BytesIO(owned_thread())).light_document_model
    stories = [story for story in model.source_stories if story.kind == 'comment']
    if failure == 'dangling':stories[1].comment_metadata.ancestor_identifier = '99'
    elif failure == 'cycle':stories[0].comment_metadata.ancestor_identifier = '1'
    else:stories[0].comment_metadata.done = 'invalid'
    destination = tmp_path / 'existing.docx'
    destination.write_bytes(b'ORIGINAL OUTPUT')
    before = model.model_dump_json()
    with pytest.raises(ValueError):
        LdmDocxWriter().write(model, destination)
    assert destination.read_bytes() == b'ORIGINAL OUTPUT'
    assert model.model_dump_json() == before


@pytest.mark.parametrize('kind', ['dtd', 'missing', 'external', 'duplicate'])
def test_extension_relationship_and_xml_boundary(kind):
    from defusedxml import DefusedXmlException
    with ZipFile(BytesIO(owned_thread())) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    if kind == 'dtd':
        parts['word/commentsExtended.xml'] = b'<!DOCTYPE commentsEx []>' + parts['word/commentsExtended.xml']
    elif kind == 'missing':
        del parts['word/commentsExtended.xml']
    else:
        root = ET.fromstring(parts['word/_rels/document.xml.rels'])
        relation = next(node for node in root if node.get('Type') == RELATION)
        if kind == 'external':relation.set('TargetMode', 'External')
        else:root.append(deepcopy(relation))
        parts['word/_rels/document.xml.rels'] = ET.tostring(root)
    raw = package(parts)
    with pytest.raises((ValueError, DefusedXmlException, KeyError)):
        aw.Document(BytesIO(raw))
    with pytest.raises((ValueError, DefusedXmlException, KeyError)):
        DocxDocument(BytesIO(raw)).to_bytes()


def test_dom_save_commits_live_ids_and_preserves_completion(tmp_path):
    document = DocxDocument(BytesIO(owned_thread()))
    parent, first, _second = document.get_child_nodes(aw.NodeType.COMMENT, True)
    held = parent.replies
    added = parent.add_reply('NEW', 'N', datetime(2026, 2, 3), 'OWNED NEW REPLY')  # noqa: DTZ001
    assert [reply.id for reply in held] == [3, 1, 2]
    assert len(document.get_child_nodes(aw.NodeType.COMMENT_RANGE_START, True)) == 2
    parent.done = False
    first.done = True
    original_ids = [reply.id for reply in held]
    document.to_bytes()
    assert [reply.id for reply in held] == original_ids  # Pure projection.
    document.save(tmp_path / 'saved.docx')
    assert [reply.id for reply in held] == [1, 2, 3]
    assert added.ancestor is parent and first.done is True and parent.done is False
    loaded = DocxDocument(tmp_path / 'saved.docx')
    assert [node.done for node in loaded.get_child_nodes(aw.NodeType.COMMENT, True)] == [False, False, True, False]


def test_unanchored_source_records_remain_but_are_not_visible_replies():
    raw = owned_thread(references=False)
    document = DocxDocument(BytesIO(raw))
    comments = document.get_child_nodes(aw.NodeType.COMMENT, True)
    assert len(comments) == 1 and comments[0].replies.count == 0
    model = aw.Document(BytesIO(raw)).light_document_model
    output = LdmDocxWriter().write_to_bytes(model)
    restored = aw.Document(BytesIO(output)).light_document_model
    assert model_snapshot(restored) == model_snapshot(model)


@pytest.mark.parametrize('case', ['range', 'table'])
def test_added_reply_retains_range_positions(case):
    document = DocxDocument(BytesIO(owned_thread(case=case)))
    parent = document.get_child_nodes(aw.NodeType.COMMENT, True)[0]
    added = parent.add_reply('NEW', 'N', datetime(2026, 2, 3), 'OWNED NEW')  # noqa: DTZ001
    output = document.to_bytes()
    with ZipFile(BytesIO(output)) as archive:
        root = ET.fromstring(archive.read('word/document.xml'))
    paragraph = next(p for p in root.iter(W + 'p') if p.find(W + 'commentRangeStart') is not None)
    offset, markers = 0, []
    for child in paragraph:
        if child.tag in (W + 'commentRangeStart', W + 'commentRangeEnd'):
            markers.append((child.tag, child.get(W + 'id'), offset))
        elif child.tag == W + 'r':
            offset += sum(len(node.text or '') for node in child.iter(W + 't'))
    parent_start = next(pos for tag, identifier, pos in markers if tag == W + 'commentRangeStart' and identifier == '0')
    parent_end = next(pos for tag, identifier, pos in markers if tag == W + 'commentRangeEnd' and identifier == '0')
    # Pure projection assigns canonical IDs without mutating the held reply ID.
    assert added.id == 3
    assert next(pos for tag, identifier, pos in markers if tag == W + 'commentRangeStart' and identifier == '1') == parent_start
    assert next(pos for tag, identifier, pos in markers if tag == W + 'commentRangeEnd' and identifier == '1') == parent_end


@pytest.mark.parametrize(('text', 'paragraphs'), [
    ('', ['']), ('a\nb', ['a', 'b']), ('a\rb', ['a', 'b']),
    ('a\r\nb', ['a', 'b']), ('\n\n', ['', '', '']),
    ('a\tb', ['a\tb']), ('  x  ', ['  x  ']), ('a\vb', ['a\vb']),
])
def test_add_reply_text_controls(text, paragraphs):
    document = DocxDocument(BytesIO(owned_thread()))
    parent = document.get_child_nodes(aw.NodeType.COMMENT, True)[0]
    reply = parent.add_reply('R', 'R', datetime(2026, 2, 3), text)  # noqa: DTZ001
    assert [p.text for p in reply.paragraphs] == paragraphs
    expected_word_text = [("\x05" if index == 0 else "") + value + "\r"
                          for index, value in enumerate(paragraphs)]
    assert [p.get_text() for p in reply.paragraphs] == expected_word_text
    for _ in range(2):
        document = DocxDocument(BytesIO(document.to_flat_opc()))
        reply = document.get_child_nodes(aw.NodeType.COMMENT, True)[1]
        assert [p.text for p in reply.paragraphs] == paragraphs
        assert [p.get_text() for p in reply.paragraphs] == expected_word_text


@pytest.mark.parametrize('kind', ['parent', 'reply', 'paragraph'])
@pytest.mark.parametrize('flat', [False, True])
def test_import_preserves_completion_and_contained_reply_links(kind, flat):
    source = DocxDocument(BytesIO(owned_import_thread()))
    target = DocxDocument(BytesIO(owned_import_thread()))
    comments = source.get_child_nodes(aw.NodeType.COMMENT, True)
    node = comments[0] if kind == 'parent' else comments[1] if kind == 'reply' else comments[0].parent_node
    imported = target.import_node(node, True)
    if kind == 'paragraph':
        target.body.append_child(imported)
    else:
        assert imported.ancestor is None and imported.replies.count == 0
        assert imported.done is (kind == 'parent')
        target.body.paragraphs[0].append_child(imported)
    for _ in range(2):
        target = DocxDocument(BytesIO(target.to_flat_opc() if flat else target.to_bytes()))
        actual = target.get_child_nodes(aw.NodeType.COMMENT, True)
        if kind == 'paragraph':
            parent, first, second = actual[-3:]
            assert parent.done and parent.replies == (first, second)
            assert first.ancestor is parent and second.ancestor is parent
        else:
            assert actual[-1].done is (kind == 'parent')
            assert actual[-1].ancestor is None and actual[-1].replies.count == 0


@pytest.mark.parametrize('flat', [False, True])
def test_import_reply_only_paragraph_detaches_missing_ancestor(flat):
    # Fixed 26.9 detaches the reply when its paragraph has no ancestor anchor.
    with ZipFile(BytesIO(owned_import_thread())) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    document = ET.fromstring(parts['word/document.xml'])
    run = next(node for node in document.iter(W + 'r')
               if any(child.tag == W + 'commentReference' and child.get(W + 'id') == '1'
                      for child in node))
    parent = next(node for node in document.iter(W + 'p') if run in list(node))
    parent.remove(run)
    paragraph = ET.Element(W + 'p')
    paragraph.append(run)
    document.find(W + 'body').insert(1, paragraph)
    parts['word/document.xml'] = ET.tostring(document)
    source = DocxDocument(BytesIO(package(parts)))
    target = DocxDocument(BytesIO(owned_import_thread()))
    reply = next(comment for comment in source.get_child_nodes(aw.NodeType.COMMENT, True)
                 if any('OWNED COMMENT 1' in p.text for p in comment.paragraphs))
    imported = target.import_node(reply.parent_node, True)
    target.body.append_child(imported)
    for _ in range(2):
        comments = target.get_child_nodes(aw.NodeType.COMMENT, True)
        assert comments[-1].ancestor is None
        assert comments[-1].replies.count == 0
        assert any('OWNED COMMENT 1' in p.text for p in comments[-1].paragraphs)
        target = DocxDocument(BytesIO(target.to_flat_opc() if flat else target.to_bytes()))
