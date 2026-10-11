"""Owned duplicate annotation declarations: logical first winner and atomic save."""
from io import BytesIO
from xml.dom import minidom
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss.dom import DocxDocument
from aspose.words_foss.dom.nodes import W
from .test_dom_note_lifecycle import create, document


def duplicate_document(tmp_path, kind, variant):
    doc = document(tmp_path)
    node = create(doc, kind)
    doc.body.paragraphs[0].append_child(node)
    with ZipFile(BytesIO(doc.to_bytes())) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    part = 'word/' + kind + 's.xml'
    tree = minidom.parseString(parts[part])
    entry = next(item for item in tree.getElementsByTagNameNS(W, kind)
                 if item.getAttributeNS(W, 'type') in {'', 'normal'})
    main = minidom.parseString(parts['word/document.xml'])
    ref = main.getElementsByTagNameNS(W, kind + 'Reference')[0]
    entry.setAttributeNS(W, 'w:id', '7')
    ref.setAttributeNS(W, 'w:id', '7')
    duplicate = entry.cloneNode(True)
    if variant != 'same':
        duplicate.getElementsByTagNameNS(W, 't')[0].firstChild.nodeValue = 'Owned conflicting duplicate'
    if variant == 'alias':
        duplicate.setAttributeNS(W, 'w:id', '+0007')
    entry.parentNode.appendChild(duplicate)
    count = 8 if kind == 'comment' else 6
    if variant == 'collision':
        for index in range(1, count):
            copied = entry.cloneNode(True)
            copied.setAttributeNS(W, 'w:id', str(20 + index))
            copied.getElementsByTagNameNS(W, 't')[0].firstChild.nodeValue = f'Owned entry {index}'
            entry.parentNode.appendChild(copied)
            anchor = ref.parentNode.cloneNode(True)
            anchor.getElementsByTagNameNS(W, kind + 'Reference')[0].setAttributeNS(W, 'w:id', str(20 + index))
            ref.parentNode.parentNode.appendChild(anchor)
    parts[part] = tree.toxml(encoding='utf-8')
    parts['word/document.xml'] = main.toxml(encoding='utf-8')
    path = tmp_path / 'duplicate.docx'
    with ZipFile(path, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return DocxDocument(path), part, count if variant == 'collision' else 1


@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
@pytest.mark.parametrize('variant', ['same', 'conflict', 'alias', 'collision'])
@pytest.mark.parametrize('edit', [False, True])
@pytest.mark.parametrize('flat', [False, True])
def test_duplicate_content_uses_first_declaration_and_saves_unique_ids(tmp_path, kind, variant, edit, flat):
    doc, part, count = duplicate_document(tmp_path, kind, variant)
    original = doc.part_xml(part)
    storage = doc.story(part).child_nodes
    node_type = aw.NodeType.COMMENT if kind == 'comment' else aw.NodeType.FOOTNOTE
    nodes = doc.get_child_nodes(node_type, True)
    assert len(nodes) == count
    expected = ['Owned entry'] + [f'Owned entry {index}' for index in range(1, count)]
    assert [node.first_paragraph.text for node in nodes] == expected
    assert all(node.parent_node is doc.body.paragraphs[0] for node in nodes)
    assert doc.part_xml(part) == original
    if edit:
        nodes[0].first_paragraph.runs[0].text = expected[0] = 'Owned edited first'
    before = doc.part_xml(part)
    assert part in doc.preservation_report()['modified']
    payload = doc.to_flat_opc() if flat else doc.to_bytes()
    assert doc.part_xml(part) == before
    assert len(doc.story(part).child_nodes) == count + 1
    cold = DocxDocument(BytesIO(payload))
    assert [node.first_paragraph.text for node in cold.get_child_nodes(node_type, True)] == expected
    entries = cold.story(part).child_nodes
    ids = [int(node._element.getAttributeNS(W, 'id')) for node in entries]
    assert len(ids) == len(set(ids)) == count
    output = tmp_path / 'saved'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(output)
        assert len(doc.story(part).child_nodes) == count
        assert storage[1].parent_node is None
        assert [node.first_paragraph.text for node in DocxDocument(output).get_child_nodes(node_type, True)] == expected


@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
@pytest.mark.parametrize('flat', [False, True])
def test_failed_duplicate_normalization_keeps_storage_and_existing_output(tmp_path, monkeypatch, kind, flat):
    from aspose.words_foss import _io
    doc, part, _ = duplicate_document(tmp_path, kind, 'conflict')
    node_type = aw.NodeType.COMMENT if kind == 'comment' else aw.NodeType.FOOTNOTE
    assert len(doc.get_child_nodes(node_type, True)) == 1
    before = doc.part_xml(part), doc.part_xml('word/document.xml')
    output = tmp_path / 'saved'
    output.write_bytes(b'Owned previous output')
    def fail(*args):
        raise OSError('Owned failed publish')
    monkeypatch.setattr(_io.os, 'replace', fail)
    with pytest.raises(OSError, match='Owned failed publish'):
        (doc.save_flat_opc if flat else doc.save)(output)
    assert (doc.part_xml(part), doc.part_xml('word/document.xml')) == before
    assert len(doc.story(part).child_nodes) == 2
    assert output.read_bytes() == b'Owned previous output'
