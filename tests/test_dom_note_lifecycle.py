"""Owned annotation content/reference graph lifecycle and atomic errors."""
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime

import pytest

import aspose.words_foss as aw
from aspose.words_foss.dom import DocxDocument
from aspose.words_foss.dom.notes import FootnoteType
from aspose.words_foss.dom.importing import ImportFormatMode
from .test_docx_dom import package
from .test_dom_additional_story_rows import additional_story


def document(tmp_path):
    path = tmp_path / 'source.docx'
    parts = package(path, body='<w:p><w:r><w:t>Owned first</w:t></w:r></w:p><w:p><w:r><w:t>Owned second</w:t></w:r></w:p><w:sectPr/>')
    parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(b'</Types>', b'<Default Extension="bin" ContentType="application/octet-stream"/></Types>')
    parts['word/document.xml'] = parts['word/document.xml'].replace(b'<?retain instruction?>', b'')
    with ZipFile(path, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return DocxDocument(path)


def create(doc, kind, *, fill=True):
    node = doc.create_comment('Owned author', 'OA', datetime(2026, 10, 11)) if kind == 'comment' else doc.create_footnote(FootnoteType.FOOTNOTE if kind == 'footnote' else FootnoteType.ENDNOTE)
    if fill:
        node.ensure_minimum()
        node.first_paragraph.append_child(doc.create_run('Owned entry', part_name=node.part_name))
    return node


@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
@pytest.mark.parametrize('operation', ['attach', 'move', 'clone', 'shallow', 'remove'])
@pytest.mark.parametrize('flat', [False, True])
def test_note_lifecycle_keeps_unique_logical_parent_and_cold_graph(tmp_path, kind, operation, flat):
    doc = document(tmp_path)
    first, second = doc.body.paragraphs
    node = create(doc, kind)
    assert node.parent_node is None
    assert node.count == 1
    first.append_child(node)
    assert first.child_nodes[-1] is node
    assert node.parent_node is node.parent_paragraph is first
    assert doc.get_child_nodes(node.node_type, True) == [node]
    if operation == 'move':
        second.append_child(node)
        assert node.parent_node is second
        assert node not in first.child_nodes
    elif operation in ['clone', 'shallow']:
        copied = node.clone(operation != 'shallow')
        assert copied.parent_node is None
        assert copied.count == (1 if operation == 'clone' else 0)
        second.append_child(copied)
    elif operation == 'remove':
        assert node.remove() is node
        assert node.parent_node is None
        assert doc.get_child_nodes(node.node_type, True) == []
    pure = doc.to_flat_opc() if flat else doc.to_bytes()
    if operation == 'shallow':
        assert copied.count == 0
    output = tmp_path / 'result'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(output)
        cold = DocxDocument(output)
        nodes = cold.get_child_nodes(node.node_type, True)
        assert len(nodes) == (0 if operation == 'remove' else 2 if operation in ['clone', 'shallow'] else 1)
        for item in nodes:
            assert item.parent_node is not None
            assert item in item.parent_node.child_nodes
            assert item.count == 1
        if kind == 'comment':
            assert [item.id for item in nodes] == list(range(len(nodes)))
            if operation in ['clone', 'shallow']:
                assert (node.id, copied.id) == (0, 1)
        if operation == 'shallow':
            assert copied.count == 1
    assert len(DocxDocument(BytesIO(pure)).get_child_nodes(node.node_type, True)) == len(nodes)


@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
@pytest.mark.parametrize('flat', [False, True])
def test_failed_save_preserves_ids_minimum_and_output(tmp_path, monkeypatch, kind, flat):
    from aspose.words_foss import _io
    doc = document(tmp_path)
    node = create(doc, kind, fill=False)
    doc.body.paragraphs[0].append_child(node)
    copied = node.clone(False)
    doc.body.paragraphs[1].append_child(copied)
    original = [node.xml, copied.xml]
    output = tmp_path / 'result'
    output.write_bytes(b'Owned previous output')
    def fail(*args):
        raise OSError('Owned failed publish')
    monkeypatch.setattr(_io.os, 'replace', fail)
    with pytest.raises(OSError):
        (doc.save_flat_opc if flat else doc.save)(output)
    assert [node.xml, copied.xml] == original
    assert node.count == copied.count == 0
    if kind == 'comment':
        assert node.id == copied.id == 0
    assert output.read_bytes() == b'Owned previous output'


@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
def test_insertion_budget_failure_preserves_pending_and_all_parts(tmp_path, monkeypatch, kind):
    from aspose.words_foss import _io
    doc = document(tmp_path)
    node = create(doc, kind)
    before = doc.to_bytes()
    original = node.xml
    names = doc.part_names
    monkeypatch.setattr(_io, 'MAX_ZIP_ENTRIES', len(names))
    with pytest.raises(ValueError, match='too many ZIP'):
        doc.body.paragraphs[0].append_child(node)
    assert doc.to_bytes() == before
    assert doc.part_names == names
    assert node.xml == original
    assert node.parent_node is None


@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
def test_loaded_note_storage_view_and_canonical_are_distinct(tmp_path, kind):
    source = tmp_path / 'source.docx'
    name, _ = additional_story(source, kind)
    doc = DocxDocument(source)
    before = doc.to_bytes()
    storage = doc.story(name).child_nodes[0]
    logical = doc.get_child_nodes(aw.NodeType.COMMENT if kind == 'comments' else aw.NodeType.FOOTNOTE, True)[0]
    assert storage is not logical
    assert storage.canonical_node is logical
    assert storage.parent_node is doc.story(name)
    assert logical.parent_node is doc.body.paragraphs[0]
    assert logical in logical.parent_node.child_nodes
    assert doc.to_bytes() == before


@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
@pytest.mark.parametrize('mode', list(ImportFormatMode))
@pytest.mark.parametrize('flat', [False, True])
def test_cross_document_entry_import_retains_content_and_source(tmp_path, kind, mode, flat):
    source = document(tmp_path)
    node = create(source, kind)
    source.body.paragraphs[0].append_child(node)
    before = source.to_bytes()
    target = document(tmp_path)
    copied = target.import_node(node, True, mode)
    assert copied.owner_document is target
    assert copied.parent_node is None
    assert copied.count == 1
    target.body.paragraphs[1].append_child(copied)
    assert copied.parent_node is target.body.paragraphs[1]
    copied.first_paragraph.runs[0].text = 'Owned changed copy'
    assert node.first_paragraph.runs[0].text == 'Owned entry'
    assert source.to_bytes() == before
    output = tmp_path / 'result'
    (target.save_flat_opc if flat else target.save)(output)
    cold = DocxDocument(output).get_child_nodes(copied.node_type, True)[0]
    assert cold.first_paragraph.runs[0].text == 'Owned changed copy'
    assert cold.parent_node is not None

@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
@pytest.mark.parametrize('flat', [False, True])
def test_paragraph_clone_remove_reinsert_keeps_inline_entry_ownership(tmp_path, kind, flat):
    doc = document(tmp_path)
    parent = doc.body.paragraphs[0]
    entry = create(doc, kind)
    parent.append_child(entry)
    copied = parent.clone(True)
    copied_entry = copied.get_child_nodes(entry.node_type, True)[0]
    assert copied_entry is not entry
    assert copied_entry.parent_node is copied
    assert copied_entry.first_paragraph.runs[0].text == 'Owned entry'
    doc.body.append_child(copied)
    assert len(doc.get_child_nodes(entry.node_type, True)) == 2
    copied.remove()
    assert copied_entry.parent_node is copied
    assert len(doc.get_child_nodes(entry.node_type, True)) == 1
    doc.body.append_child(copied)
    output = tmp_path / 'result'
    (doc.save_flat_opc if flat else doc.save)(output)
    cold = DocxDocument(output)
    actual = cold.get_child_nodes(entry.node_type, True)
    assert len(actual) == 2
    assert all(item in item.parent_node.child_nodes for item in actual)

@pytest.mark.parametrize('flat', [False, True])
def test_comment_ranges_keep_warm_ids_but_filter_orphans_on_output(tmp_path, flat):
    doc = document(tmp_path)
    parent = doc.body.paragraphs[0]
    comment = create(doc, 'comment')
    parent.append_child(comment)
    start = doc.create_comment_range_start(comment.id)
    end = doc.create_comment_range_end(comment.id)
    parent.insert_before(start, parent.child_nodes[0])
    parent.insert_before(end, comment)
    assert start.parent_node is end.parent_node is parent
    assert parent.child_nodes[-2:] == (end, comment)
    comment.id = 17
    output = tmp_path / 'result'
    (doc.save_flat_opc if flat else doc.save)(output)
    assert (comment.id, start.id, end.id) == (1, 0, 0)
    assert start.parent_node is end.parent_node is None
    (doc.save_flat_opc if flat else doc.save)(output)
    assert comment.id == 0
    cold = DocxDocument(output)
    assert cold.get_child_nodes(aw.NodeType.COMMENT_RANGE_START, True) == []
    assert cold.get_child_nodes(aw.NodeType.COMMENT_RANGE_END, True) == []

@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
@pytest.mark.parametrize('mode', list(ImportFormatMode))
@pytest.mark.parametrize('flat', [False, True])
def test_cross_document_paragraph_bundle_import_keeps_content_and_identity(tmp_path, kind, mode, flat):
    source = document(tmp_path)
    paragraph = source.body.paragraphs[0]
    entry = create(source, kind)
    paragraph.append_child(entry)
    before = source.to_bytes()
    target = document(tmp_path)
    copied = target.import_node(paragraph, True, mode)
    imported = copied.get_child_nodes(entry.node_type, True)[0]
    assert imported.parent_node is copied
    assert imported.owner_document is target
    assert imported.first_paragraph.runs[0].text == 'Owned entry'
    target.body.append_child(copied)
    assert source.to_bytes() == before
    assert entry.parent_node is paragraph
    output = tmp_path / 'result'
    for _ in range(2):
        (target.save_flat_opc if flat else target.save)(output)
        cold = DocxDocument(output).get_child_nodes(entry.node_type, True)
        assert len(cold) == 1
        assert cold[0].first_paragraph.runs[0].text == 'Owned entry'

@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
@pytest.mark.parametrize('flat', [False, True])
def test_minimal_package_without_relationships_supports_new_entry(tmp_path, kind, flat):
    doc = document(tmp_path)
    parts = dict(doc._package._payloads)
    parts.pop('word/_rels/document.xml.rels')
    source = tmp_path / 'minimal.docx'
    with ZipFile(source, 'w') as z:
        for name, data in parts.items():
            z.writestr(name, data)
    doc = DocxDocument(source)
    node = create(doc, kind)
    doc.body.paragraphs[0].append_child(node)
    cold = DocxDocument(BytesIO(doc.to_flat_opc() if flat else doc.to_bytes()))
    assert len(cold.get_child_nodes(node.node_type, True)) == 1

@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
def test_unsupported_annotation_resources_fail_before_target_mutation(tmp_path, kind):
    source = document(tmp_path)
    node = create(source, kind)
    source.body.paragraphs[0].append_child(node)
    element = node.first_paragraph._element
    link = element.ownerDocument.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:hyperlink')
    link.setAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'r:id', 'owned-missing')
    element.appendChild(link)
    target = document(tmp_path)
    before = target.to_bytes()
    for subject in [node, source.body.paragraphs[0]]:
        with pytest.raises(NotImplementedError, match='resource'):
            target.import_node(subject, True, ImportFormatMode.USE_DESTINATION_STYLES)
        assert target.to_bytes() == before
        assert target.get_child_nodes(node.node_type, True) == []

@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
@pytest.mark.parametrize('flat', [False, True])
def test_custom_reference_mark_preserves_warm_value_and_actual_cold_text(tmp_path, kind, flat):
    doc = document(tmp_path)
    node = create(doc, kind)
    doc.body.paragraphs[0].append_child(node)
    node.reference_mark = 'Owned*'
    assert node.is_auto is False
    output = tmp_path / 'result'
    (doc.save_flat_opc if flat else doc.save)(output)
    cold = DocxDocument(output).get_child_nodes(aw.NodeType.FOOTNOTE, True)[0]
    assert cold.reference_mark == 'Owned*'
    assert cold.is_auto is False
    node.reference_mark = None
    assert node.reference_mark is None and node.is_auto is True
    (doc.save_flat_opc if flat else doc.save)(output)
    assert DocxDocument(output).get_child_nodes(aw.NodeType.FOOTNOTE, True)[0].reference_mark == ''

@pytest.mark.parametrize('outside', ['header', 'nested'])
@pytest.mark.parametrize('flat', [False, True])
def test_successful_save_removes_unsupported_outside_note_anchor_only_on_commit(tmp_path, outside, flat):
    from .test_dom_story_row_height import story_package
    if outside == 'header':
        path = tmp_path / 'header.docx'
        name, _ = story_package(path, 'header')
        doc = DocxDocument(path)
        parent = doc.story(name).get_child_nodes(aw.NodeType.PARAGRAPH, True)[0]
    else:
        doc = document(tmp_path)
        outer = create(doc, 'footnote')
        doc.body.paragraphs[0].append_child(outer)
        parent = outer.first_paragraph
    inner = create(doc, 'footnote')
    parent.append_child(inner)
    assert inner.parent_node is parent
    output = tmp_path / 'result'
    pure = doc.to_flat_opc() if flat else doc.to_bytes()
    assert inner.parent_node is parent
    (doc.save_flat_opc if flat else doc.save)(output)
    assert inner.parent_node is None
    assert inner not in parent.child_nodes
    assert inner.count == 1
    cold = DocxDocument(BytesIO(pure))
    assert len(cold.get_child_nodes(aw.NodeType.FOOTNOTE, True)) == (1 if outside == 'nested' else 0)

@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
def test_detached_paragraph_bundle_never_enters_pure_output(tmp_path, kind):
    doc = document(tmp_path)
    entry = create(doc, kind)
    paragraph = doc.body.paragraphs[0]
    paragraph.append_child(entry)
    copied = paragraph.clone(True)
    copied_entry = copied.get_child_nodes(entry.node_type, True)[0]
    cold = DocxDocument(BytesIO(doc.to_bytes()))
    assert len(cold.get_child_nodes(entry.node_type, True)) == 1
    assert copied_entry.parent_node is copied
    assert copied.parent_node is None

@pytest.mark.parametrize('operation', ['clone', 'import'])
def test_paragraph_comment_bundle_allocates_unique_warm_ids_and_ranges(tmp_path, operation):
    source = document(tmp_path)
    p = source.body.paragraphs[0]
    comment = create(source, 'comment')
    p.append_child(comment)
    start = source.create_comment_range_start(comment.id)
    end = source.create_comment_range_end(comment.id)
    p.insert_before(start, p.child_nodes[0]);p.insert_before(end, comment)
    if operation == 'clone':
        target = source
        copied = p.clone(True)
    else:
        target = document(tmp_path)
        target.body.paragraphs[0].append_child(create(target, 'comment'))
        copied = target.import_node(p, True, ImportFormatMode.USE_DESTINATION_STYLES)
    imported = copied.get_child_nodes(aw.NodeType.COMMENT, True)[0]
    assert imported.id == 1
    assert copied.get_child_nodes(aw.NodeType.COMMENT_RANGE_START, True)[0].id == 1
    assert copied.get_child_nodes(aw.NodeType.COMMENT_RANGE_END, True)[0].id == 1
    target.body.append_child(copied)
    target.save(tmp_path / 'result.docx')
    assert imported.id == 1
    assert copied.get_child_nodes(aw.NodeType.COMMENT_RANGE_START, True)[0].id == 1


def test_comment_constructor_id_does_not_follow_edited_literal_id(tmp_path):
    doc = document(tmp_path)
    first = create(doc, 'comment')
    first.id = 17
    second = create(doc, 'comment')
    assert second.id == 1

@pytest.mark.parametrize('bundle', [False, True])
@pytest.mark.parametrize('operation', ['clone', 'import'])
@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
def test_custom_mark_survives_entry_and_paragraph_copy(tmp_path, bundle, operation, kind):
    source = document(tmp_path)
    entry = create(source, kind)
    parent = source.body.paragraphs[0]
    parent.append_child(entry)
    entry.reference_mark = 'Owned*'
    subject = parent if bundle else entry
    if operation == 'clone':
        target = source
        copied = subject.clone(True)
    else:
        target = document(tmp_path)
        copied = target.import_node(subject, True, ImportFormatMode.USE_DESTINATION_STYLES)
    if bundle:
        actual = copied.get_child_nodes(aw.NodeType.FOOTNOTE, True)[0]
        target.body.append_child(copied)
    else:
        actual = copied
        target.body.paragraphs[1].append_child(copied)
    assert actual.reference_mark == 'Owned*'
    output = tmp_path / 'result.docx'
    target.save(output)
    assert DocxDocument(output).get_child_nodes(aw.NodeType.FOOTNOTE, True)[-1].reference_mark == 'Owned*'

@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
@pytest.mark.parametrize('changed', ['entry', 'anchor'])
def test_decimal_annotation_ids_match_when_lexical_forms_differ(tmp_path, kind, changed):
    source = tmp_path / 'input.docx'
    name, parts = additional_story(source, kind)
    part = name if changed == 'entry' else 'word/document.xml'
    parts[part] = parts[part].replace(b'w:id="1"', b'w:id="+001"', 1)
    with ZipFile(source, 'w') as archive:
        for key, value in parts.items():
            archive.writestr(key, value)
    doc = DocxDocument(source)
    original = doc.to_bytes()
    nodes = doc.get_child_nodes(aw.NodeType.COMMENT if kind == 'comments' else aw.NodeType.FOOTNOTE, True)
    assert len(nodes) == 1
    assert nodes[0] in nodes[0].parent_node.child_nodes
    assert doc.to_bytes() == original

@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
def test_bad_annotation_identifier_does_not_echo_sensitive_source(tmp_path, kind):
    source = tmp_path / 'input.docx'
    name, parts = additional_story(source, kind)
    parts[name] = parts[name].replace(b'w:id="1"', b'w:id="Owned sensitive identifier"', 1)
    with ZipFile(source, 'w') as archive:
        for key, value in parts.items():
            archive.writestr(key, value)
    doc = DocxDocument(source)
    before = doc.to_bytes()
    with pytest.raises(ValueError, match='Invalid annotation identifier') as error:
        doc.get_child_nodes(aw.NodeType.ANY, True)
    assert 'Owned' not in str(error.value)
    assert doc.to_bytes() == before

@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
def test_annotation_cycle_is_rejected_before_graph_or_content_changes(tmp_path, kind):
    doc = document(tmp_path)
    node = create(doc, kind)
    before = doc.to_bytes();content = node.xml
    with pytest.raises(ValueError, match='descendant'):
        node.first_paragraph.append_child(node)
    assert node.xml == content
    assert node.parent_node is None
    assert doc.to_bytes() == before

@pytest.mark.parametrize('value', [None, True, 0, 'FOOTNOTE'])
def test_footnote_factory_requires_typed_enum(tmp_path, value):
    doc = document(tmp_path)
    before = doc.to_bytes()
    with pytest.raises(TypeError, match='FootnoteType'):
        doc.create_footnote(value)
    assert doc.to_bytes() == before

@pytest.mark.parametrize('value,error', [(None,TypeError),(True,TypeError),(1.5,TypeError),('7',TypeError),(2**31,OverflowError)])
def test_comment_and_range_id_invalid_values_are_atomic(tmp_path, value, error):
    doc = document(tmp_path)
    node = create(doc, 'comment')
    doc.body.paragraphs[0].append_child(node)
    before = doc.to_bytes();initial=node.id
    with pytest.raises(error):node.id=value
    with pytest.raises(error):doc.create_comment_range_start(value)
    with pytest.raises(error):doc.create_comment_range_end(value)
    assert node.id == initial and doc.to_bytes() == before

@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
def test_loaded_table_entry_has_native_minimum_paragraph_without_changing_projection(tmp_path, kind):
    path=tmp_path/'input.docx';name,_=additional_story(path,kind)
    doc=DocxDocument(path);before=doc.to_bytes()
    storage=doc.story(name).child_nodes[0]
    assert storage.count == 2
    assert storage.first_paragraph is not None
    assert storage.first_paragraph.runs == ()
    logical=storage.canonical_node
    assert logical.count == 2 and logical.first_paragraph is not None
    assert doc.to_bytes() == before

@pytest.mark.parametrize('kind', ['footnotes', 'endnotes', 'comments'])
def test_mixed_anchor_run_preserves_comments_pi_and_read_only_row_projection(tmp_path, kind):
    import re
    path=tmp_path/'input.docx';name,parts=additional_story(path,kind)
    ref={'footnotes':'footnoteReference','endnotes':'endnoteReference','comments':'commentReference'}[kind]
    table=b'<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="40000" w:hRule="exact"/></w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
    body=('<w:body><w:p><w:r><w:t>Owned before</w:t><!--Owned marker--><?owned preserve?>'+f'<w:{ref} w:id="1"/>'+'<w:t>Owned after</w:t></w:r></w:p>').encode()+table+b'<w:sectPr/></w:body>'
    parts['word/document.xml']=re.sub(rb'<w:body>.*?</w:body>',body,parts['word/document.xml'],flags=re.S)
    with ZipFile(path,'w')as z:
        for key,value in parts.items():z.writestr(key,value)
    doc=DocxDocument(path);before=doc.to_bytes()
    paragraph=doc.body.paragraphs[0]
    expected=aw.NodeType.COMMENT if kind=='comments'else aw.NodeType.FOOTNOTE
    assert [n.node_type for n in paragraph.child_nodes]==[aw.NodeType.RUN,expected,aw.NodeType.RUN]
    assert [n.text for n in paragraph.runs]==['Owned before','Owned after']
    assert doc.to_bytes()==before
    paragraph.runs[0].text='Owned changed'
    output=tmp_path/'result.docx';doc.save(output)
    with ZipFile(output)as z:
        main=z.read('word/document.xml')
    assert b'<!--Owned marker-->' in main and b'<?owned preserve?>' in main
    cold=DocxDocument(output)
    assert [n.text for n in cold.body.paragraphs[0].runs]==['Owned changed','Owned after']

@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
@pytest.mark.parametrize('flat', [False, True])
@pytest.mark.parametrize('return_to_body', [False, True])
def test_note_move_between_body_and_header_preserves_unique_parent_and_safe_save(tmp_path, kind, flat, return_to_body):
    from .test_dom_story_row_height import story_package
    path=tmp_path/'header.docx';name,_=story_package(path,'header')
    doc=DocxDocument(path);body=doc.body.paragraphs[0]
    header=doc.story(name).get_child_nodes(aw.NodeType.PARAGRAPH,True)[0]
    entry=create(doc,kind);body.append_child(entry)
    header.append_child(entry)
    assert entry.parent_node is header and entry in header.child_nodes
    assert entry not in body.child_nodes
    if return_to_body:body.append_child(entry)
    output=tmp_path/'result'
    (doc.save_flat_opc if flat else doc.save)(output)
    cold=DocxDocument(output)
    assert len(cold.get_child_nodes(entry.node_type,True))==(1 if return_to_body else 0)
    assert entry.parent_node is (body if return_to_body else None)

@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
@pytest.mark.parametrize('operation', ['clone', 'import'])
def test_entry_copy_preserves_inline_anchor_direct_format(tmp_path, kind, operation):
    from xml.etree.ElementTree import fromstring
    source=document(tmp_path);node=create(source,kind);source.body.paragraphs[0].append_child(node)
    # Owned input puts formatting on the inline marker, independently of note content.
    xml=source.to_bytes();parts={}
    with ZipFile(BytesIO(xml))as z:parts={name:z.read(name)for name in z.namelist()}
    ref={'footnote':'footnoteReference','endnote':'endnoteReference','comment':'commentReference'}[kind]
    parts['word/document.xml']=parts['word/document.xml'].replace(('<w:r><w:'+ref).encode(),('<w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:'+ref).encode(),1)
    path=tmp_path/'styled.docx'
    with ZipFile(path,'w')as z:
        for name,data in parts.items():z.writestr(name,data)
    source=DocxDocument(path);node=source.get_child_nodes(aw.NodeType.COMMENT if kind=='comment'else aw.NodeType.FOOTNOTE,True)[0]
    if operation=='clone':target=source;copied=node.clone(True)
    else:target=document(tmp_path);copied=target.import_node(node,True,ImportFormatMode.USE_DESTINATION_STYLES)
    target.body.paragraphs[1].append_child(copied)
    output=tmp_path/'result.docx';target.save(output)
    root=fromstring(target.part_xml('word/document.xml'));w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    anchors=[run for run in root.iter(w+'r')if run.find(w+ref)is not None]
    assert len(anchors)==(2 if operation=='clone'else 1)
    for run in anchors:
        assert run.find(w+'rPr/'+w+'b')is not None
        assert run.find(w+'rPr/'+w+'sz').get(w+'val')=='36'


@pytest.mark.parametrize('kind', ['footnote', 'endnote', 'comment'])
@pytest.mark.parametrize('operation', ['move', 'clone'])
@pytest.mark.parametrize('flat', [False, True])
def test_detached_entry_keeps_marker_format_and_custom_mark(tmp_path, kind, operation, flat):
    from xml.etree.ElementTree import fromstring
    doc = document(tmp_path)
    note = create(doc, kind)
    doc.body.paragraphs[0].append_child(note)
    ref = {'footnote': 'footnoteReference', 'endnote': 'endnoteReference', 'comment': 'commentReference'}[kind]
    with ZipFile(BytesIO(doc.to_bytes())) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    parts['word/document.xml'] = parts['word/document.xml'].replace(
        ('<w:r><w:' + ref).encode(),
        ('<w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:' + ref).encode(), 1)
    source = tmp_path / 'marker.docx'
    with ZipFile(source, 'w') as package:
        for name, value in parts.items():
            package.writestr(name, value)
    doc = DocxDocument(source)
    note = doc.get_child_nodes(aw.NodeType.COMMENT if kind == 'comment' else aw.NodeType.FOOTNOTE, True)[0]
    if kind != 'comment':
        note.reference_mark = 'Owned*'
    note.remove()
    assert note.parent_node is None
    if kind != 'comment':
        assert note.reference_mark == 'Owned*'
    moved = note.clone(True) if operation == 'clone' else note
    doc.body.paragraphs[1].append_child(moved)
    output = tmp_path / 'result'
    (doc.save_flat_opc if flat else doc.save)(output)
    cold = DocxDocument(output)
    root = fromstring(cold.part_xml('word/document.xml'))
    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    anchors = [run for run in root.iter(w + 'r') if run.find(w + ref) is not None]
    assert len(anchors) == 1
    assert anchors[0].find(w + 'rPr/' + w + 'b') is not None
    assert anchors[0].find(w + 'rPr/' + w + 'sz').get(w + 'val') == '36'
    if kind != 'comment':
        assert cold.get_child_nodes(aw.NodeType.FOOTNOTE, True)[0].reference_mark == 'Owned*'
