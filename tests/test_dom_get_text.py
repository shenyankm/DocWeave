"""Word control characters and ordered logical stories, independently calibrated."""
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

import dockiln as dk
from dockiln import light_document_model as ldm
from dockiln._flat_opc import encode
from dockiln.writers.docx import LdmDocxWriter
from dockiln.dom import DocxDocument
from tests.test_comment_threads import owned_import_thread
from tests.test_conversion_comments import W, owned_comment, package


def controlled_document(inline, *, section=False):
    with ZipFile(BytesIO(owned_comment())) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    root = ET.fromstring(parts['word/document.xml'])
    paragraph = ET.fromstring('<w:p xmlns:w="' + W[1:-1] + '">' + inline + '</w:p>')
    if section:
        ET.SubElement(ET.SubElement(paragraph, W + 'pPr'), W + 'sectPr')
    root.find(W + 'body').insert(0, paragraph)
    parts['word/document.xml'] = ET.tostring(root)
    return package(parts)


@pytest.mark.parametrize('flat', [False, True])
@pytest.mark.parametrize(('inline', 'expected'), [
    ('<w:r><w:t xml:space="preserve"> OWNED </w:t><w:tab/><w:br/><w:br w:type="page"/><w:br w:type="column"/><w:cr/><w:noBreakHyphen/><w:softHyphen/></w:r>', ' OWNED \t\x0b\x0c\x0e\r\x1e\x1f\r'),
    ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>OWNED RESULT</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>', '\x13 DATE \x14OWNED RESULT\x15\r'),
    ('<w:fldSimple w:instr=" DATE "><w:r><w:t>OWNED RESULT</w:t></w:r></w:fldSimple>', '\x13 DATE \x14OWNED RESULT\x15\r'),
    ('<w:bookmarkStart w:id="3" w:name="Owned"/><w:r><w:t>OWNED</w:t></w:r><w:bookmarkEnd w:id="3"/>', 'OWNED\r'),
    ('<w:ins w:id="1"><w:r><w:t>OWNED INSERT</w:t></w:r></w:ins><w:del w:id="2"><w:r><w:delText>DELETE</w:delText></w:r></w:del>', 'OWNED INSERTDELETE\r'),
    ('<w:r><w:t>OWNED</w:t><w:sym w:font="Wingdings" w:char="F021"/></w:r>', 'OWNED\uf021\r'),
])
def test_paragraph_control_characters_and_saves(flat, inline, expected):
    raw = controlled_document(inline)
    document = DocxDocument(BytesIO(encode(raw) if flat else raw))
    for _ in range(2):
        assert document.body.paragraphs[0].get_text() == expected
        document = DocxDocument(BytesIO(document.to_flat_opc() if flat else document.to_bytes()))


def test_last_paragraph_section_and_cell_terminators():
    document = DocxDocument(BytesIO(controlled_document('<w:r><w:t>OWNED SECTION</w:t></w:r>', section=True)))
    assert document.body.paragraphs[0].get_text() == 'OWNED SECTION\x0c'
    assert document.body.paragraphs[-1].get_text().endswith('OWNED AFTER\x0c')
    table_document = DocxDocument(BytesIO(owned_comment('table')))
    table = table_document.body.tables[0]
    cell = table.rows[0].cells[0]
    assert cell.paragraphs[-1].get_text().endswith('\x07')
    assert cell.get_text().endswith('\x07')
    assert table.rows[0].get_text() == cell.get_text() + '\x07'
    assert table.get_text() == table.rows[0].get_text()
    assert table_document.body.get_text().endswith('\x07\x07\x0c')


@pytest.mark.parametrize('flat', [False, True])
def test_inline_comment_content_edits_and_getter_is_not_raw_text(flat):
    document = DocxDocument(BytesIO(owned_comment()))
    comment = document.get_child_nodes(dk.NodeType.COMMENT, True)[0]
    assert comment.get_text() == '\x05OWNED COMMENT α\r'
    assert 'OWNED COMMENT α' not in document.body.paragraphs[0].text
    comment.paragraphs[0].runs[-1].text = 'OWNED EDITED'
    for _ in range(2):
        assert document.get_text() == 'OWNED BEFORE OWNED ANCHOR\x05OWNED EDITED\r OWNED AFTER\x0c'
        document = DocxDocument(BytesIO(document.to_flat_opc() if flat else document.to_bytes()))


def test_get_text_clone_import_and_node_context():
    document = DocxDocument(BytesIO(controlled_document('<w:r><w:t>OWNED FIRST</w:t></w:r>')))
    paragraph = document.body.paragraphs[0]
    copied = paragraph.clone(True)
    assert copied.get_text() == 'OWNED FIRST\r'
    target = DocxDocument(BytesIO(owned_import_thread()))
    imported = target.import_node(paragraph, True)
    target.body.append_child(imported)
    assert imported.get_text() == 'OWNED FIRST\x0c'
    assert target.body.paragraphs[0].get_text().endswith('OWNED COMMENT 2\r\r')


@pytest.mark.parametrize('order', [list(range(6)), list(reversed(range(6))), [2, 0, 4, 3, 1, 5]])
def test_document_section_story_order_and_no_inherited_duplication(order):
    def paragraph(text):
        return ldm.Paragraph(children=[ldm.Run(text=text)])
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[paragraph('OWNED BODY 0')]),
        headers_footers=[ldm.HeaderFooter(header_footer_type=k, children=[paragraph(f'OWNED HF {k}')]) for k in order]),
        ldm.Section(body=ldm.Body(children=[paragraph('OWNED BODY 1')]))])
    document = DocxDocument(BytesIO(LdmDocxWriter().write_to_bytes(model)))
    assert document.get_text() == ('OWNED BODY 0\x0c' + ''.join(f'OWNED HF {k}\r' for k in [5, 1, 3, 4, 0, 2]) + 'OWNED BODY 1\x0c')
    assert document.body.get_text() == 'OWNED BODY 0\x0cOWNED BODY 1\x0c'


@pytest.mark.parametrize('history', [None, '0', 'false', 'off', '1', 'true', 'on'])
def test_hyperlink_field_code_includes_history_and_anchor(history):
    document = DocxDocument(BytesIO(owned_comment()))
    link = document.body.paragraphs[0].add_hyperlink('OWNED LINK', 'https://example.invalid/a#bookmark')
    if history is not None:
        link._element.setAttributeNS(W[1:-1], 'w:history', history)
    expected = '\x13 HYPERLINK "https://example.invalid/a" \\l "bookmark" '
    if history not in {'1', 'true', 'on'}:
        expected += ' \\h '
    assert link.get_text() == expected + '\x14OWNED LINK\x15'


@pytest.mark.parametrize('inline', [
    '<w:r><w:customUnknown><w:t>SENSITIVE BODY</w:t></w:customUnknown></w:r>',
    '<w:customUnknown><w:r><w:t>SENSITIVE BODY</w:t></w:r></w:customUnknown>',
])
def test_unsupported_structure_fails_without_partial_text_or_diagnostic_body(inline):
    document = DocxDocument(BytesIO(controlled_document(inline)))
    with pytest.raises(NotImplementedError) as error:
        document.get_text()
    assert 'SENSITIVE BODY' not in str(error.value)


@pytest.mark.parametrize('inline', ['<w:r><w:br w:type="invalid"/></w:r>', '<w:r><w:fldChar w:fldCharType="invalid"/></w:r>', '<w:r><w:sym w:char="invalid"/></w:r>', '<w:r><w:t><w:t>SENSITIVE BODY</w:t></w:t></w:r>'])
def test_malformed_word_controls_report_clear_errors(inline):
    document = DocxDocument(BytesIO(controlled_document(inline)))
    with pytest.raises(ValueError) as error:
        document.get_text()
    assert 'SENSITIVE BODY' not in str(error.value)


@pytest.mark.parametrize('kind', ['document', 'paragraph', 'run'])
def test_get_text_has_no_formatting_arguments(kind):
    document = DocxDocument(BytesIO(owned_comment()))
    node = document if kind == 'document' else document.body.paragraphs[0] if kind == 'paragraph' else document.body.paragraphs[0].runs[0]
    with pytest.raises(TypeError):
        node.get_text(None)


@pytest.mark.parametrize('kind', ['footnote', 'endnote'])
def test_created_note_text_does_not_invent_reference_character(tmp_path, kind):
    from tests.test_dom_note_lifecycle import create, document
    doc = document(tmp_path)
    note = create(doc, kind, fill=False)
    assert note.get_text() == ''
    note.ensure_minimum()
    assert note.get_text() == '\r'
    note.first_paragraph.append_child(doc.create_run('OWNED CREATED', part_name=note.part_name))
    assert note.get_text() == 'OWNED CREATED\r'
    doc.body.paragraphs[0].append_child(note)
    assert note.get_text() == 'OWNED CREATED\r'


def test_logical_story_cycle_reports_error_without_recursion(tmp_path):
    from tests.test_dom_additional_story_rows import additional_story
    name, parts = additional_story(tmp_path / 'cycle.docx', 'footnotes')
    body = ET.fromstring(parts['word/document.xml'])
    for run in body.iter(W + 'r'):
        for child in list(run):
            if child.tag == W + 'footnoteReference':
                run.remove(child)
    parts['word/document.xml'] = ET.tostring(body)
    notes = ET.fromstring(parts[name])
    entry = notes.find(W + 'footnote')
    paragraph = ET.SubElement(entry, W + 'p')
    run = ET.SubElement(paragraph, W + 'r')
    ET.SubElement(run, W + 'footnoteReference', {W + 'id': '1'})
    parts[name] = ET.tostring(notes)
    document = DocxDocument(BytesIO(package(parts)))
    note = document.story(name).child_nodes[0].canonical_node
    with pytest.raises(ValueError, match='Cyclic logical story reference'):
        note.get_text()


@pytest.mark.parametrize('unknown', [False, True])
def test_getter_preserves_bound_nodes_and_package_payloads(unknown):
    raw = controlled_document('<w:customUnknown><w:r><w:t>SENSITIVE BODY</w:t></w:r></w:customUnknown>') if unknown else owned_comment()
    document = DocxDocument(BytesIO(raw))
    # Existing annotation binding normalizes logical nodes; text reading adds no mutation.
    nodes = tuple(document.get_child_nodes(dk.NodeType.ANY, True))
    before = tuple((node.part_name, node.node_type, node.xml) for node in nodes)
    payloads = dict(document._package._payloads)
    dirty = set(document._package._dirty)
    if unknown:
        with pytest.raises(NotImplementedError):
            document.get_text()
    else:
        document.get_text()
        for node in nodes:
            try:
                node.get_text()
            except NotImplementedError:
                # Raw property/extension wrappers are deliberately outside the text contract.
                pass
    assert tuple((node.part_name, node.node_type, node.xml) for node in nodes) == before
    assert document._package._payloads == payloads
    assert document._package._dirty == dirty


def test_empty_body_without_section_properties_has_section_terminator():
    with ZipFile(BytesIO(owned_comment())) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    root = ET.fromstring(parts['word/document.xml'])
    root.find(W + 'body').clear()
    parts['word/document.xml'] = ET.tostring(root)
    document = DocxDocument(BytesIO(package(parts)))
    assert document.body.get_text() == '\x0c'
    assert document.get_text() == '\x0c'
