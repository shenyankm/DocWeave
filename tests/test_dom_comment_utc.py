"""Owned modern comment dates, readonly UTC projection and atomic lifecycle."""
from datetime import datetime, timedelta, timezone
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as E

import pytest

from aspose.words_foss import NodeType
from aspose.words_foss.dom import DocxDocument
from aspose.words_foss.dom.comment_utc import W14, CID, CEX, ROLES, get_utc, utc_date
from .test_dom_comment_metadata import source_date, comment

REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
UTC = timezone.utc


def modern(tmp_path, raw='2030-01-02T03:04:05Z', *, omit=None, custom=False):
    source = source_date(tmp_path, '2026-10-11T12:34:56Z')
    with ZipFile(BytesIO(source.to_bytes())) as z:parts = {n: z.read(n) for n in z.namelist()}
    tree = E.fromstring(parts['word/comments.xml'])
    tree.find('.//{' + W + '}p').set('{' + W14 + '}paraId', '12345678')
    parts['word/comments.xml'] = E.tostring(tree)
    rels, types = E.fromstring(parts['word/_rels/document.xml.rels']), E.fromstring(parts['[Content_Types].xml'])
    for role, (ns, local, rel_type) in ROLES.items():
        if role == omit:continue
        tree = E.Element('{' + ns + '}' + role)
        item = E.SubElement(tree, '{' + ns + '}' + local, {'{' + ns + '}durableId': '23456789'})
        if role == 'commentsIds':item.set('{' + ns + '}paraId', '12345678')
        elif raw is not None:item.set('{' + ns + '}dateUtc', raw)
        name = 'word/owned/' + role + '.xml' if custom else 'word/' + role + '.xml'
        parts[name] = E.tostring(tree)
        E.SubElement(types, '{' + CT + '}Override', PartName='/' + name,
                     ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.' + role + '+xml')
        E.SubElement(rels, '{' + REL + '}Relationship', Id='owned-' + role, Type=rel_type, Target=name.removeprefix('word/'))
    parts['word/_rels/document.xml.rels'], parts['[Content_Types].xml'] = E.tostring(rels), E.tostring(types)
    path = tmp_path / 'modern.docx'
    with ZipFile(path, 'w') as z:
        for n, payload in parts.items():z.writestr(n, payload)
    return DocxDocument(path)


@pytest.mark.parametrize('raw, expected', [
    ('2030-01-02T03:04:05Z', datetime(2030, 1, 2, 3, 4, 5, tzinfo=UTC)),
    ('2030-01-02T03:04:05+08:00', datetime(2030, 1, 1, 19, 4, 5, tzinfo=UTC)),
    ('2030-01-02T03:04:05', datetime(2030, 1, 2, 3, 4, 5)),
    ('2030-01-02T03:04:05.1234567Z', datetime(2030, 1, 2, 3, 4, 5, 123456, tzinfo=UTC)),
    ('0001-01-01T00:00:00.0000001Z', datetime.min.replace(tzinfo=UTC)),
    ('0001-01-01T00:00:01+08:00', datetime(1, 1, 1, 16, 0, 1, tzinfo=UTC)),
    (None, None), ('', None), ('Owned invalid date', None), ('0001-01-01T00:00:00Z', None),
    ('9999-12-31T23:59:59-05:00', None), ('2030-01-02T24:00:00Z', None),
])
@pytest.mark.parametrize('flat', [False, True])
def test_modern_utc_source_and_two_cold_saves(tmp_path, raw, expected, flat):
    doc = modern(tmp_path, raw, custom=True);node = comment(doc)
    fallback = get_utc(comment(source_date(tmp_path, '2026-10-11T12:34:56Z'))._element)
    expected = fallback if expected is None else expected
    held = node.date_time_utc
    assert held == expected and held.tzinfo == expected.tzinfo
    assert node.date_time.year == 2026
    output = tmp_path / 'result'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(output)
        assert node.date_time_utc == held and node.date_time_utc.tzinfo == held.tzinfo
        cold = comment(DocxDocument(output))
        cold_expected = expected.replace(microsecond=0, tzinfo=UTC)
        if cold_expected == datetime.min.replace(tzinfo=UTC):
            # A zero modern UTC date is absent; the legacy date becomes the fallback.
            cold_expected = fallback
        assert cold.date_time_utc == cold_expected


@pytest.mark.parametrize('value', [None, 0, datetime.min, datetime(2040, 1, 2, tzinfo=UTC)])
def test_utc_is_readonly_for_all_values(tmp_path, value):
    doc = modern(tmp_path);node = comment(doc);before = doc.to_bytes();held = node.date_time_utc
    with pytest.raises(AttributeError):node.date_time_utc = value
    assert doc.to_bytes() == before and node.date_time_utc is held


@pytest.mark.parametrize('value', [datetime.min, datetime(2040, 1, 2, 3, 4, 5, 123456),
                                  datetime(2040, 1, 2, 3, 4, 5, 123456, tzinfo=UTC)])
@pytest.mark.parametrize('flat', [False, True])
def test_date_setter_replaces_modern_utc_and_keeps_warm_precision(tmp_path, value, flat):
    doc = modern(tmp_path);node = comment(doc);node.date_time = value
    expected = get_utc(node._element)
    assert node.date_time_utc == expected and node.date_time_utc.year == value.year
    output = tmp_path / 'result'
    (doc.save_flat_opc if flat else doc.save)(output)
    assert node.date_time_utc == expected
    cold = comment(DocxDocument(output));assert cold.date_time_utc.replace(tzinfo=None) == expected.replace(tzinfo=None, microsecond=0)


@pytest.mark.parametrize('operation', ['clone', 'shallow', 'import', 'paragraph-clone'])
@pytest.mark.parametrize('flat', [False, True])
def test_modern_utc_copy_and_import_have_independent_identity(tmp_path, operation, flat):
    source = modern(tmp_path, '2030-01-02T03:04:05.1234567Z');node = comment(source);held = node.date_time_utc
    doc = source
    if operation == 'paragraph-clone':
        paragraph = source.body.paragraphs[0].clone(True);source.body.append_child(paragraph)
        copied = next(n for n in paragraph.child_nodes if n.node_type == NodeType.COMMENT)
    else:
        if operation == 'import':
            doc = source_date(tmp_path, None);copied = doc.import_node(node, True)
        else:copied = node.clone(operation != 'shallow')
        doc.body.paragraphs[1].append_child(copied)
    assert copied.date_time_utc == held
    output = tmp_path / 'result'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(output)
        assert copied.date_time_utc == held
        cold = DocxDocument(output).get_child_nodes(NodeType.COMMENT, True)
        assert cold[-1].date_time_utc == held.replace(microsecond=0)
    copied.date_time = datetime(2040, 1, 2, tzinfo=UTC)
    assert node.date_time_utc == held and copied.date_time_utc.year == 2040


@pytest.mark.parametrize('flat', [False, True])
def test_modern_utc_projection_runs_without_prior_story_access(tmp_path, flat):
    doc = modern(tmp_path, '2030-01-02T03:04:05.1234567Z');assert doc._note_graph is None
    output = tmp_path / 'result';(doc.save_flat_opc if flat else doc.save)(output)
    assert comment(DocxDocument(output)).date_time_utc == datetime(2030, 1, 2, 3, 4, 5, tzinfo=UTC)


@pytest.mark.parametrize('flat', [False, True])
def test_new_modern_parts_atomic_failure_and_budget_preserve_held_source(tmp_path, monkeypatch, flat):
    from aspose.words_foss import _io
    source = modern(tmp_path);original = comment(source);target = source_date(tmp_path, None)
    copied = target.import_node(original, True);target.body.paragraphs[1].append_child(copied)
    before = target.to_bytes();names = target.part_names;held = copied.date_time_utc;pid = copied.paragraphs[-1]._element.getAttributeNS(W14, 'paraId')
    path = tmp_path / 'result';path.write_bytes(b'Owned prior output')
    def fail(*args):raise OSError('Owned publish failure')
    monkeypatch.setattr(_io.os, 'replace', fail)
    with pytest.raises(OSError):(target.save_flat_opc if flat else target.save)(path)
    assert path.read_bytes() == b'Owned prior output' and target.part_names == names and target.to_bytes() == before
    assert copied.date_time_utc is held and copied.paragraphs[-1]._element.getAttributeNS(W14, 'paraId') == pid
    monkeypatch.setattr(_io, 'MAX_PART_BYTES', 20)
    with pytest.raises(ValueError):(target.save_flat_opc if flat else target.save)(path)
    assert path.read_bytes() == b'Owned prior output' and copied.date_time_utc is held


@pytest.mark.parametrize('placement', ['two-first', 'two-last', 'table-last', 'table-inner', 'table-follow'])
@pytest.mark.parametrize('flat', [False, True])
def test_utc_binding_follows_recognized_source_paragraphs_and_last_output_paragraph(tmp_path, placement, flat):
    doc = modern(tmp_path);tree = doc._package.tree('word/comments.xml');entry = tree.documentElement.getElementsByTagNameNS(W, 'comment')[0]
    first = entry.getElementsByTagNameNS(W, 'p')[0]
    if placement.startswith('two'):
        paragraph = tree.createElementNS(W, 'w:p');paragraph.setAttribute('xmlns:w', W);entry.appendChild(paragraph)
        if placement == 'two-last':
            first.removeAttributeNS(W14, 'paraId');paragraph.setAttribute('xmlns:w14', W14);paragraph.setAttributeNS(W14, 'w14:paraId', '12345678')
    else:
        tbl = tree.createElementNS(W, 'w:tbl');tbl.setAttribute('xmlns:w', W)
        tr = tree.createElementNS(W, 'w:tr');tc = tree.createElementNS(W, 'w:tc');p = tree.createElementNS(W, 'w:p');entry.appendChild(tbl);tbl.appendChild(tr);tr.appendChild(tc);tc.appendChild(p)
        if placement == 'table-inner':
            first.removeAttributeNS(W14, 'paraId');p.setAttribute('xmlns:w14', W14);p.setAttributeNS(W14, 'w14:paraId', '12345678')
        if placement == 'table-follow':entry.appendChild(tree.createElementNS(W, first.prefix + ':p'))
    doc._package._dirty.add('word/comments.xml')
    node = comment(doc);expected = datetime(2030, 1, 2, 3, 4, 5, tzinfo=UTC)
    assert node.date_time_utc == expected
    output = tmp_path / 'result';(doc.save_flat_opc if flat else doc.save)(output)
    cold = DocxDocument(output);assert comment(cold).date_time_utc == expected
    root = E.fromstring(cold.part_xml('word/comments.xml'));ids = E.fromstring(cold.part_xml('word/commentsIds.xml'))
    assert ids[0].get('{' + CID + '}paraId') == list(root.iter('{' + W + '}p'))[-1].get('{' + W14 + '}paraId')


@pytest.mark.parametrize('flat', [False, True])
def test_delete_modern_comment_removes_known_utc_records_but_keeps_opaque_metadata(tmp_path, flat):
    doc = modern(tmp_path);node = comment(doc)
    tree = doc._package.tree('word/commentsExtensible.xml');shadow = tree.createElementNS('urn:owned:opaque', 'owned:shadow')
    shadow.setAttribute('xmlns:owned', 'urn:owned:opaque');shadow.setAttribute('dateUtc', 'Owned opaque content');tree.documentElement.appendChild(shadow);doc._package._dirty.add('word/commentsExtensible.xml')
    node.remove();output = tmp_path / 'result';(doc.save_flat_opc if flat else doc.save)(output)
    cold = DocxDocument(output);assert cold.get_child_nodes(NodeType.COMMENT, True) == []
    root = E.fromstring(cold.part_xml('word/commentsExtensible.xml'));assert root.findall('{' + CEX + '}commentExtensible') == []
    assert root.find('{urn:owned:opaque}shadow').get('dateUtc') == 'Owned opaque content'
    assert E.fromstring(cold.part_xml('word/commentsIds.xml')).findall('{' + CID + '}commentId') == []


@pytest.mark.parametrize('raw, expected', [('0001-01-01T00:00:00+08:00', datetime(1, 1, 1, 16, tzinfo=UTC)),
                                        ('0001-01-01T00:00:00Z', None), ('0000-01-01T00:00:00+08:00', None)])
def test_utc_source_conversion_precedes_minimum_fallback(raw, expected):
    assert utc_date(raw) == expected


@pytest.mark.parametrize('value, expected', [(datetime.min, datetime.min),
        (datetime.min.replace(tzinfo=UTC), datetime.min),
        (datetime.max.replace(tzinfo=UTC), datetime.max.replace(tzinfo=UTC))])
def test_fallback_utc_minimum_ignores_datetime_kind(tmp_path, value, expected):
    doc = source_date(tmp_path, None);node = doc.create_comment('', '', value)
    assert node.date_time_utc == expected and node.date_time_utc.tzinfo == expected.tzinfo


def test_typed_utc_projection_retains_extension_list_after_entries(tmp_path):
    doc = modern(tmp_path);tree = doc._package.tree('word/commentsExtensible.xml')
    item = tree.createElementNS(CEX, 'm:extLst');item.setAttribute('xmlns:m', CEX);tree.documentElement.appendChild(item)
    doc._package._dirty.add('word/commentsExtensible.xml')
    root = E.fromstring(DocxDocument(BytesIO(doc.to_bytes())).part_xml('word/commentsExtensible.xml'))
    assert [child.tag for child in root] == ['{' + CEX + '}commentExtensible', '{' + CEX + '}extLst']


@pytest.mark.parametrize('value', ['00000000', '80000000', 'FFFFFFFF'])
def test_invalid_paragraph_id_does_not_bind_modern_utc(tmp_path, value):
    doc = modern(tmp_path);entry = doc._package.tree('word/comments.xml').documentElement.getElementsByTagNameNS(W, 'p')[0]
    entry.setAttributeNS(W14, 'w14:paraId', value);entry.setAttribute('xmlns:w14', W14)
    ids = doc._package.tree('word/commentsIds.xml');ids.documentElement.childNodes[0].setAttributeNS(CID, 'm:paraId', value);ids.documentElement.childNodes[0].setAttribute('xmlns:m', CID)
    node = comment(doc)
    assert node.date_time_utc.year == 2026
    output = tmp_path / 'result.docx';doc.save(output)
    assert comment(DocxDocument(output)).date_time_utc.year == 2026


@pytest.mark.parametrize('value', ['80000000', 'FFFFFFFF'])
def test_large_durable_id_keeps_utc_and_source_identifier(tmp_path, value):
    doc = modern(tmp_path)
    for role, ns in [('commentsIds', CID), ('commentsExtensible', CEX)]:
        tree = doc._package.tree('word/' + role + '.xml');tree.documentElement.childNodes[0].setAttributeNS(ns, 'm:durableId', value)
    assert comment(doc).date_time_utc.year == 2030
    output = tmp_path / 'result.docx';doc.save(output);cold = DocxDocument(output)
    assert comment(cold).date_time_utc.year == 2030
    assert E.fromstring(cold.part_xml('word/commentsIds.xml'))[0].get('{' + CID + '}durableId') == value


def test_prepared_metadata_additions_reject_unsafe_part_names(tmp_path):
    doc = modern(tmp_path)
    with pytest.raises(ValueError, match='Unsafe DOCX part'):
        doc._package.to_bytes({'../owned.xml': b'<owned/>'})


def test_short_durable_id_matches_equivalent_eight_digit_value(tmp_path):
    doc = modern(tmp_path)
    ids = doc._package.tree('word/commentsIds.xml')
    items = ids.getElementsByTagNameNS(CID, 'commentId')
    items[0].setAttributeNS(CID, 'm:durableId', '2345678')
    ext = doc._package.tree('word/commentsExtensible.xml')
    ext.getElementsByTagNameNS(CEX, 'commentExtensible')[0].setAttributeNS(CEX, 'm:durableId', '02345678')
    node = comment(doc)
    assert node.date_time_utc == datetime(2030, 1, 2, 3, 4, 5, tzinfo=UTC)
    for flat in [False, True]:
        output = tmp_path / ('short-id-' + str(flat))
        (doc.save_flat_opc if flat else doc.save)(output)
        assert comment(DocxDocument(output)).date_time_utc == node.date_time_utc


@pytest.mark.parametrize('kind', ['orphan', 'duplicate-cid', 'duplicate-cex'])
@pytest.mark.parametrize('flat', [False, True])
def test_unassociated_records_survive_two_saves_and_date_edit(tmp_path, kind, flat):
    doc = modern(tmp_path)
    ids = doc._package.tree('word/commentsIds.xml').documentElement
    ext = doc._package.tree('word/commentsExtensible.xml').documentElement
    ns, local = (CEX, 'commentExtensible') if kind == 'duplicate-cex' else (CID, 'commentId')
    parent = ext if ns == CEX else ids
    node = parent.ownerDocument.createElementNS(ns, 'm:' + local)
    node.setAttribute('xmlns:m', ns)
    if ns == CID:
        node.setAttributeNS(CID, 'm:paraId', '00000001' if kind == 'orphan' else '12345678')
        node.setAttributeNS(CID, 'm:durableId', '00000001')
    else:
        node.setAttributeNS(CEX, 'm:durableId', '23456789')
        node.setAttributeNS(CEX, 'm:dateUtc', '2044-01-02T03:04:05Z')
    node.setAttribute('owned', kind);parent.appendChild(node)
    if kind != 'duplicate-cex':
        extra = ext.ownerDocument.createElementNS(CEX, 'm:commentExtensible')
        extra.setAttribute('xmlns:m', CEX)
        extra.setAttributeNS(CEX, 'm:durableId', '00000001')
        extra.setAttributeNS(CEX, 'm:dateUtc', '2044-01-02T03:04:05Z')
        extra.setAttribute('owned', 'extra');ext.appendChild(extra)
    # Keep the orphan and the non-winning duplicate records, including their attributes.
    unselected = ([node, extra] if kind == 'orphan' else
                  ([ids.getElementsByTagNameNS(CID, 'commentId')[0], ext.getElementsByTagNameNS(CEX, 'commentExtensible')[0]]
                   if kind == 'duplicate-cid' else [ext.getElementsByTagNameNS(CEX, 'commentExtensible')[0]]))
    expected_raw = [n.toxml() for n in unselected]
    held = comment(doc)
    assert held.date_time_utc.year == (2030 if kind == 'orphan' else 2044)
    if kind == 'orphan':
        copied = held.clone(True);doc.body.paragraphs[1].append_child(copied)
        assert copied.date_time_utc == held.date_time_utc
    output = tmp_path / 'record-output'
    for edit in [False, False, True]:
        if edit:held.date_time = datetime(2040, 1, 2, tzinfo=UTC)
        (doc.save_flat_opc if flat else doc.save)(output)
        cold = DocxDocument(output)
        records = [n.toxml() for p in ['word/commentsIds.xml', 'word/commentsExtensible.xml']
                   for n in cold._package.tree(p).documentElement.childNodes if n.nodeType == n.ELEMENT_NODE]
        assert all(raw in records for raw in expected_raw)
        assert comment(cold).date_time_utc == held.date_time_utc.replace(microsecond=0)
        if kind == 'orphan':
            assert cold.get_child_nodes(NodeType.COMMENT, True)[1].date_time_utc == copied.date_time_utc
    if kind == 'orphan':
        copied.remove();(doc.save_flat_opc if flat else doc.save)(output)
        cold = DocxDocument(output);assert len(cold.get_child_nodes(NodeType.COMMENT, True)) == 1
        records = [n.toxml() for p in ['word/commentsIds.xml', 'word/commentsExtensible.xml']
                   for n in cold._package.tree(p).documentElement.childNodes if n.nodeType == n.ELEMENT_NODE]
        assert len(cold._package.tree('word/commentsIds.xml').getElementsByTagNameNS(CID, 'commentId')) == 2
        assert all(raw in records for raw in expected_raw)


@pytest.mark.parametrize('flat', [False, True])
def test_modern_projection_failure_preserves_unassociated_records(tmp_path, monkeypatch, flat):
    doc = modern(tmp_path);root = doc._package.tree('word/commentsExtensible.xml').documentElement
    extra = root.getElementsByTagNameNS(CEX, 'commentExtensible')[0].cloneNode(True)
    extra.setAttribute('xmlns:m', CEX);extra.setAttributeNS(CEX, 'm:durableId', '00000001');extra.setAttribute('owned', 'orphan');root.appendChild(extra)
    held = comment(doc);snapshot = root.toxml();names = doc._package.part_names;value = held.date_time_utc
    output = tmp_path / 'atomic-records';output.write_bytes(b'previous')
    def fail(*args):raise OSError('owned failure')
    monkeypatch.setattr('os.replace', fail)
    with pytest.raises(OSError, match='owned failure'):(doc.save_flat_opc if flat else doc.save)(output)
    assert output.read_bytes() == b'previous' and root.toxml() == snapshot
    assert doc._package.part_names == names and held.date_time_utc == value


@pytest.mark.parametrize('reverse_ids', [False, True])
@pytest.mark.parametrize('reverse_paragraphs', [False, True])
@pytest.mark.parametrize('flat', [False, True])
def test_multiple_paragraphs_select_last_cid_record_and_keep_loser(tmp_path, reverse_ids, reverse_paragraphs, flat):
    doc = modern(tmp_path)
    entry = doc._package.tree('word/comments.xml').getElementsByTagNameNS(W, 'comment')[0]
    extra_p = entry.ownerDocument.createElementNS(W, 'w:p');extra_p.setAttribute('xmlns:w', W)
    extra_p.setAttribute('xmlns:w14', W14);extra_p.setAttributeNS(W14, 'w14:paraId', '34567890')
    entry.insertBefore(extra_p, entry.firstChild) if reverse_paragraphs else entry.appendChild(extra_p)
    ids = doc._package.tree('word/commentsIds.xml').documentElement
    second = ids.firstChild.cloneNode(True);second.setAttribute('xmlns:m', CID)
    second.setAttributeNS(CID, 'm:paraId', '34567890');second.setAttributeNS(CID, 'm:durableId', '45678901')
    ids.insertBefore(second, ids.firstChild) if reverse_ids else ids.appendChild(second)
    ext = doc._package.tree('word/commentsExtensible.xml').documentElement
    extra = ext.firstChild.cloneNode(True);extra.setAttribute('xmlns:m', CEX)
    extra.setAttributeNS(CEX, 'm:durableId', '45678901');extra.setAttributeNS(CEX, 'm:dateUtc', '2044-01-02T03:04:05Z');ext.appendChild(extra)
    from defusedxml.minidom import parseString
    for name in ['word/commentsIds.xml', 'word/commentsExtensible.xml']:
        doc._package._trees[name] = parseString(doc._package.tree(name).toxml(encoding='utf-8'), forbid_dtd=True)
    ids = doc._package.tree('word/commentsIds.xml').documentElement
    loser = ids.firstChild.toxml()
    held = comment(doc);assert held.date_time_utc.year == (2030 if reverse_ids else 2044)
    copy = held.clone(True);doc.body.paragraphs[1].append_child(copy)
    for _ in range(2):
        output = tmp_path / 'multi';(doc.save_flat_opc if flat else doc.save)(output)
        cold = DocxDocument(output);nodes = cold.get_child_nodes(NodeType.COMMENT, True)
        assert [n.date_time_utc for n in nodes] == [held.date_time_utc, copy.date_time_utc]
        identifiers = [p.getAttributeNS(W14, 'paraId') for p in cold._package.tree('word/comments.xml').getElementsByTagNameNS(W, 'p')]
        assert len(identifiers) == len(set(identifiers))
        assert loser in [n.toxml() for n in cold._package.tree('word/commentsIds.xml').documentElement.childNodes]
