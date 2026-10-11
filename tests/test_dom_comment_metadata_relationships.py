"""Owned modern comment part selection and CID-only lifecycle controls."""
from datetime import datetime, timezone
from io import BytesIO
import pytest
from aspose.words_foss import NodeType
from aspose.words_foss.dom import DocxDocument
from aspose.words_foss.dom.comment_utc import CID, CEX, ROLES, elements
from .test_dom_comment_utc import modern, REL, CT
from .test_dom_comment_metadata import comment, source_date


def conflicting(tmp_path, role, first, second, *, reverse=False):
    doc = modern(tmp_path)
    rels = doc._package.tree('word/_rels/document.xml.rels').documentElement
    original = next(n for n in elements(rels, REL, 'Relationship')
                    if n.getAttribute('Type') == ROLES[role][2])
    original.setAttribute('Id', first)
    duplicate = original.cloneNode(True)
    duplicate.setAttribute('Id', second)
    duplicate.setAttribute('Target', 'owned-second-' + role + '.xml')
    rels.insertBefore(duplicate, original) if reverse else rels.appendChild(duplicate)
    part = doc._package.tree('word/' + role + '.xml').cloneNode(True)
    ns = CID if role == 'commentsIds' else CEX
    part.documentElement.setAttribute('xmlns:m', ns)
    node = elements(part.documentElement, ns, ROLES[role][1])[0]
    if role == 'commentsIds':
        node.setAttributeNS(CID, 'm:durableId', '45678901')
        ext = doc._package.tree('word/commentsExtensible.xml').documentElement
        other = elements(ext, CEX, 'commentExtensible')[0].cloneNode(True)
        other.setAttribute('xmlns:m', CEX)
        other.setAttributeNS(CEX, 'm:durableId', '45678901')
        other.setAttributeNS(CEX, 'm:dateUtc', '2044-01-02T03:04:05Z')
        ext.appendChild(other)
    else:node.setAttributeNS(CEX, 'm:dateUtc', '2044-01-02T03:04:05Z')
    from aspose.words_foss.utils.xml_helpers import serialize_xml
    doc._package._commit_parts({'word/owned-second-' + role + '.xml': serialize_xml(part)})
    types = doc._package.tree('[Content_Types].xml').documentElement
    override = types.ownerDocument.createElementNS(CT, 'Override')
    override.setAttribute('xmlns', CT)
    override.setAttribute('PartName', '/word/owned-second-' + role + '.xml')
    override.setAttribute('ContentType', 'application/vnd.ms-word.' + role + '+xml')
    types.appendChild(override)
    doc._package._dirty.update({'word/_rels/document.xml.rels', '[Content_Types].xml', 'word/commentsExtensible.xml'})
    return doc


@pytest.mark.parametrize('role', list(ROLES))
@pytest.mark.parametrize('first,second,year', [('rId2', 'rId10', 2044), ('rId10', 'rId2', 2030),
    ('Z', 'a', 2030), ('a', 'A', 2044), ('same', 'same', 2044),
    ('rId\U00010000', 'rId\ue000', 2030), ('rId\ue000', 'rId\U00010000', 2044),
    ('rIdÅ', 'rIdz', 2044), ('rIdz', 'rIdÅ', 2030)])
@pytest.mark.parametrize('flat', [False, True])
def test_modern_part_selection_follows_relation_id_order(tmp_path, role, first, second, year, flat):
    doc = conflicting(tmp_path, role, first, second)
    held = comment(doc)
    assert held.date_time_utc.year == year
    before = {n: doc._package.payload(n) for n in doc.part_names}
    data = doc.to_flat_opc() if flat else doc.to_bytes()
    assert comment(DocxDocument(BytesIO(data))).date_time_utc.year == year
    assert {n: doc._package.payload(n) for n in doc.part_names} == before
    out = tmp_path / 'result'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(out)
        assert comment(doc) is held and held.date_time_utc.year == year
        assert comment(DocxDocument(out)).date_time_utc.year == year


@pytest.mark.parametrize('operation', ['leave', 'edit', 'clone', 'import'])
@pytest.mark.parametrize('flat', [False, True])
def test_cid_only_does_not_create_extensible_part(tmp_path, operation, flat):
    donor = modern(tmp_path, omit='commentsExtensible')
    node = comment(donor)
    doc = donor
    if operation == 'edit':node.date_time = datetime(2040, 1, 2, tzinfo=timezone.utc)
    elif operation in {'clone', 'import'}:
        if operation == 'import':doc = source_date(tmp_path, None)
        node = doc.import_node(node, True) if operation == 'import' else node.clone(True)
        doc.body.paragraphs[-1].append_child(node)
    held = node.date_time_utc
    from aspose.words_foss.dom.comment_metadata import _local_wall
    expected_cold = _local_wall(held.replace(tzinfo=None)).astimezone(timezone.utc) if operation == 'edit' else held
    out = tmp_path / 'result'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(out)
        assert node.date_time_utc == held
        cold = DocxDocument(out)
        rels = cold._package.tree('word/_rels/document.xml.rels').documentElement
        roles = {n.getAttribute('Type') for n in elements(rels, REL, 'Relationship')}
        assert ROLES['commentsIds'][2] in roles
        assert ROLES['commentsExtensible'][2] not in roles
        assert all('commentsExtensible' not in n for n in cold.part_names)
        assert any(n.date_time_utc == expected_cold for n in cold.get_child_nodes(NodeType.COMMENT, True))


@pytest.mark.parametrize('flat', [False, True])
def test_losing_part_removal_is_atomic_and_visible(tmp_path, flat, monkeypatch):
    doc = conflicting(tmp_path, 'commentsExtensible', 'rId2', 'rId10')
    held = comment(doc)
    losing = 'word/commentsExtensible.xml'
    chosen = 'word/owned-second-commentsExtensible.xml'
    before = {n: doc._package.payload(n) for n in doc.part_names}
    held_value = held.date_time_utc
    assert losing in doc.preservation_report()['removed']
    data = doc.to_flat_opc() if flat else doc.to_bytes()
    cold = DocxDocument(BytesIO(data))
    assert losing not in cold.part_names and chosen in cold.part_names
    assert {n: doc._package.payload(n) for n in doc.part_names} == before
    output = tmp_path / 'prior'
    output.write_bytes(b'Owned prior output')
    import os
    def fail(*args):raise OSError('Owned publication failure')
    with monkeypatch.context() as context:
        context.setattr(os, 'replace', fail)
        with pytest.raises(OSError, match='publication'):
            (doc.save_flat_opc if flat else doc.save)(output)
    assert output.read_bytes() == b'Owned prior output'
    assert held.date_time_utc is held_value and comment(doc) is held
    assert {n: doc._package.payload(n) for n in doc.part_names} == before
    from aspose.words_foss import _io
    with monkeypatch.context() as context:
        context.setattr(_io, 'MAX_EXPANDED_BYTES', 1)
        with pytest.raises(ValueError, match='safety limit'):
            (doc.save_flat_opc if flat else doc.save)(output)
    assert output.read_bytes() == b'Owned prior output'
    assert held.date_time_utc is held_value
    assert {n: doc._package.payload(n) for n in doc.part_names} == before
    (doc.save_flat_opc if flat else doc.save)(output)
    assert losing not in doc.part_names and comment(doc) is held
    assert comment(DocxDocument(output)).date_time_utc == held_value


@pytest.mark.parametrize('flat', [False, True])
def test_losing_metadata_root_with_other_relation_is_preserved(tmp_path, flat):
    doc = conflicting(tmp_path, 'commentsExtensible', 'rId2', 'rId10')
    rels = doc._package.tree('word/_rels/document.xml.rels').documentElement
    link = rels.ownerDocument.createElementNS(REL, 'Relationship')
    link.setAttribute('xmlns', REL)
    for key, value in [('Id', 'retained'), ('Type', 'urn:owned:opaque'), ('Target', 'commentsExtensible.xml')]:
        link.setAttribute(key, value)
    rels.appendChild(link)
    data = doc.to_flat_opc() if flat else doc.to_bytes()
    cold = DocxDocument(BytesIO(data))
    assert 'word/commentsExtensible.xml' in cold.part_names
    assert comment(cold).date_time_utc.year == 2044


@pytest.mark.parametrize('empty', [True, False])
def test_empty_metadata_relationship_id_fails_at_load(tmp_path, empty):
    doc = modern(tmp_path)
    rels = doc._package.tree('word/_rels/document.xml.rels').documentElement
    link = next(n for n in elements(rels, REL, 'Relationship') if n.getAttribute('Type') == ROLES['commentsIds'][2])
    if empty:link.setAttribute('Id', '')
    else:link.removeAttribute('Id')
    doc._package._dirty.add('word/_rels/document.xml.rels')
    # Build the raw OPC bytes without applying the DOM serialization projection.
    with pytest.raises(ValueError, match='non-empty relationship identifier'):
        DocxDocument(BytesIO(doc._package.to_bytes()))


@pytest.mark.parametrize('role', list(ROLES))
@pytest.mark.parametrize('mode', ['external', 'missing', 'foreign-root', 'invalid-mode'])
@pytest.mark.parametrize('flat', [False, True])
def test_selected_invalid_target_does_not_fall_through(tmp_path, role, mode, flat):
    doc = modern(tmp_path)
    rels = doc._package.tree('word/_rels/document.xml.rels').documentElement
    original = next(n for n in elements(rels, REL, 'Relationship') if n.getAttribute('Type') == ROLES[role][2])
    link = original.cloneNode(True)
    link.setAttribute('Id', '000-selected')
    if mode == 'external':
        link.setAttribute('TargetMode', 'External')
        link.setAttribute('Target', 'https://owned.invalid/not-fetched')
    elif mode == 'missing':link.setAttribute('Target', 'owned-missing.xml')
    elif mode == 'invalid-mode':link.setAttribute('TargetMode', 'OwnedMode')
    else:
        link.setAttribute('Target', 'owned-opaque.xml')
        doc._package._commit_parts({'word/owned-opaque.xml': b'<owned-root>Owned opaque content</owned-root>'})
    rels.appendChild(link)
    doc._package._dirty.add('word/_rels/document.xml.rels')
    held = comment(doc)
    expected = held.date_time_utc
    assert expected.year == (2030 if mode == 'invalid-mode' else 2026)
    before = {n: doc._package.payload(n) for n in doc.part_names}
    out = tmp_path / 'result'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(out)
        assert held.date_time_utc == expected and comment(doc) is held
        cold = DocxDocument(out)
        assert comment(cold).date_time_utc == expected
        links = elements(cold._package.tree('word/_rels/document.xml.rels').documentElement, REL, 'Relationship')
        assert sum(n.getAttribute('Type') == ROLES[role][2] for n in links) <= 1
        if mode == 'foreign-root':
            if flat:
                from defusedxml.ElementTree import fromstring
                from xml.etree.ElementTree import tostring
                assert tostring(fromstring(cold._package.payload('word/owned-opaque.xml'))) == tostring(fromstring(before['word/owned-opaque.xml']))
            else:
                assert cold._package.payload('word/owned-opaque.xml') == before['word/owned-opaque.xml']


@pytest.mark.parametrize('role', list(ROLES))
@pytest.mark.parametrize('flat', [False, True])
def test_typed_relationship_reads_xml_despite_non_xml_content_type(tmp_path, role, flat):
    doc = modern(tmp_path)
    types = doc._package.tree('[Content_Types].xml').documentElement
    link = next(n for n in elements(types, CT, 'Override') if n.getAttribute('PartName') == '/word/' + role + '.xml')
    link.setAttribute('ContentType', 'application/octet-stream')
    doc._package._dirty.add('[Content_Types].xml')
    assert comment(doc).date_time_utc.year == 2030
    out = tmp_path / 'result'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(out)
        assert comment(DocxDocument(out)).date_time_utc.year == 2030


@pytest.mark.parametrize('flat', [False, True])
def test_selected_relationship_extensions_are_retained(tmp_path, flat):
    doc = conflicting(tmp_path, 'commentsExtensible', 'rId2', 'rId10')
    rels = doc._package.tree('word/_rels/document.xml.rels').documentElement
    selected = next(n for n in elements(rels, REL, 'Relationship') if n.getAttribute('Id') == 'rId10')
    selected.setAttribute('xmlns:owned', 'urn:owned:extension')
    selected.setAttributeNS('urn:owned:extension', 'owned:value', 'Owned retained attribute')
    extension = selected.ownerDocument.createElementNS('urn:owned:extension', 'owned:payload')
    selected.appendChild(extension)
    data = doc.to_flat_opc() if flat else doc.to_bytes()
    cold = DocxDocument(BytesIO(data))
    rels = cold._package.tree('word/_rels/document.xml.rels').documentElement
    selected = next(n for n in elements(rels, REL, 'Relationship') if n.getAttribute('Type') == ROLES['commentsExtensible'][2])
    assert selected.getAttributeNS('urn:owned:extension', 'value') == 'Owned retained attribute'
    assert len(elements(selected, 'urn:owned:extension', 'payload')) == 1
    assert comment(cold).date_time_utc.year == 2044


def test_public_package_updates_do_not_accept_removal_markers(tmp_path):
    doc = modern(tmp_path)
    before = {n: doc._package.payload(n) for n in doc.part_names}
    with pytest.raises(ValueError, match='Unsafe DOCX part'):
        doc._package.set_parts({'word/comments.xml': None})
    assert {n: doc._package.payload(n) for n in doc.part_names} == before


@pytest.mark.parametrize('role', list(ROLES))
@pytest.mark.parametrize('mode', ['empty', 'absent', 'empty-id', 'absent-id'])
@pytest.mark.parametrize('flat', [False, True])
def test_empty_target_is_skipped_before_identifier_validation(tmp_path, role, mode, flat):
    doc = modern(tmp_path)
    rels = doc._package.tree('word/_rels/document.xml.rels').documentElement
    original = next(n for n in elements(rels, REL, 'Relationship') if n.getAttribute('Type') == ROLES[role][2])
    ignored = original.cloneNode(True)
    ignored.setAttribute('Id', '000-ignored')
    ignored.setAttribute('Target', '')
    if mode == 'absent':ignored.removeAttribute('Target')
    if mode == 'empty-id':ignored.setAttribute('Id', '')
    if mode == 'absent-id':ignored.removeAttribute('Id')
    rels.appendChild(ignored)
    doc._package._dirty.add('word/_rels/document.xml.rels')
    assert comment(doc).date_time_utc.year == 2030
    out = tmp_path / 'result'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(out)
        assert comment(DocxDocument(out)).date_time_utc.year == 2030


@pytest.mark.parametrize('flat', [False, True])
def test_missing_cid_part_reuses_content_type_override(tmp_path, flat):
    doc = modern(tmp_path)
    doc._package._commit_parts({'word/commentsIds.xml': None})
    doc = DocxDocument(BytesIO(doc._package.to_bytes()))
    held = comment(doc)
    expected = held.date_time_utc
    original = {name: doc._package.payload(name) for name in doc.part_names}
    data = doc.to_flat_opc() if flat else doc.to_bytes()
    assert comment(DocxDocument(BytesIO(data))).date_time_utc == expected
    assert {name: doc._package.payload(name) for name in doc.part_names} == original
    out = tmp_path / 'recovered'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(out)
        cold = DocxDocument(out)
        types = cold._package.tree('[Content_Types].xml').documentElement
        names = [n.getAttribute('PartName') for n in elements(types, CT, 'Override')]
        assert names.count('/word/commentsIds.xml') == 1
        assert 'word/commentsIds.xml' in cold.part_names
        assert comment(cold).date_time_utc == expected
        assert comment(doc) is held and held.date_time_utc == expected


@pytest.mark.parametrize('flat', [False, True])
def test_available_foreign_root_keeps_extensible_role_without_overwriting_opaque(tmp_path, flat):
    doc = modern(tmp_path)
    opaque = b'<owned-root owned="untouched"/>'
    doc._package._commit_parts({'word/commentsExtensible.xml': opaque})
    doc = DocxDocument(BytesIO(doc._package.to_bytes()))
    held = comment(doc)
    expected = held.date_time_utc
    before = {name: doc._package.payload(name) for name in doc.part_names}
    out = tmp_path / 'recovered-foreign'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(out)
        cold = DocxDocument(out)
        assert comment(cold).date_time_utc == expected
        rels = cold._package.tree('word/_rels/document.xml.rels').documentElement
        links = [n for n in elements(rels, REL, 'Relationship') if n.getAttribute('Type') == ROLES['commentsExtensible'][2]]
        assert len(links) == 1
        assert links[0].getAttribute('Target') != 'commentsExtensible.xml'
        if flat:
            from xml.etree import ElementTree as E
            root = E.fromstring(cold._package.payload('word/commentsExtensible.xml'))
            assert root.tag == 'owned-root' and root.attrib == {'owned': 'untouched'}
        else:
            assert cold._package.payload('word/commentsExtensible.xml') == opaque
        assert comment(doc) is held and held.date_time_utc == expected
    assert before['word/commentsExtensible.xml'] == doc._package.payload('word/commentsExtensible.xml')
