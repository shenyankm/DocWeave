"""Settings role recovery leaves generic resource validation and publication atomic."""
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as E

import pytest
import dockiln as dk
from dockiln._flat_opc import encode, decode
from dockiln._opc import ordered_relationships
from dockiln.dom import DocxDocument
from dockiln.writers.docx.constants import CT_SETTINGS, REL_SETTINGS, PKG_RELS_URI
from tests.test_settings_relationship_loading import settings_package, W, CT

REL = 'word/_rels/document.xml.rels'


def package(records, *, first=None, keep_first=False):
    with ZipFile(BytesIO(settings_package('internal'))) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    links = E.fromstring(parts[REL])
    for node in list(links):
        if node.get('Type') == REL_SETTINGS:
            links.remove(node)
    parts['word/first.xml'] = first if first is not None else (
        f'<w:settings xmlns:w="{W}"><w:defaultTabStop w:val="444"/>'
        '<w:compat><w:compatSetting w:name="compatibilityMode" '
        'w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat></w:settings>').encode()
    for record in records:
        E.SubElement(links, f'{{{PKG_RELS_URI}}}Relationship', record)
    if keep_first:
        E.SubElement(links, f'{{{PKG_RELS_URI}}}Relationship',
                     Id='owned-retained', Type='urn:owned:retained', Target='first.xml')
    parts[REL] = E.tostring(links)
    types = E.fromstring(parts['[Content_Types].xml'])
    E.SubElement(types, f'{{{CT}}}Override', PartName='/word/first.xml', ContentType=CT_SETTINGS)
    parts['[Content_Types].xml'] = E.tostring(types)
    out = BytesIO()
    with ZipFile(out, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return out.getvalue()


def link(identifier, target='first.xml', kind=REL_SETTINGS, **extra):
    return dict(Id=identifier, Type=kind, Target=target, **extra)


CASES = [
    ([link('rId2'), link('rId10', 'settings.xml')], 61.7),
    ([link('rId10', 'settings.xml'), link('rId2')], 61.7),
    ([link('rId1'), link('rId1', 'settings.xml')], 61.7),
    ([link('rId1', 'settings.xml'), link('rId1')], 22.2),
    ([link('rId1'), link('rId1', kind='urn:foreign/settings'), link('rId2', 'settings.xml')], 61.7),
    ([link('rId1', kind='urn:foreign/settings'), link('rId1'), link('rId2', 'settings.xml')], 22.2),
    ([link('rId1'), link('rId1', '', kind='urn:foreign/settings'), link('rId2', 'settings.xml')], 22.2),
    ([link(''), link('rId2', 'settings.xml')], None),
    ([link('', ''), link('rId2', 'settings.xml')], 61.7),
    ([dict(Type=REL_SETTINGS), link('rId2', 'settings.xml')], 61.7),
    ([link('rId𐀀'), link('rId\ue000', 'settings.xml')], 22.2),
    ([link('rId\ue000'), link('rId𐀀', 'settings.xml')], 61.7),
    ([link('rIdÅ'), link('rIdz', 'settings.xml')], 61.7),
    ([link('rId1', TargetMode='External'), link('rId2', 'settings.xml')], 22.2),
    ([link('rId1', TargetMode='OwnedMode'), link('rId2', 'settings.xml')], 22.2),
    ([link('rId1', 'https://example.invalid/settings.xml', TargetMode='External'), link('rId2', 'settings.xml')], 36),
    ([link('rId1', 'missing.xml'), link('rId2', 'settings.xml')], 36),
]


@pytest.mark.parametrize('records,expected', CASES)
@pytest.mark.parametrize('flat', [False, True])
@pytest.mark.parametrize('model', ['conversion', 'dom'])
def test_selection_and_twice_saved_cold(tmp_path, records, expected, flat, model):
    raw = package(records)
    raw = encode(raw) if flat else raw
    factory = dk.Document if model == 'conversion' else DocxDocument
    if expected is None:
        with pytest.raises(ValueError, match='identifier'):
            factory(BytesIO(raw))
        return
    doc = factory(BytesIO(raw))
    assert doc.default_tab_stop == expected
    if model == 'conversion':
        assert doc.light_document_model.compatibility_mode == (14 if expected == 22.2 else 15 if expected == 61.7 else 12)
    for number in [1, 2]:
        if number == 2:
            doc.default_tab_stop = 42.1
        value = expected if number == 1 else 42.1
        destination = tmp_path / f'owned-{number}.bin'
        if model == 'conversion':
            doc.save(destination, dk.SaveFormat.FLAT_OPC if flat else dk.SaveFormat.DOCX)
        elif flat:
            doc.save_flat_opc(destination)
        else:
            doc.save(destination)
        assert doc.default_tab_stop == value
        assert dk.Document(destination).default_tab_stop == value
        assert DocxDocument(destination).default_tab_stop == value
        with ZipFile(BytesIO(decode(destination.read_bytes()) if flat else destination.read_bytes())) as archive:
            links = E.fromstring(archive.read(REL))
            typed = [node for node in links if node.get('Type') == REL_SETTINGS]
            assert len(typed) == 1
            assert not typed[0].get('TargetMode')
            target = 'word/' + typed[0].get('Target')
            assert E.fromstring(archive.read(target)).tag == f'{{{W}}}settings'


def test_primitive_preserves_records_and_orders_before_type_filter():
    records = [link('rId𐀀'), link('rId\ue000'), link('rId𐀀', kind='urn:owned'), link('', '')]
    snapshot = [dict(record) for record in records]
    result = ordered_relationships(records, dict.get)
    assert result[0] is records[2] and result[1] is records[1]
    assert records == snapshot


@pytest.mark.parametrize('flat', [False, True])
@pytest.mark.parametrize('failure', ['budget', 'publication'])
def test_projection_keeps_live_parts_and_held_nodes_until_success(tmp_path, monkeypatch, flat, failure):
    doc = DocxDocument(BytesIO(package([link('rId1'), link('rId2', 'settings.xml')])))
    held = doc.body.paragraphs[0]
    before = {name: doc._package.payload(name) for name in doc.part_names}
    output = doc.to_flat_opc() if flat else doc.to_bytes()
    assert {name: doc._package.payload(name) for name in doc.part_names} == before
    assert held is doc.body.paragraphs[0]
    assert doc.default_tab_stop == 22.2
    with ZipFile(BytesIO(decode(output) if flat else output)) as archive:
        assert 'word/settings.xml' not in archive.namelist()
        assert 'word/first.xml' in archive.namelist()
    destination = tmp_path / 'owned.bin'
    destination.write_bytes(b'owned previous output')
    with monkeypatch.context() as patch:
        if failure == 'budget':
            from dockiln import _io
            patch.setattr(_io, 'MAX_EXPANDED_BYTES', 1)
        else:
            import os
            def fail(*args):
                raise OSError('owned publication failure')
            patch.setattr(os, 'replace', fail)
        with pytest.raises((ValueError, OSError)):
            doc.save_flat_opc(destination) if flat else doc.save(destination)
    assert destination.read_bytes() == b'owned previous output'
    assert {name: doc._package.payload(name) for name in doc.part_names} == before
    assert held is doc.body.paragraphs[0] and doc.default_tab_stop == 22.2
    doc.save_flat_opc(destination) if flat else doc.save(destination)
    assert held is doc.body.paragraphs[0] and doc.default_tab_stop == 22.2
    assert 'word/settings.xml' not in doc.part_names


@pytest.mark.parametrize('first', [b'<broken', b'<!DOCTYPE settings [<!ENTITY secret "owned">]><settings/>'])
@pytest.mark.parametrize('model', [dk.Document, DocxDocument])
def test_selected_malformed_settings_fail_at_construction(first, model):
    from xml.parsers.expat import ExpatError
    from defusedxml.common import DefusedXmlException
    with pytest.raises((E.ParseError, ExpatError, DefusedXmlException)):
        model(BytesIO(package([link('rId1')], first=first)))


@pytest.mark.parametrize('flat', [False, True])
def test_foreign_root_and_other_referenced_settings_stay_preserved(tmp_path, flat):
    raw = package([link('rId1')], first=b'<owned xmlns="urn:owned">foreign</owned>')
    doc = DocxDocument(BytesIO(raw))
    assert doc.default_tab_stop == 36
    result = doc.to_flat_opc() if flat else doc.to_bytes()
    with ZipFile(BytesIO(decode(result) if flat else result)) as archive:
        data = archive.read('word/first.xml')
        if not flat:
            assert data == b'<owned xmlns="urn:owned">foreign</owned>'
        root = E.fromstring(data)
        assert root.tag == '{urn:owned}owned' and root.text == 'foreign'
    doc = DocxDocument(BytesIO(package([link('rId1', 'settings.xml')], keep_first=True)))
    result = doc.to_bytes()
    with ZipFile(BytesIO(result)) as archive:
        assert 'word/first.xml' in archive.namelist()


@pytest.mark.parametrize('kind', ['image', 'hyperlink', 'styles', 'numbering', 'theme',
    'header', 'footer', 'fontTable', 'stylesWithEffects', 'footnotes', 'endnotes', 'comments'])
@pytest.mark.parametrize('invalid', ['missing', 'duplicate', 'settings-collision'])
def test_generic_main_resource_relationship_ids_remain_strict(kind, invalid):
    resource = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + kind
    first = link('owned', 'https://example.invalid/owned', kind=resource, TargetMode='External')
    records = [first]
    if invalid == 'missing':
        first.pop('Id')
    elif invalid == 'duplicate':
        records.append(dict(first))
    else:
        records.append(link('owned'))
    with pytest.raises(ValueError, match='relationship (ID|identifier)'):
        dk.Document(BytesIO(package(records)))


@pytest.mark.parametrize('conflict', ['wrong', 'duplicate'])
def test_projection_content_type_failure_preserves_output_and_source(tmp_path, conflict):
    raw = package([link('rId1', 'missing.xml')])
    with ZipFile(BytesIO(raw)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    types = E.fromstring(parts['[Content_Types].xml'])
    for _ in range(2 if conflict == 'duplicate' else 1):
        E.SubElement(types, f'{{{CT}}}Override', PartName='/word/settings2.xml',
                     ContentType=CT_SETTINGS if conflict == 'duplicate' else 'application/octet-stream')
    parts['[Content_Types].xml'] = E.tostring(types)
    stream = BytesIO()
    with ZipFile(stream, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    doc = DocxDocument(BytesIO(stream.getvalue()))
    snapshot = {name: doc._package.payload(name) for name in doc.part_names}
    destination = tmp_path / 'owned.bin'
    destination.write_bytes(b'owned previous output')
    with pytest.raises(ValueError, match='Conflicting settings content type'):
        doc.save(destination)
    assert destination.read_bytes() == b'owned previous output'
    assert {name: doc._package.payload(name) for name in doc.part_names} == snapshot
    assert doc.default_tab_stop == 36


def test_normalization_diagnostic_and_report_do_not_include_source_data():
    from dockiln.diagnostics import collect_diagnostics
    doc = DocxDocument(BytesIO(package([link('owned-private-identifier'), link('z', 'settings.xml')])))
    diagnostics = []
    with collect_diagnostics(diagnostics):
        report = doc.preservation_report()
    assert len(diagnostics) == 1 and diagnostics[0].code == 'settings.relationships_normalized'
    assert 'owned-private-identifier' not in diagnostics[0].message
    assert diagnostics[0].location is None
    assert 'word/settings.xml' in report['removed']


@pytest.mark.parametrize('role', ['commentsIds', 'commentsExtensible'])
@pytest.mark.parametrize('last', ['settings', 'metadata'])
@pytest.mark.parametrize('flat', [False, True])
def test_settings_choice_is_frozen_before_other_role_projection(tmp_path, monkeypatch, role, last, flat):
    from tests.test_dom_comment_utc import modern
    from tests.test_dom_comment_metadata import comment
    from dockiln.dom.comment_utc import ROLES
    source = modern(tmp_path)
    parts = {name: source._package.payload(name) for name in source.part_names}
    links = E.fromstring(parts[REL])
    metadata = next(node for node in links if node.get('Type') == ROLES[role][2])
    links.remove(metadata)
    metadata.set('Id', 'owned-collision')
    settings = E.Element(f'{{{PKG_RELS_URI}}}Relationship', link('owned-collision', 'settings.xml'))
    links.extend([metadata, settings] if last == 'settings' else [settings, metadata])
    parts[REL] = E.tostring(links)
    parts['word/settings.xml'] = f'<w:settings xmlns:w="{W}"><w:defaultTabStop w:val="444"/></w:settings>'.encode()
    types = E.fromstring(parts['[Content_Types].xml'])
    E.SubElement(types, f'{{{CT}}}Override', PartName='/word/settings.xml', ContentType=CT_SETTINGS)
    parts['[Content_Types].xml'] = E.tostring(types)
    stream = BytesIO()
    with ZipFile(stream, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    doc = DocxDocument(BytesIO(encode(stream.getvalue()) if flat else stream.getvalue()))
    held = comment(doc)
    tabs, utc = doc.default_tab_stop, held.date_time_utc
    assert tabs == (22.2 if last == 'settings' else 36)
    assert utc.year == (2026 if last == 'settings' else 2030)
    original = {name: doc._package.payload(name) for name in doc.part_names}
    pure = doc.to_flat_opc() if flat else doc.to_bytes()
    assert DocxDocument(BytesIO(pure)).default_tab_stop == tabs
    assert {name: doc._package.payload(name) for name in doc.part_names} == original
    import os
    def fail(*args):
        raise OSError('owned publication failure')
    destination = tmp_path / 'owned.bin'
    destination.write_bytes(b'owned previous output')
    with monkeypatch.context() as patch:
        patch.setattr(os, 'replace', fail)
        with pytest.raises(OSError):
            doc.save_flat_opc(destination) if flat else doc.save(destination)
    assert destination.read_bytes() == b'owned previous output'
    assert {name: doc._package.payload(name) for name in doc.part_names} == original
    assert held is comment(doc) and held.date_time_utc == utc and doc.default_tab_stop == tabs
    for _ in range(2):
        doc.save_flat_opc(destination) if flat else doc.save(destination)
        assert held is comment(doc) and held.date_time_utc == utc and doc.default_tab_stop == tabs
        cold = DocxDocument(destination)
        assert cold.default_tab_stop == tabs and comment(cold).date_time_utc == utc
    doc.default_tab_stop = 42.1
    doc.save_flat_opc(destination) if flat else doc.save(destination)
    cold = DocxDocument(destination)
    assert doc.default_tab_stop == cold.default_tab_stop == 42.1
    assert held is comment(doc) and held.date_time_utc == comment(cold).date_time_utc == utc


@pytest.mark.parametrize('opaque', [False, True])
@pytest.mark.parametrize('flat', [False, True])
def test_removed_settings_owner_has_no_dangling_relationship_part(opaque, flat):
    with ZipFile(BytesIO(package([link('rId2'), link('rId1', 'settings.xml')]))) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    related = 'word/_rels/first.xml.rels'
    parts[related] = (b'<owned xmlns="urn:owned">opaque</owned>' if opaque else
        f'<Relationships xmlns="{PKG_RELS_URI}"><Relationship Id="owned" '
        'Type="urn:owned:resource" Target="owned.bin"/></Relationships>'.encode())
    parts['word/owned.bin'] = b'Owned harmless data'
    types = E.fromstring(parts['[Content_Types].xml'])
    E.SubElement(types, f'{{{CT}}}Override', PartName='/word/owned.bin', ContentType='application/octet-stream')
    if opaque:
        E.SubElement(types, f'{{{CT}}}Override', PartName='/' + related, ContentType='application/octet-stream')
    parts['[Content_Types].xml'] = E.tostring(types)
    raw = BytesIO()
    with ZipFile(raw, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    doc = DocxDocument(BytesIO(raw.getvalue()))
    snapshot = {name: doc._package.payload(name) for name in doc.part_names}
    output = doc.to_flat_opc() if flat else doc.to_bytes()
    assert {name: doc._package.payload(name) for name in doc.part_names} == snapshot
    with ZipFile(BytesIO(decode(output) if flat else output)) as archive:
        assert ('word/first.xml' in archive.namelist()) == opaque
        assert (related in archive.namelist()) == opaque
        assert archive.read('word/owned.bin') == b'Owned harmless data'
        if opaque:
            assert archive.read(related) == b'<owned xmlns="urn:owned">opaque</owned>'
    assert DocxDocument(BytesIO(output)).default_tab_stop == 61.7


@pytest.mark.parametrize('target', ['https://example.invalid/owned.xml',
    'http://example.invalid/owned.xml', 'file:///word/first.xml',
    'mailto:owned@example.invalid', '//example.invalid/owned.xml',
    'first.xml#owned', 'first.xml?owned=1', '../../owned.xml'])
@pytest.mark.parametrize('mode', ['', 'External', 'OwnedMode'])
@pytest.mark.parametrize('model', ['conversion', 'dom'])
@pytest.mark.parametrize('flat', [False, True])
def test_unresolved_settings_uri_is_ignored_without_io(tmp_path, monkeypatch, target, mode, model, flat):
    import socket
    def network(*args, **kwargs):
        raise AssertionError('Settings must not open a network connection')
    monkeypatch.setattr(socket, 'create_connection', network)
    raw = package([link('a', target, **({'TargetMode': mode} if mode else {})), link('z', 'settings.xml')])
    factory = dk.Document if model == 'conversion' else DocxDocument
    doc = factory(BytesIO(encode(raw) if flat else raw))
    assert doc.default_tab_stop == 36
    output = tmp_path / 'owned.bin'
    for value in [36, 42.1]:
        doc.default_tab_stop = value
        if model == 'conversion':
            doc.save(output, dk.SaveFormat.FLAT_OPC if flat else dk.SaveFormat.DOCX)
        elif flat:
            doc.save_flat_opc(output)
        else:
            doc.save(output)
        assert doc.default_tab_stop == factory(output).default_tab_stop == value
    # Generic package resolution still rejects remote/absolute URIs.
    from dockiln._opc import resolve_target
    with pytest.raises(ValueError):
        resolve_target('word/document.xml', target)
