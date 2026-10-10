"""Theme dependency capture preserves binary data and rejects broken graphs."""

import runpy
import xml.etree.ElementTree as ET
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from defusedxml.common import DefusedXmlException

from aspose.words_foss import Document
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._opc import related_part_snapshot

CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
THEME = 'custom/themes/owned.xml'
BINARY = 'custom/assets/owned.bin'


def relationships(*items):
    root = ET.Element(f'{{{RELS}}}Relationships')
    for item in items:
        ET.SubElement(root, f'{{{RELS}}}Relationship', item)
    return ET.tostring(root)


def edge(target, rid='resource', mode=None):
    result = {'Id': rid, 'Type': 'urn:owned:resource', 'Target': target}
    if mode is not None:
        result['TargetMode'] = mode
    return result


def package_parts():
    corpus = Path(__file__).parents[1] / 'docs/benchmarks/corpus/font-theme-corpus-26.9.zip'
    with ZipFile(corpus) as outer, ZipFile(BytesIO(outer.read(outer.namelist()[0]))) as inner:
        parts = {name: inner.read(name) for name in inner.namelist()}
    parts[THEME] = parts.pop('word/theme/theme1.xml')
    types = ET.fromstring(parts['[Content_Types].xml'])
    for item in types:
        if item.get('PartName') == '/word/theme/theme1.xml':
            item.set('PartName', '/' + THEME)
    ET.SubElement(types, f'{{{CT}}}Override', {'PartName': '/' + BINARY, 'ContentType': 'application/octet-stream'})
    parts['[Content_Types].xml'] = ET.tostring(types)
    rels = ET.fromstring(parts['word/_rels/document.xml.rels'])
    for item in rels:
        if item.get('Type', '').endswith('/theme'):
            item.set('Target', '../' + THEME)
    parts['word/_rels/document.xml.rels'] = ET.tostring(rels)
    parts[BINARY] = b'\x00\xffowned binary\x80'
    parts['custom/themes/_rels/owned.xml.rels'] = relationships(
        edge('../assets/owned.bin'), edge('../assets/owned.bin', 'shared'),
        edge('https://invalid.example/never-fetch', 'external', 'External'))
    parts['custom/assets/_rels/owned.bin.rels'] = relationships(edge('../themes/owned.xml', 'cycle'))
    return parts


def archive_bytes(parts):
    stream = BytesIO()
    with ZipFile(stream, 'w', compression=ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return stream.getvalue()


def test_theme_graph_survives_loading_and_json_without_copying_unrelated_parts():
    parts = package_parts()
    model = Document(BytesIO(archive_bytes(parts))).light_document_model
    snapshot = model.source_theme
    assert snapshot.part_name == THEME
    captured = {part.name: part for part in snapshot.parts}
    assert set(captured) == {THEME, BINARY, 'custom/themes/_rels/owned.xml.rels',
                             'custom/assets/_rels/owned.bin.rels'}
    assert all(part.data == parts[name] for name, part in captured.items())
    assert captured[BINARY].content_type == 'application/octet-stream'
    assert ldm.Document.model_validate_json(model.model_dump_json()).source_theme == snapshot
    with pytest.raises(ValueError, match='frozen'):
        captured[BINARY].data = b'changed'


@pytest.mark.parametrize('kind', ['missing', 'escape', 'duplicate', 'empty', 'mode', 'wrong-root',
                                 'infrastructure', 'missing-type', 'duplicate-type', 'dtd'])
def test_invalid_dependency_graph_is_rejected(kind):
    parts = package_parts()
    rels_name = 'custom/themes/_rels/owned.xml.rels'
    if kind == 'missing':
        del parts[BINARY]
    elif kind == 'escape':
        parts[rels_name] = relationships(edge('../../../outside.bin'))
    elif kind == 'duplicate':
        parts[rels_name] = relationships(edge('../assets/owned.bin'), edge('../assets/owned.bin'))
    elif kind == 'empty':
        parts[rels_name] = relationships(edge(''))
    elif kind == 'mode':
        parts[rels_name] = relationships(edge('../assets/owned.bin', mode='Other'))
    elif kind == 'wrong-root':
        parts[rels_name] = b'<Relationships/>'
    elif kind == 'infrastructure':
        parts[rels_name] = relationships(edge('/[Content_Types].xml'))
    elif kind in {'missing-type', 'duplicate-type'}:
        root = ET.fromstring(parts['[Content_Types].xml'])
        item = next(item for item in root if item.get('PartName') == '/' + BINARY)
        if kind == 'missing-type':
            root.remove(item)
        else:
            ET.SubElement(root, item.tag, item.attrib)
        parts['[Content_Types].xml'] = ET.tostring(root)
    else:
        parts[rels_name] = b'<!DOCTYPE Relationships [<!ENTITY x "secret">]><Relationships/>'
    with ZipFile(BytesIO(archive_bytes(parts))) as archive, pytest.raises((ValueError, DefusedXmlException)):
        related_part_snapshot(archive, THEME)


@pytest.mark.parametrize('format', ['docx', 'flat_opc'])
@pytest.mark.parametrize('json_roundtrip', [False, True])
def test_graph_is_rebound_and_retained_through_saving(format, json_roundtrip):
    from aspose.words_foss import SaveFormat
    from aspose.words_foss._opc import resolve_target
    from aspose.words_foss.docx_writer import LdmDocxWriter
    from aspose.words_foss.saving import OoxmlSaveOptions
    model = Document(BytesIO(archive_bytes(package_parts()))).light_document_model
    if json_roundtrip:
        model = ldm.Document.model_validate_json(model.model_dump_json())
    before = model.source_theme
    options = OoxmlSaveOptions(SaveFormat.DOCX if format == 'docx' else SaveFormat.FLAT_OPC)
    saved = LdmDocxWriter(options).write_to_bytes(model)
    reread = Document(BytesIO(saved)).light_document_model.source_theme
    assert model.source_theme == before
    assert reread.part_name == 'word/theme/theme1.xml'
    captured = {part.name: part for part in reread.parts}
    binary = next(part for part in reread.parts if part.content_type == 'application/octet-stream')
    assert binary.data == package_parts()[BINARY]
    assert len(captured) == 4
    for part in reread.parts:
        if part.name.endswith('.rels'):
            parent = part.name.replace('/_rels/', '/').removesuffix('.rels')
            for item in ET.fromstring(part.data):
                if item.get('TargetMode') == 'External':
                    assert item.get('Target') == 'https://invalid.example/never-fetch'
                else:
                    assert resolve_target(parent, item.get('Target')) in captured


@pytest.mark.parametrize('damage', ['missing', 'duplicate', 'dangling-id', 'escape'])
def test_damaged_snapshot_does_not_replace_existing_output(tmp_path, damage):
    from aspose.words_foss.docx_writer import LdmDocxWriter
    model = Document(BytesIO(archive_bytes(package_parts()))).light_document_model
    theme = model.source_theme
    parts = list(theme.parts)
    if damage == 'missing':
        parts = [part for part in parts if part.name != BINARY]
    elif damage == 'duplicate':
        parts.append(parts[0])
    elif damage == 'escape':
        parts.append(ldm.SourceThemePart(name='../escape.bin', content_type='application/octet-stream', data=b'owned'))
    else:
        root = ET.fromstring(theme.data)
        root.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', 'missing')
        data = ET.tostring(root)
        parts = [part.model_copy(update={'data': data}) if part.name == theme.part_name else part for part in parts]
        theme = theme.model_copy(update={'data': data})
    model.source_theme = theme.model_copy(update={'parts': tuple(parts)})
    output = tmp_path / 'existing.docx'
    output.write_bytes(b'existing')
    with pytest.raises(ValueError):
        LdmDocxWriter().write(model, output)
    assert output.read_bytes() == b'existing'


def test_resource_directory_collision_uses_an_unused_name():
    from aspose.words_foss.docx_writer.package import _import_theme_resources
    model = Document(BytesIO(archive_bytes(package_parts()))).light_document_model
    text = [('[Content_Types].xml', f'<Types xmlns="{CT}"/>')]
    binary = [('word/theme/resources1/custom/assets/owned.bin', b'existing')]
    _import_theme_resources(model.source_theme, text, binary)
    assert binary[0][1] == b'existing'
    assert ('word/theme/resources2/' + BINARY, package_parts()[BINARY]) in binary
    assert len({name.casefold() for name, _ in binary}) == len(binary)


def test_frozen_native_cold_reads_retain_the_reachable_theme_image():
    root = Path(__file__).parents[1]
    check = runpy.run_path(str(root / "scripts/verify_font_themes.py"))["verify_theme_resources"]
    assert check(root / "docs/benchmarks") == 6


@pytest.mark.parametrize('part', ['custom/themes/_rels/owned.xml.rels', '[Content_Types].xml'])
def test_dependency_parse_errors_do_not_echo_private_entity_names(part):
    parts = package_parts()
    parts[part] = b'<owned>&PRIVATE_ACCOUNT_123;</owned>'
    with ZipFile(BytesIO(archive_bytes(parts))) as archive, pytest.raises(ValueError) as error:
        related_part_snapshot(archive, THEME)
    assert str(error.value) == 'Invalid related OPC XML'
    assert error.value.__suppress_context__


@pytest.mark.parametrize('filename', ['space name.bin', 'literal%20.bin', '颜色.bin'])
@pytest.mark.parametrize('format', ['docx', 'flat_opc'])
def test_encoded_resource_names_preserve_literal_names_and_content_types(filename, format):
    from urllib.parse import quote

    from aspose.words_foss import SaveFormat
    from aspose.words_foss._opc import relationships_path
    from aspose.words_foss.docx_writer import LdmDocxWriter
    from aspose.words_foss.saving import OoxmlSaveOptions

    parts = package_parts()
    name = 'custom/assets/' + filename
    parts[name] = parts.pop(BINARY)
    parts[relationships_path(name)] = parts.pop(relationships_path(BINARY))
    types = ET.fromstring(parts['[Content_Types].xml'])
    for item in types:
        if item.get('PartName') == '/' + BINARY:
            item.set('PartName', quote('/' + name, safe='/'))
    parts['[Content_Types].xml'] = ET.tostring(types)
    links_name = relationships_path(THEME)
    links = ET.fromstring(parts[links_name])
    for item in links:
        if item.get('TargetMode') != 'External':
            item.set('Target', quote('../assets/' + filename, safe='/'))
    parts[links_name] = ET.tostring(links)
    model = Document(BytesIO(archive_bytes(parts))).light_document_model
    model = ldm.Document.model_validate_json(model.model_dump_json())
    fmt = SaveFormat.DOCX if format == 'docx' else SaveFormat.FLAT_OPC
    raw = LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(model)
    reopened = Document(BytesIO(raw)).light_document_model.source_theme
    binary = [part for part in reopened.parts if part.content_type == 'application/octet-stream']
    assert len(binary) == 1 and binary[0].name.endswith('/' + filename)
    assert binary[0].data == parts[name]
