"""Flat OPC loading, package preservation, and rejection boundaries."""

from io import BytesIO
from xml.parsers.expat import ExpatError
from zipfile import ZipFile

import pytest
from defusedxml.common import DTDForbidden
from defusedxml.ElementTree import fromstring
from defusedxml.minidom import parseString

import aspose.words_foss as aw
from aspose.words_foss import _io
from aspose.words_foss._flat_opc import MAIN_TYPES, decode, encode, is_flat_opc
from aspose.words_foss.docx_reader.document_reader import DocumentReader

PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"


def xml_shape(node):
    attributes = (tuple(sorted((item.namespaceURI or '', item.localName or item.name, item.value)
                               for item in node.attributes.values())) if node.attributes else ())
    return (node.nodeType, node.namespaceURI, node.localName or node.nodeName, attributes,
            getattr(node, 'data', None), tuple(xml_shape(child) for child in node.childNodes))


def package(parts):
    return (f'<pkg:package xmlns:pkg="{PKG}" xmlns:w="urn:word" xmlns:keep="urn:opaque">{parts}</pkg:package>').encode()


def part(payload, name="/word/document.xml"):
    return f'<pkg:part pkg:name="{name}" pkg:contentType="application/xml">{payload}</pkg:part>'


def test_decode_preserves_namespace_context_xml_siblings_and_binary_bytes():
    source = package(part('<pkg:xmlData><!--before--><?test keep?><w:document keep:flag="yes"><w:t>文字</w:t></w:document><!--after--></pkg:xmlData>')
                     + part('<pkg:binaryData> AAH/\n </pkg:binaryData>', "/word/media/image.bin"))
    with ZipFile(BytesIO(decode(source))) as archive:
        xml = archive.read("word/document.xml")
        root = fromstring(xml)
        assert root.tag == "{urn:word}document" and root.attrib["{urn:opaque}flag"] == "yes"
        assert root.find("{urn:word}t").text == "文字"
        assert b"<!--before-->" in xml and b"<?test keep?>" in xml and b"<!--after-->" in xml
        assert archive.read("word/media/image.bin") == b"\x00\x01\xff"
        types = fromstring(archive.read("[Content_Types].xml"))
        assert len(types) == 2


@pytest.mark.parametrize("payload", [
    '<pkg:xmlData><w:a/><w:b/></pkg:xmlData>',
    '<pkg:binaryData>invalid!</pkg:binaryData>',
    '<pkg:binaryData><w:a/></pkg:binaryData>',
    '<pkg:unknown/>',
    '<pkg:xmlData><w:a/></pkg:xmlData><pkg:binaryData>AA==</pkg:binaryData>',
])
def test_invalid_payload_rejected(payload):
    with pytest.raises(ValueError):
        decode(package(part(payload)))


@pytest.mark.parametrize("name", ["word/document.xml", "/../outside", "//word/document.xml", "/word\\outside", "/[Content_Types].xml"])
def test_unsafe_part_names_rejected(name):
    with pytest.raises(ValueError):
        decode(package(part('<pkg:xmlData><w:a/></pkg:xmlData>', name)))


def test_duplicate_part_and_dtd_rejected():
    value = part('<pkg:xmlData><w:a/></pkg:xmlData>')
    with pytest.raises(ValueError):
        decode(package(value + value))
    with pytest.raises(DTDForbidden):
        decode(b'<!DOCTYPE pkg:package [<!ENTITY x "expanded">]>' + package(value))


@pytest.mark.parametrize("name", ["/WORD/document.xml", "/word/%64ocument.xml"])
def test_case_and_uri_alias_collisions_rejected(name):
    value = '<pkg:xmlData><w:a/></pkg:xmlData>'
    with pytest.raises(ValueError):
        decode(package(part(value) + part(value, name)))


@pytest.mark.parametrize("name", ["/word/%2e%2e/outside", "/word/%5coutside", "/word/%00outside", "/word/a?b", "/word/a#b", "/word/a%2fb"])
def test_encoded_unsafe_paths_rejected(name):
    with pytest.raises(ValueError):
        decode(package(part('<pkg:xmlData><w:a/></pkg:xmlData>', name)))


def test_uri_escaped_names_match_relationship_resolution():
    from aspose.words_foss._opc import resolve_target

    source = package(part('<pkg:xmlData><w:a/></pkg:xmlData>', '/word/space%20name.xml'))
    with ZipFile(BytesIO(decode(source))) as archive:
        assert archive.read(resolve_target('word/document.xml', 'space%20name.xml'))


@pytest.mark.parametrize("limit,value", [("MAX_INPUT_BYTES", 1), ("MAX_PART_BYTES", 1),
                                       ("MAX_EXPANDED_BYTES", 1), ("MAX_ZIP_ENTRIES", 1)])
def test_size_and_part_limits_rejected(monkeypatch, limit, value):
    monkeypatch.setattr(_io, limit, value)
    with pytest.raises(ValueError):
        decode(package(part('<pkg:xmlData><w:a/></pkg:xmlData>')))


def test_detector_handles_utf16_and_nonpackages():
    source = package(part('<pkg:xmlData><w:a/></pkg:xmlData>'))
    assert is_flat_opc(source.decode().encode('utf-16'))
    assert not is_flat_opc(b'ordinary text')
    assert not is_flat_opc(b'<package/>')
    assert not is_flat_opc(b'PK\x03\x04')


@pytest.mark.parametrize("payload", ['<pkg:xmlData>lost<w:a/></pkg:xmlData>',
                                   '<other:xmlData xmlns:other="urn:other"><w:a/></other:xmlData>'])
def test_nonpackage_payload_and_uncontained_text_rejected(payload):
    with pytest.raises(ValueError):
        decode(package(part(payload)))


def test_explicit_text_does_not_decode_xml_package():
    options = aw.LoadOptions()
    options.load_format = aw.LoadFormat.TEXT
    source = package(part('<pkg:xmlData><w:a/></pkg:xmlData>'))
    assert '<pkg:package' in aw.Document(BytesIO(source), options).get_text()


def test_missing_content_type_rejected():
    source = package(part('<pkg:xmlData><w:a/></pkg:xmlData>')).replace(b'pkg:contentType="application/xml"', b'')
    with pytest.raises(ValueError):
        decode(source)


@pytest.mark.parametrize("format_name", ["FLAT_OPC", "FLAT_OPC_MACRO_ENABLED",
                                       "FLAT_OPC_TEMPLATE", "FLAT_OPC_TEMPLATE_MACRO_ENABLED"])
def test_native_variants_load_edit_and_preserve_parts(tmp_path, format_name):
    from pathlib import Path

    fixture = Path(__file__).parent / 'fixtures' / 'flat-opc-26.9.zip'
    with ZipFile(fixture) as archive:
        source = archive.read(format_name + '.xml')
    path = tmp_path / 'input.xml'
    path.write_bytes(source)
    options = aw.LoadOptions()
    options.load_format = getattr(aw.LoadFormat, format_name)
    for document in (aw.Document(path), aw.Document(data=source), aw.Document(stream=BytesIO(source)),
                     aw.Document(data=source, load_options=options)):
        text = document.get_text()
        assert all(label in text for label in ('ALPHA', 'BETA', 'CELL'))
        assert '<pkg:package' not in text
        output = tmp_path / 'converted.docx'
        document.save(output, aw.SaveFormat.DOCX)
        assert aw.DocxDocument(output).body.tables
    reader = DocumentReader()
    reader.load_file(path)
    assert reader.to_light_document().sections
    editable = aw.DocxDocument(BytesIO(source))
    with ZipFile(BytesIO(decode(source))) as archive:
        before = {name: archive.read(name) for name in archive.namelist()}
    run = next(run for paragraph in editable.body.paragraphs for run in paragraph.runs if run.text == 'ALPHA')
    run.text = 'EDITED'
    with ZipFile(BytesIO(editable.to_bytes())) as archive:
        after = {name: archive.read(name) for name in archive.namelist()}
    expected_types = {
        'FLAT_OPC': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
        'FLAT_OPC_MACRO_ENABLED': 'application/vnd.ms-word.document.macroEnabled.main+xml',
        'FLAT_OPC_TEMPLATE': 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
        'FLAT_OPC_TEMPLATE_MACRO_ENABLED': 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml',
    }
    types = fromstring(after['[Content_Types].xml'])
    assert next(item.attrib['ContentType'] for item in types
                if item.attrib.get('PartName') == '/word/document.xml') == expected_types[format_name]
    assert before.keys() == after.keys()
    assert [name for name in before if before[name] != after[name]] == ['word/document.xml']
    reopened = aw.DocxDocument(BytesIO(editable.to_bytes()))
    assert reopened.body.tables and any('EDITED' in p.text for p in reopened.body.paragraphs)
    flat = editable.to_flat_opc()
    with ZipFile(BytesIO(decode(flat))) as archive:
        roundtrip = {name: archive.read(name) for name in archive.namelist()}
    assert roundtrip.keys() == after.keys()
    for name, value in after.items():
        if name.endswith(('.xml', '.rels')):
            assert xml_shape(parseString(value, forbid_dtd=True)) == xml_shape(parseString(roundtrip[name], forbid_dtd=True))
        else:
            assert roundtrip[name] == value


@pytest.mark.parametrize("format_value", [24, 25, 26, 27])
@pytest.mark.parametrize("kind", ["enum", "options", "legacy_string"])
def test_conversion_flat_opc_file_bytes_and_direct_writer(tmp_path, format_value, kind):
    from aspose.words_foss.docx_writer import LdmDocxWriter

    fmt = aw.SaveFormat(format_value)
    value = (fmt if kind == "enum" else aw.saving.OoxmlSaveOptions(fmt) if kind == "options"
             else fmt.name.lower())
    doc = aw.Document(BytesIO(b'FLAT_SAVE_SENTINEL'))
    data = doc.to_bytes(value)
    assert is_flat_opc(data)
    assert 'FLAT_SAVE_SENTINEL' in aw.Document(BytesIO(data)).get_text()
    types = fromstring(data)
    assert next(item.attrib[f'{{{PKG}}}contentType'] for item in types
                if item.attrib[f'{{{PKG}}}name'] == '/word/document.xml') == MAIN_TYPES[format_value]
    output = tmp_path / 'output.bin'
    doc.save(output, value)
    assert is_flat_opc(output.read_bytes())
    assert 'FLAT_SAVE_SENTINEL' in aw.Document(BytesIO(output.read_bytes())).get_text()
    options = aw.saving.OoxmlSaveOptions(fmt.name.lower())
    writer = LdmDocxWriter(options)
    writer.write(doc.light_document_model, output)
    assert is_flat_opc(output.read_bytes()) and is_flat_opc(writer.write_to_bytes(doc.light_document_model))


def test_flatten_preserves_binary_and_xml_namespace_comments():
    source = package(part('<pkg:xmlData><!--before--><?test keep?><w:document keep:flag="yes"><w:t>文字</w:t></w:document><!--after--></pkg:xmlData>')
                     + part('<pkg:binaryData>AAH/</pkg:binaryData>', '/word/media/image.bin').replace('application/xml', 'image/png'))
    actual = decode(encode(decode(source)))
    with ZipFile(BytesIO(actual)) as archive:
        assert archive.read('word/media/image.bin') == b'\x00\x01\xff'
        xml = archive.read('word/document.xml')
        assert all(value in xml for value in (b'<!--before-->', b'<?test keep?>', b'<!--after-->'))
        assert fromstring(xml).attrib['{urn:opaque}flag'] == 'yes'


def test_dom_flatten_after_edit_and_atomic_failure(tmp_path, monkeypatch):
    from pathlib import Path

    with ZipFile(Path(__file__).parent / 'fixtures' / 'flat-opc-26.9.zip') as archive:
        source = archive.read('FLAT_OPC.xml')
    doc = aw.DocxDocument(BytesIO(source))
    next(r for p in doc.body.paragraphs for r in p.runs if r.text == 'ALPHA').text = 'EDITED'
    data = doc.to_flat_opc()
    assert is_flat_opc(data)
    reopened = aw.DocxDocument(BytesIO(data))
    assert reopened.body.tables and any('EDITED' in p.text for p in reopened.body.paragraphs)
    output = tmp_path / 'edited.xml'
    doc.save_flat_opc(output)
    assert output.read_bytes() == data
    output.write_bytes(b'ORIGINAL')
    monkeypatch.setattr(_io, 'MAX_INPUT_BYTES', 1)
    with pytest.raises(ValueError):
        doc.save_flat_opc(output)
    assert output.read_bytes() == b'ORIGINAL'
    assert list(tmp_path.iterdir()) == [output]


def test_flatten_rejects_missing_content_type_and_bad_xml():
    data = decode(package(part('<pkg:binaryData>AA==</pkg:binaryData>')))
    with pytest.raises(ExpatError):
        encode(data)  # application/xml with malformed XML bytes must not be emitted as binary.
    buffer = BytesIO()
    with ZipFile(buffer, 'w') as archive:
        archive.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        archive.writestr('word/document.xml', '<document/>')
    with pytest.raises(ValueError, match='content type'):
        encode(buffer.getvalue())


def test_macro_resource_preserved_by_dom_and_diagnosed_by_conversion():
    import warnings
    from pathlib import Path
    from xml.etree.ElementTree import SubElement, tostring

    with ZipFile(Path(__file__).parent / 'fixtures' / 'flat-opc-26.9.zip') as archive:
        source = decode(archive.read('FLAT_OPC_MACRO_ENABLED.xml'))
    with ZipFile(BytesIO(source)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    relationships = fromstring(parts['word/_rels/document.xml.rels'])
    SubElement(relationships, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',
               Id='vbaSentinel', Type='http://schemas.microsoft.com/office/2006/relationships/vbaProject',
               Target='vbaProject.bin')
    parts['word/_rels/document.xml.rels'] = tostring(relationships)
    types = fromstring(parts['[Content_Types].xml'])
    SubElement(types, '{http://schemas.openxmlformats.org/package/2006/content-types}Override',
               PartName='/word/vbaProject.bin', ContentType='application/vnd.ms-office.vbaProject')
    parts['[Content_Types].xml'] = tostring(types)
    parts['word/vbaProject.bin'] = b'opaque resource sentinel; not executable VBA'
    buffer = BytesIO()
    with ZipFile(buffer, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    source = buffer.getvalue()
    flattened = aw.DocxDocument(BytesIO(source)).to_flat_opc()
    with ZipFile(BytesIO(decode(flattened))) as archive:
        assert archive.read('word/vbaProject.bin') == parts['word/vbaProject.bin']
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        doc = aw.Document(BytesIO(flattened))
    assert any(item.code == 'load.vba_project_omitted' for item in doc.diagnostics)


@pytest.mark.parametrize('name', ['../outside.xml', 'word//outside.xml', 'word/./outside.xml',
                                'WORD/document.xml', 'word/directory/'])
def test_flatten_rejects_unsafe_and_colliding_zip_names(name):
    stream = BytesIO()
    with ZipFile(stream, 'w') as archive:
        archive.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                         '<Default Extension="xml" ContentType="application/xml"/></Types>')
        archive.writestr('word/document.xml', '<document/>')
        archive.writestr(name, '<opaque/>')
    with pytest.raises(ValueError):
        encode(stream.getvalue())


def test_flatten_rejects_dtd_and_duplicate_type_declarations():
    for manifest in ('<!DOCTYPE Types><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
                     ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Default Extension="XML" ContentType="application/xml"/></Types>')):
        stream = BytesIO()
        with ZipFile(stream, 'w') as archive:
            archive.writestr('[Content_Types].xml', manifest)
            archive.writestr('word/document.xml', '<document/>')
        with pytest.raises((DTDForbidden, ValueError)):
            encode(stream.getvalue())
