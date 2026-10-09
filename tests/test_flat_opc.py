"""Flat OPC loading, package preservation, and rejection boundaries."""

from io import BytesIO
from zipfile import ZipFile

import pytest
from defusedxml.common import DTDForbidden
from defusedxml.ElementTree import fromstring

import aspose.words_foss as aw
from aspose.words_foss import _io
from aspose.words_foss._flat_opc import decode, is_flat_opc
from aspose.words_foss.docx_reader.document_reader import DocumentReader

PKG = "http://schemas.microsoft.com/office/2006/xmlPackage"


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
