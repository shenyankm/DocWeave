"""XML character preservation through conversion, package edits and Flat OPC."""

from io import BytesIO
from xml.dom.minidom import parseString
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.docx_writer.xml_utils import escape_attr, escape_text
from .test_docx_dom import package, payloads

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
VALUES = ["A\tB\nC\rD", "A😀𠀀\U0010ffffB", 'A & < > " B']


@pytest.mark.parametrize("value", VALUES)
def test_writer_keeps_xml_characters_in_text_and_attributes(value):
    node = ET.fromstring(f'<r value="{escape_attr(value)}">{escape_text(value)}</r>')
    assert node.get("value") == value
    assert node.text == value.replace("\r", "\n")


def test_writer_still_filters_xml_illegal_characters():
    value = "a\x00\x0b\x0c\ud800\udfff\ufffe\uffffb"
    assert escape_attr(value) == escape_text(value) == "ab"


@pytest.mark.parametrize("value", VALUES)
def test_dom_serializer_preserves_namespaces_comments_and_attribute_values(value):
    from aspose.words_foss.utils.xml_helpers import serialize_xml

    tree = parseString('<r xmlns="urn:owned"><child xmlns=""/><!--keep--><?keep data?></r>')
    tree.documentElement.setAttribute("value", value)
    raw = serialize_xml(tree)
    node = ET.fromstring(raw)
    assert node.get("value") == value
    assert node[0].tag == "child"
    assert b"<!--keep-->" in raw and b"<?keep data?>" in raw


@pytest.mark.parametrize("value", VALUES)
@pytest.mark.parametrize("flat_opc", [False, True])
def test_public_package_edit_preserves_font_attributes(tmp_path, value, flat_opc):
    path = tmp_path / "owned.docx"
    body = ('<w:p><w:r><w:rPr><w:rFonts w:ascii="' + escape_attr(value)
            + '"/></w:rPr><w:t>before</w:t></w:r></w:p>')
    original = package(path, body)
    types = original["[Content_Types].xml"].replace(
        b"</Types>", b'<Default Extension="bin" ContentType="application/octet-stream"/></Types>')
    package(path, body, extras={"[Content_Types].xml": types})
    doc = aw.DocxDocument(path)
    doc.body.paragraphs[0].runs[0].text = "after 😀"
    saved = doc.to_flat_opc() if flat_opc else doc.to_bytes()
    reopened = aw.DocxDocument(BytesIO(saved))
    data = payloads(reopened.to_bytes())
    fonts = ET.fromstring(data["word/document.xml"]).find(".//" + W + "rFonts")
    assert fonts.get(W + "ascii") == value
    assert reopened.body.paragraphs[0].text == "after 😀"
    assert ET.tostring(ET.fromstring(data["customXml/item.xml"])) == ET.tostring(
        ET.fromstring(original["customXml/item.xml"]))


@pytest.mark.parametrize("value", VALUES)
def test_conversion_writer_preserves_font_name_and_supplementary_text(tmp_path, value):
    path = tmp_path / "owned.docx"
    package(path, '<w:p><w:r><w:t>text 😀𠀀</w:t></w:r></w:p>')
    model = aw.Document(path).light_document_model
    run = next(node for node in model.get_child_nodes(aw.NodeType.RUN, True))
    run.font.name = value
    saved = LdmDocxWriter().write_to_bytes(model)
    with ZipFile(BytesIO(saved)) as archive:
        node = ET.fromstring(archive.read("word/document.xml"))
    assert node.find(".//" + W + "rFonts").get(W + "ascii") == value
    assert node.find(".//" + W + "t").text == "text 😀𠀀"
