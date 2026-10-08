"""Original-package edits retain unknown parts and fail before overwriting on unsupported edits."""

from zipfile import ZipFile

from defusedxml.ElementTree import fromstring
from defusedxml.common import DTDForbidden
import pytest

from aspose.words_foss.docx_edit import replace_text

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def source_package(path, text='<w:t>客户 &amp; 公司</w:t>', extras=()):
    parts = {
        "word/document.xml": f'<w:document xmlns:w="{W}"><w:body><w:p><w:r>{text}</w:r></w:p></w:body></w:document>'.encode(),
        "word/header1.xml": f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>独立页眉</w:t></w:r></w:p></w:hdr>'.encode(),
        "word/footnotes.xml": f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1"><w:p><w:r><w:t>原始脚注</w:t></w:r></w:p></w:footnote></w:footnotes>'.encode(),
        "customXml/item1.xml": b'<custom>unknown payload</custom>',
        "word/media/image1.bin": b'original binary bytes',
        **dict(extras),
    }
    with ZipFile(path, "w") as archive:
        archive.comment = b"original comment"
        for name, data in parts.items():
            archive.writestr(name, data)
    return parts


def test_literal_patch_retains_all_other_part_payloads_and_xml_prefixes(tmp_path):
    source, output = tmp_path / "original.docx", tmp_path / "edited.docx"
    parts = source_package(source)
    assert replace_text(source, output, {"客户 & 公司": "新 <公司> & 客户"}) == 1
    with ZipFile(output) as archive:
        assert archive.comment == b"original comment"
        for name, payload in parts.items():
            if name != "word/document.xml":
                assert archive.read(name) == payload
        xml = archive.read("word/document.xml")
        assert b"<w:document" in xml and b"<w:t>" in xml
        assert "新 &lt;公司&gt; &amp; 客户" in xml.decode()
        assert next(fromstring(xml).iter(f"{{{W}}}t")).text == "新 <公司> & 客户"


def test_replacements_are_simultaneous_and_preserve_new_edge_whitespace(tmp_path):
    source = tmp_path / "original.docx"
    source_package(source, '<w:t>A B</w:t>')
    assert replace_text(source, source, {"A": " B ", "B": "C"}) == 2
    with ZipFile(source) as archive:
        xml = archive.read("word/document.xml")
        text = next(fromstring(xml).iter(f"{{{W}}}t"))
        assert text.text == " B  C"
        assert text.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"


def test_markup_inside_comments_is_not_edited(tmp_path):
    source, output = tmp_path / "original.docx", tmp_path / "edited.docx"
    source_package(source, '<!-- <w:t>客户</w:t> --><w:t>客户</w:t>')
    replace_text(source, output, {"客户": "甲方"})
    with ZipFile(output) as archive:
        assert '<!-- <w:t>客户</w:t> -->' in archive.read("word/document.xml").decode()


@pytest.mark.parametrize("case", ["split", "signed", "signature_relationship", "dtd", "missing"])
def test_unsupported_edit_preserves_existing_output(tmp_path, case):
    source, output = tmp_path / "original.docx", tmp_path / "edited.docx"
    output.write_bytes(b"old document")
    extras = [("_xmlsignatures/sig1.xml", b"signature")] if case == "signed" else []
    if case == "signature_relationship":
        extras.append(("_rels/.rels", b'<Relationships><Relationship Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="custom/signature.xml"/></Relationships>'))
    if case == "dtd":
        extras.append(("word/styles.xml", b'<!DOCTYPE x><x/>'))
    text = '<w:t>客</w:t><w:t>户</w:t>' if case == "split" else '<w:t>客户</w:t>'
    source_package(source, text, extras)
    with pytest.raises(DTDForbidden if case == "dtd" else ValueError):
        replace_text(source, output, {"不存在" if case == "missing" else "客户": "甲方"})
    assert output.read_bytes() == b"old document"
    assert sorted(path.name for path in tmp_path.iterdir()) == ["edited.docx", "original.docx"]


@pytest.mark.parametrize("replacements", [[], {}, {"": "x"}, {"x": "\x00"}, {"x": "\n"}, {"x": "\ud800"}])
def test_invalid_replacements_are_rejected_before_reading(replacements):
    with pytest.raises(ValueError):
        replace_text("nonexistent.docx", "output.docx", replacements)
