"""Explicit export semantics and independently verified package preservation."""

from io import BytesIO

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
import pytest

import aspose.words_foss as aw
from aspose.words_foss.md_writer import LdmMarkdownWriter
from aspose.words_foss.models import ConversionOptions
from .test_docx_dom import package, payloads
from .test_docx_dom_resources import image_bytes


def business_document():
    source = Document()
    parent = source.styles.add_style("业务标题", WD_STYLE_TYPE.PARAGRAPH)
    child = source.styles.add_style("子标题", WD_STYLE_TYPE.PARAGRAPH)
    child.base_style = parent
    source.add_paragraph("业务章节", style=child)
    stream = BytesIO()
    source.save(stream)
    return aw.Document(BytesIO(stream.getvalue()))


@pytest.mark.parametrize("target, expected", [
    ("Heading 3", "### 业务章节"), ("Quote", "> 业务章节"),
    ("Code", "```\n业务章节\n```"), ("Normal", "业务章节"),
])
def test_business_style_map_inherits_without_mutating_source(tmp_path, target, expected):
    doc = business_document()
    before = doc.light_document_model.model_dump_json()
    options = aw.saving.MarkdownSaveOptions()
    options.style_map = {"业务标题": target}
    options.paragraph_break = "\n"
    assert doc.to_bytes(options).decode().strip() == expected
    output = tmp_path / "mapped.md"
    doc.save(output, options)
    assert output.read_text().strip() == expected
    assert doc.light_document_model.model_dump_json() == before
    assert doc.to_bytes("md").decode().strip() == "业务章节"


def test_style_map_own_style_wins_and_normal_overrides_heading():
    doc = business_document()
    options = aw.saving.MarkdownSaveOptions()
    options.style_map = {"业务标题": "Heading 1", "子标题": "Quote"}
    assert doc.to_bytes(options).decode().strip() == "> 业务章节"
    paragraph = doc.light_document_model.all_paragraphs[0]
    paragraph.paragraph_format.is_heading = True
    paragraph.paragraph_format.outline_level = 0
    options.style_map = {"子标题": "Normal"}
    assert doc.to_bytes(options).decode().strip() == "业务章节"


@pytest.mark.parametrize("mapping", [None, [], {"": "Normal"}, {"x": "Heading 7"},
                                     {"x": []}, {1: "Quote"}])
def test_bad_style_map_is_rejected_before_rendering(mapping):
    with pytest.raises(ValueError, match="style_map"):
        LdmMarkdownWriter(ConversionOptions(style_map=mapping))


def assert_independent_report(doc, original):
    current = payloads(doc.to_bytes())
    report = doc.preservation_report()
    assert report == {
        "added": sorted(current.keys() - original.keys()),
        "removed": sorted(original.keys() - current.keys()),
        "modified": sorted(name for name in current.keys() & original.keys()
                           if current[name] != original[name]),
        "unchanged": sorted(name for name in current.keys() & original.keys()
                            if current[name] == original[name]),
    }
    return report


def test_preservation_report_includes_resources_and_does_not_reset_on_save(tmp_path):
    source = tmp_path / "original.docx"
    original = package(source)
    doc = aw.DocxDocument(source)
    doc.part_xml("word/styles.xml")
    doc.create_paragraph("detached")
    assert assert_independent_report(doc, original)["modified"] == []
    doc.body.paragraphs[0].runs[0].text = "新客户"
    doc.body.paragraphs[0].add_picture(image_bytes(), width=30)
    report = assert_independent_report(doc, original)
    assert "word/document.xml" in report["modified"]
    assert "[Content_Types].xml" in report["modified"]
    assert any(name.startswith("word/media/") for name in report["added"])
    assert "customXml/item.xml" in report["unchanged"]
    doc.save(tmp_path / "edited.docx")
    assert doc.preservation_report() == report
    reopened = aw.DocxDocument(tmp_path / "edited.docx")
    assert reopened.preservation_report()["modified"] == []


def test_resource_replacement_is_compared_with_original_bytes(tmp_path):
    source = tmp_path / "original.docx"
    original = package(source)
    doc = aw.DocxDocument(source)
    name = "customXml/item.xml"
    doc._package.set_parts({name: b"changed"})
    assert assert_independent_report(doc, original)["modified"] == [name]
    doc._package.set_parts({name: original[name]})
    assert assert_independent_report(doc, original)["modified"] == []
