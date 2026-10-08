"""Resource edits retain original payloads, use part-local IDs and fail before mutation."""

from io import BytesIO
import pytest
from defusedxml.ElementTree import fromstring
from PIL import Image

import aspose.words_foss as aw
from aspose.words_foss import _io
from aspose.words_foss.docx_writer.constants import CT_URI, PKG_RELS_URI, REL_HYPERLINK
from aspose.words_foss.docx_writer.drawing import REL_IMAGE, WP_URI
from aspose.words_foss.dom import Hyperlink
from aspose.words_foss.dom.nodes import W
from .test_docx_dom import package, payloads


def image_bytes(kind="PNG"):
    stream = BytesIO()
    Image.new("RGB", (40, 20), "navy").save(stream, format=kind)
    return stream.getvalue()


def relations(raw, part="word/_rels/document.xml.rels"):
    root = fromstring(payloads(raw)[part])
    return {node.get("Id"): node.attrib for node in root.findall(f"{{{PKG_RELS_URI}}}Relationship")}


@pytest.mark.parametrize("default_namespace", [False, True])
def test_hyperlink_ids_reuse_and_retarget_without_affecting_shared_links(tmp_path, default_namespace):
    source = tmp_path / "source.docx"
    before = package(source, '<w:p><w:r><w:t>before </w:t></w:r></w:p><w:sectPr/>',
                     default_namespace=default_namespace)
    doc = aw.DocxDocument(source)
    p = doc.body.paragraphs[0]
    first = p.add_hyperlink("中文 & <链接>", "https://example.com/a?q=1&b=2")
    second = p.add_hyperlink("second", first.target)
    assert isinstance(first, Hyperlink) and first.text == "中文 & <链接>"
    first.runs[0].font.bold = True
    assert first.runs[0].effective_font.bold
    external = [node for node in relations(doc.to_bytes()).values() if node["Type"] == REL_HYPERLINK]
    assert len(external) == 1
    first.target = "mailto:someone@example.com"
    assert second.target == "https://example.com/a?q=1&b=2"
    first.target = "#bookmark"
    assert first.target == "#bookmark"
    reopened = aw.DocxDocument(BytesIO(doc.to_bytes()))
    links = [node for node in reopened.body.paragraphs[0].child_nodes if isinstance(node, Hyperlink)]
    assert [node.target for node in links] == ["#bookmark", second.target]
    actual = payloads(doc.to_bytes())
    for name in before.keys() - {"word/document.xml", "word/_rels/document.xml.rels"}:
        assert actual[name] == before[name]
    assert p.text == "before 中文 & <链接>second"


@pytest.mark.parametrize("source_kind", ["bytes", "stream", "path"])
@pytest.mark.parametrize("kind", ["PNG", "JPEG"])
def test_picture_dimensions_alt_text_and_original_parts(tmp_path, source_kind, kind):
    source = tmp_path / "source.docx"
    before = package(source, '<w:p/><w:sectPr/>')
    doc = aw.DocxDocument(source)
    data = image_bytes(kind)
    image_path = tmp_path / "picture"
    image_path.write_bytes(data)
    picture = {"bytes": data, "stream": BytesIO(data), "path": image_path}[source_kind]
    run = doc.body.paragraphs[0].add_picture(picture, width=72, alternative_text="中文图 & <说明>")
    with pytest.raises(NotImplementedError):
        run.text = "do not delete the picture"
    raw = doc.to_bytes()
    actual = payloads(raw)
    for name in before.keys() - {"word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"}:
        assert actual[name] == before[name]
    parsed = aw.Document(BytesIO(raw))
    shape = parsed.get_child_nodes(aw.NodeType.SHAPE, True)[0]
    assert (shape.width, shape.height) == (72, 36)
    assert shape.alternative_text == "中文图 & <说明>"
    assert shape.image_data.image_bytes == data
    rebuilt = aw.Document(BytesIO(parsed.to_bytes("docx")))
    assert rebuilt.get_child_nodes(aw.NodeType.SHAPE, True)[0].alternative_text == shape.alternative_text
    root = fromstring(actual["[Content_Types].xml"])
    part = next(name for name in actual.keys() - before.keys() if name.startswith("word/media/image"))
    assert root.find(f"{{{CT_URI}}}Override[@PartName='/{part}']").get("ContentType") == (
        "image/png" if kind == "PNG" else "image/jpeg")


def test_picture_dedup_part_local_relationships_and_package_wide_drawing_ids(tmp_path):
    source = tmp_path / "source.docx"
    package(source, '<w:p/><w:sectPr/>', extras={"word/header1.xml": (
        f'<w:hdr xmlns:w="{W}" xmlns:wp="{WP_URI}"><w:p><wp:docPr id="1" name="existing"/></w:p></w:hdr>'
    ).encode()})
    doc = aw.DocxDocument(source)
    p = doc.body.paragraphs[0]
    for _ in range(2):
        p.add_picture(image_bytes(), height=30)
    doc.story("word/header1.xml").paragraphs[0].add_picture(image_bytes())
    raw = doc.to_bytes()
    parts = payloads(raw)
    assert len([name for name in parts if name.endswith(".png")]) == 1
    assert len([node for node in relations(raw).values() if node["Type"] == REL_IMAGE]) == 1
    header_rels = relations(raw, "word/_rels/header1.xml.rels")
    assert len(header_rels) == 1
    ids = [node.get("id") for name in ("word/document.xml", "word/header1.xml")
           for node in fromstring(parts[name]).iter(f"{{{WP_URI}}}docPr")]
    assert len(ids) == len(set(ids)) == 4
    body_targets = {node["Target"] for node in relations(raw).values() if node["Type"] == REL_IMAGE}
    assert body_targets == {node["Target"] for node in header_rels.values()}
    reopened = aw.DocxDocument(BytesIO(raw))
    assert len(reopened.part_names) == len(parts)
    assert payloads(reopened.to_bytes()) == parts


@pytest.mark.parametrize("part", ["stories/header.xml", "stories/header.part"])
def test_drawing_ids_include_declared_xml_parts_outside_word(tmp_path, part):
    source = tmp_path / "source.docx"
    before = package(source, '<w:p/>')
    types = before["[Content_Types].xml"].replace(b'</Types>', (
        f'<Override PartName="/{part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '</Types>').encode())
    package(source, '<w:p/>', extras={"[Content_Types].xml": types, part: (
        f'<w:hdr xmlns:w="{W}" xmlns:wp="{WP_URI}"><w:p><wp:docPr id="1"/></w:p></w:hdr>'
    ).encode()})
    doc = aw.DocxDocument(source)
    doc.body.paragraphs[0].add_picture(image_bytes())
    root = fromstring(payloads(doc.to_bytes())["word/document.xml"])
    assert root.find(f".//{{{WP_URI}}}docPr").get("id") == "2"


@pytest.mark.parametrize("target", ["", "#", "javascript:alert(1)", "file:///etc/passwd",
                                     "https://example.com\n", "https://example.com/a b"])
def test_invalid_links_are_atomic(tmp_path, target):
    source = tmp_path / "source.docx"
    before = package(source, '<w:p/>')
    doc = aw.DocxDocument(source)
    with pytest.raises(ValueError):
        doc.body.paragraphs[0].add_hyperlink("text", target)
    assert payloads(doc.to_bytes()) == before


@pytest.mark.parametrize("width,height", [(0, None), (-1, None), (float("nan"), 3),
                                         (True, None), (None, "invalid"), (3, float("inf"))])
def test_invalid_picture_dimensions_are_atomic(tmp_path, width, height):
    source = tmp_path / "source.docx"
    before = package(source, '<w:p/>')
    doc = aw.DocxDocument(source)
    with pytest.raises(ValueError):
        doc.body.paragraphs[0].add_picture(image_bytes(), width=width, height=height)
    assert payloads(doc.to_bytes()) == before


def test_picture_and_package_limits_fail_before_mutation(tmp_path, monkeypatch):
    source = tmp_path / "source.docx"
    before = package(source, '<w:p/>')
    doc = aw.DocxDocument(source)
    monkeypatch.setattr(_io, "MAX_IMAGE_PIXELS", 10)
    with pytest.raises(ValueError, match="pixels"):
        doc.body.paragraphs[0].add_picture(image_bytes())
    monkeypatch.setattr(_io, "MAX_IMAGE_PIXELS", 25_000_000)
    monkeypatch.setattr(_io, "MAX_ZIP_ENTRIES", len(before))
    with pytest.raises(ValueError, match="too many"):
        doc.body.paragraphs[0].add_picture(image_bytes())
    assert payloads(doc.to_bytes()) == before


def test_missing_rels_are_created_and_duplicates_are_not_silently_repaired(tmp_path):
    source = tmp_path / "source.docx"
    package(source, '<w:p/>')
    doc = aw.DocxDocument(source)
    link = doc.story("word/header1.xml").paragraphs[0].add_hyperlink("header", "https://example.com")
    assert link.target == "https://example.com"
    assert "word/_rels/header1.xml.rels" in doc.part_names
    bad_rels = (f'<Relationships xmlns="{PKG_RELS_URI}">'
                '<Relationship Id="same" Type="a" Target="a"/>'
                '<Relationship Id="same" Type="b" Target="b"/></Relationships>').encode()
    before = package(source, '<w:p/>', extras={"word/_rels/document.xml.rels": bad_rels})
    doc = aw.DocxDocument(source)
    with pytest.raises(ValueError, match="duplicate"):
        doc.body.paragraphs[0].add_hyperlink("text", "https://example.com")
    assert payloads(doc.to_bytes()) == before


def test_python_docx_can_open_generated_resource_relationships(tmp_path):
    python_docx = pytest.importorskip("docx")
    stream = BytesIO()
    original = python_docx.Document()
    original.add_paragraph("source")
    original.save(stream)
    doc = aw.DocxDocument(BytesIO(stream.getvalue()))
    doc.body.paragraphs[0].add_picture(image_bytes(), width=90)
    doc.body.paragraphs[0].add_hyperlink("link", "https://example.com")
    loaded = python_docx.Document(BytesIO(doc.to_bytes()))
    assert len(loaded.inline_shapes) == 1
    assert loaded.inline_shapes[0].width.pt == 90
    assert loaded.paragraphs[0].hyperlinks[0].url == "https://example.com"
