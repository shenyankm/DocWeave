"""Content order and source-story extraction; rendering limits stay explicit."""

from io import BytesIO
from xml.etree.ElementTree import SubElement, tostring
import json
import warnings
from zipfile import ZipFile

import pytest
from defusedxml.ElementTree import fromstring
import pymupdf

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._opc import resolve_target
from aspose.words_foss.docx_reader import DocumentReader
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.docx_writer.constants import CT_URI, PKG_RELS_URI, R_URI, W_URI
from aspose.words_foss.md_writer import LdmMarkdownWriter

W = "{" + W_URI + "}"


def paragraph(text):
    return ldm.Paragraph(children=[ldm.Run(text=text)])


def interleaved_document():
    nested = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph("nested")])])])
    cell = ldm.Cell(children=[paragraph("before"), nested, paragraph("after")])
    table = ldm.Table(rows=[ldm.Row(cells=[cell])])
    return ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[table]))])


def source_package():
    raw = LdmDocxWriter().write_to_bytes(ldm.Document(sections=[ldm.Section(body=ldm.Body(
        children=[paragraph("body")]))]))
    with ZipFile(BytesIO(raw)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    body = ('<w:p><w:pPr><w:sectPr><w:headerReference w:type="first" r:id="first"/>'
            '</w:sectPr></w:pPr><w:r><w:t>body</w:t><w:footnoteReference w:id="7"/>'
            '<w:endnoteReference w:id="9"/></w:r></w:p>'
            '<w:p><w:r><w:t>second</w:t></w:r></w:p>'
            '<w:sectPr><w:headerReference w:type="default" r:id="primary"/></w:sectPr>')
    parts["word/document.xml"] = (f'<w:document xmlns:w="{W_URI}" xmlns:r="{R_URI}">'
                                   f'<w:body>{body}</w:body></w:document>').encode()
    targets = {"first": ("header", "/word/headers/first.xml"),
               "primary": ("header", "headers/primary.xml"),
               "notes": ("footnotes", "../notes/fn%20text.xml"),
               "ends": ("endnotes", "endnotes.xml")}
    # Keep the normal ancillary relationships and add unconventional, URI-encoded targets.
    old = parts["word/_rels/document.xml.rels"].decode().replace("</Relationships>", "")
    parts["word/_rels/document.xml.rels"] = (old + "".join(
        f'<Relationship Id="{rid}" Type="{R_URI}/{kind}" Target="{target}"/>'
        for rid, (kind, target) in targets.items()) + "</Relationships>").encode()
    def p(text):
        return f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'
    parts["word/headers/first.xml"] = (f'<w:hdr xmlns:w="{W_URI}" xmlns:r="{R_URI}">'
        f'{p("first header")}<w:p><w:hyperlink r:id="url"><w:r><w:t>header link</w:t></w:r>'
        '</w:hyperlink></w:p></w:hdr>').encode()
    parts["word/headers/_rels/first.xml.rels"] = (f'<Relationships xmlns="{PKG_RELS_URI}">'
        f'<Relationship Id="url" Type="{R_URI}/hyperlink" Target="https://example.com/header" '
        'TargetMode="External"/></Relationships>').encode()
    parts["word/headers/primary.xml"] = (f'<w:hdr xmlns:w="{W_URI}">{p("primary header")}</w:hdr>').encode()
    parts["notes/fn text.xml"] = (f'<w:footnotes xmlns:w="{W_URI}">'
        f'<w:footnote w:type="separator" w:id="-1">{p("separator")}</w:footnote>'
        f'<w:footnote w:id="7">{p("footnote text")}</w:footnote></w:footnotes>').encode()
    parts["word/endnotes.xml"] = (f'<w:endnotes xmlns:w="{W_URI}">'
        f'<w:endnote w:id="9">{p("endnote text")}</w:endnote></w:endnotes>').encode()
    types = fromstring(parts["[Content_Types].xml"])
    for part, kind in (("notes/fn text.xml", "footnotes"), ("word/endnotes.xml", "endnotes"),
                       ("word/headers/first.xml", "header"), ("word/headers/primary.xml", "header")):
        SubElement(types, f"{{{CT_URI}}}Override", {
            "PartName": "/" + part,
            "ContentType": f"application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml",
        })
    parts["[Content_Types].xml"] = tostring(types, encoding="utf-8", xml_declaration=True)
    output = BytesIO()
    with ZipFile(output, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return output.getvalue()


def load_with_loss(raw):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", aw.ContentLossWarning)
        return aw.Document(BytesIO(raw))


def test_cell_order_survives_model_json_docx_and_structured_export():
    original = interleaved_document()
    restored = ldm.Document.model_validate_json(original.model_dump_json(by_alias=True))
    assert [type(child).__name__ for child in restored.tables[0].rows[0].cells[0].children] == [
        "Paragraph", "Table", "Paragraph"]
    doc = aw.Document(BytesIO(LdmDocxWriter().write_to_bytes(restored)))
    assert doc.get_text() == "before\nnested\nafter"
    assert doc.to_bytes("text").decode() == doc.get_text()
    cell = doc.to_dict()["blocks"][0]["rows"][0]["cells"][0]
    assert [node["type"] for node in cell["blocks"]] == ["paragraph", "table", "paragraph"]
    assert cell["blocks"][2]["text"] == "after"
    assert cell["blocks"][2]["location"].endswith(".children[2]")
    with ZipFile(BytesIO(doc.to_bytes("docx"))) as archive:
        root = fromstring(archive.read("word/document.xml"))
    outer_cell = root.find(".//" + W + "tc")
    assert [node.tag for node in outer_cell if node.tag != W + "tcPr"] == [W + "p", W + "tbl", W + "p"]


def test_pdf_and_markdown_do_not_reorder_or_drop_nested_table_content():
    doc = aw.Document(BytesIO(LdmDocxWriter().write_to_bytes(interleaved_document())))
    with pymupdf.open(stream=doc.to_bytes("pdf"), filetype="pdf") as pdf:
        text = pdf[0].get_text()
        assert text.index("before") < text.index("nested") < text.index("after")
        rectangles = [pdf[0].search_for(value)[0] for value in ("before", "nested", "after")]
        assert rectangles[0].y0 < rectangles[1].y0 < rectangles[2].y0
    with pytest.warns(aw.ContentLossWarning, match="Nested tables"):
        markdown = doc.to_bytes("md").decode()
    assert markdown.index("before") < markdown.index("nested") < markdown.index("after")
    assert any(item.code == "markdown.nested_table_flattened" for item in doc.diagnostics)


def test_legacy_cell_lists_and_horizontal_merge_keep_order():
    legacy = ldm.Cell(paragraphs=[paragraph("first")], tables=[interleaved_document().tables[0]])
    assert legacy.children == legacy.paragraphs + legacy.tables
    a = interleaved_document().tables[0].rows[0].cells[0]
    a.cell_format.horizontal_merge = 1
    b = ldm.Cell(paragraphs=[paragraph("last")], cell_format=ldm.CellFormat(horizontal_merge=2))
    merged, column, span = next(ldm.iter_grid_cells(ldm.Row(cells=[a, b])))
    assert (column, span) == (0, 2)
    assert [type(child).__name__ for child in merged.children] == ["Paragraph", "Table", "Paragraph", "Paragraph"]
    assert merged.children[-1].text == "last"


def test_markdown_grid_span_and_hidden_field_content():
    content = ldm.Paragraph(children=[
        ldm.Run(text="visible"), ldm.Run(text="hidden", font=ldm.Font(hidden=True)),
        ldm.FieldStart(), ldm.Run(text=" DATE "), ldm.FieldSeparator(),
        ldm.Run(text="2026-01-01"), ldm.FieldEnd(),
    ])
    table = ldm.Table(rows=[
        ldm.Row(cells=[ldm.Cell(paragraphs=[content], cell_format=ldm.CellFormat(grid_span=2))]),
        ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph("A")]), ldm.Cell(paragraphs=[paragraph("B")])]),
    ])
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[table]))])
    assert model.text == "visible2026-01-01\nA\nB"
    with pytest.warns(aw.ContentLossWarning, match="Merged cells"):
        markdown = LdmMarkdownWriter().write(model)
    assert "visible2026-01-01" in markdown and "hidden" not in markdown and "DATE" not in markdown
    assert all(line.count("|") == 3 for line in markdown.splitlines())


def test_source_stories_have_note_ids_links_and_section_variant_inheritance():
    doc = load_with_loss(source_package())
    result = json.loads(json.dumps(doc.to_dict()))
    stories = {(story["kind"], story["identifier"] or story["part_name"]): story
               for story in result["source_stories"]}
    assert stories["footnote", "7"]["blocks"][0]["text"] == "footnote text"
    assert stories["endnote", "9"]["blocks"][0]["text"] == "endnote text"
    assert not any(story["identifier"] == "-1" for story in stories.values())
    assert result["blocks"][0]["note_references"] == [
        {"kind": "footnote", "identifier": "7"}, {"kind": "endnote", "identifier": "9"}]
    first = stories["header", "word/headers/first.xml"]
    assert first["references"] == [{"section": 0, "variant": "first", "inherited": False},
                                    {"section": 1, "variant": "first", "inherited": True}]
    assert first["blocks"][1]["runs"][0]["link"] == "https://example.com/header"
    assert stories["header", "word/headers/primary.xml"]["references"] == [
        {"section": 1, "variant": "default", "inherited": False}]


@pytest.mark.parametrize("fmt", ["docx", "pdf", "md", "text"])
def test_extraction_is_not_misrepresented_as_note_rendering(fmt):
    doc = load_with_loss(source_package())
    with pytest.warns(aw.ContentLossWarning, match="Footnotes/endnotes"):
        doc.to_bytes(fmt)
    prefix = "markdown" if fmt == "md" else fmt
    assert any(item.code == prefix + ".notes_omitted" for item in doc.diagnostics)


@pytest.mark.parametrize("fmt", ["docx", "pdf", "md", "text"])
def test_reference_only_notes_also_warn_on_output(fmt):
    model = interleaved_document()
    model.tables[0].rows[0].cells[0].paragraphs[0].note_references = [
        ldm.NoteReference(kind="footnote", identifier="7")]
    doc = aw.Document()
    doc._document = model
    with pytest.warns(aw.ContentLossWarning, match="Footnotes/endnotes"):
        doc.to_bytes(fmt)
    assert any(item.code.endswith(".notes_omitted") for item in doc.diagnostics)


def test_reusing_reader_does_not_leak_stories_relationships_or_styles():
    reader = DocumentReader()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", aw.ContentLossWarning)
        reader.load_bytes(source_package())
    assert reader.to_light_document().source_stories
    plain = LdmDocxWriter().write_to_bytes(ldm.Document(sections=[ldm.Section(body=ldm.Body(
        children=[paragraph("fresh")]))]))
    reader.load_bytes(plain)
    result = reader.to_light_document()
    assert result.text == "fresh" and not result.source_stories and not result.header_paragraphs


def test_dom_horizontal_merges_preserve_text_and_grid_and_reject_partial_cells():
    table = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph(f"{r}-{c}")])
                                        for c in range(3)]) for r in range(2)])
    raw = LdmDocxWriter().write_to_bytes(ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[table]))]))
    doc = aw.DocxDocument(BytesIO(raw))
    grid = doc.body.tables[0]
    target = grid.merge_cells(0, 0, 2)
    assert [p.text for p in target.paragraphs] == ["0-0", "0-1"]
    assert len(grid.rows[0].cells) == 2 and len(grid.rows[1].cells) == 3
    before = doc.to_bytes()
    with pytest.raises(ValueError, match="split"):
        grid.merge_cells(0, 1, 3)
    assert doc.to_bytes() == before
    assert grid.merge_cells(0, 0, 2) is target
    grid.merge_cells(1, 0, 3)
    loaded = aw.Document(BytesIO(doc.to_bytes()))
    assert [c.cell_format.grid_span for c in loaded.light_document_model.tables[0].rows[0].cells] == [2, 1]
    assert loaded.light_document_model.tables[0].rows[1].cells[0].cell_format.grid_span == 3
    assert loaded.get_text().splitlines() == [f"{r}-{c}" for r in range(2) for c in range(3)]


@pytest.mark.parametrize("args", [(True, 0, 1), (2, 0, 1), (0, -1, 1), (0, 0, 4)])
def test_dom_invalid_merge_is_atomic(args):
    table = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph("A")]),
                                        ldm.Cell(paragraphs=[paragraph("B")])])])
    raw = LdmDocxWriter().write_to_bytes(ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[table]))]))
    doc = aw.DocxDocument(BytesIO(raw))
    with pytest.raises(ValueError):
        doc.body.tables[0].merge_cells(*args)
    assert doc.to_bytes() == raw


@pytest.mark.parametrize("target,expected", [("../notes/a%20b.xml", "notes/a b.xml"),
                                            ("/word/styles.xml", "word/styles.xml"),
                                            ("styles.xml", "word/styles.xml")])
def test_opc_internal_uri_resolution(target, expected):
    assert resolve_target("word/document.xml", target) == expected


@pytest.mark.parametrize("target", ["../../outside", "https://example.com/a", "//host/a", "a?query=1",
                                    "a#fragment", "a%5Cb", "a%00b", ""])
def test_opc_targets_cannot_escape_or_become_external(target):
    with pytest.raises(ValueError):
        resolve_target("word/document.xml", target)
