"""Original-package DOM: local edits, ordered nodes, and explicit safety boundaries."""

from io import BytesIO
from zipfile import ZipFile, ZipInfo

import pytest
from defusedxml.common import DTDForbidden
from defusedxml.ElementTree import fromstring

import aspose.words_foss as aw
from aspose.words_foss.dom import UnknownNode
from aspose.words_foss.dom.nodes import W

MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
EXT = "urn:retained-extension"


def package(path, body=None, extras=None, default_namespace=False):
    if body is None:
        body = ('<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>'
                '<w:r><w:rPr><w:b/></w:rPr><w:t>客户</w:t></w:r>'
                '<w:r><w:rPr><w:i/></w:rPr><w:t>公司</w:t></w:r></w:p>'
                '<w:p><ext:opaque ext:value="unchanged"/></w:p><w:sectPr/>')
    xml = (f'<?xml version="1.0" encoding="UTF-8"?>'
           f'<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:ext="{EXT}" '
           'mc:Ignorable="ext"><!--retain comment--><?retain instruction?>'
           f'<w:body>{body}</w:body></w:document>')
    if default_namespace:
        xml = xml.replace(f'xmlns:w="{W}"', f'xmlns="{W}" xmlns:w="{W}"')
        xml = xml.replace('<w:', '<').replace('</w:', '</')
    parts = {
        "[Content_Types].xml": (
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            b'<Default Extension="xml" ContentType="application/xml"/>'
            b'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            b'<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            b'</Types>'),
        "_rels/.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/document.xml": xml.encode(),
        "word/_rels/document.xml.rels": b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        "word/styles.xml": f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style></w:styles>'.encode(),
        "word/header1.xml": f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>页眉</w:t></w:r></w:p></w:hdr>'.encode(),
        "customXml/item.xml": b'<opaque xmlns="urn:custom">original</opaque>',
        "word/media/image.bin": b'original media',
        **(extras or {}),
    }
    with ZipFile(path, "w") as archive:
        archive.comment = b"original archive comment"
        for name, data in parts.items():
            info = ZipInfo(name, date_time=(2020, 2, 3, 4, 5, 6))
            info.comment = b"entry comment"
            archive.writestr(info, data)
    return parts


def payloads(data):
    with ZipFile(BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def test_no_edit_save_retains_every_part_even_after_reading_and_detached_creation(tmp_path):
    source, output = tmp_path / "source.docx", tmp_path / "output.docx"
    parts = package(source)
    doc = aw.DocxDocument(source)
    assert doc.body.paragraphs[0].text == "客户公司"
    assert doc.part_xml("word/styles.xml")
    doc.create_paragraph("not inserted")
    doc.save(output)
    assert payloads(output.read_bytes()) == parts
    with ZipFile(output) as archive:
        assert archive.comment == b"original archive comment"
        assert archive.getinfo("word/document.xml").date_time == (2020, 2, 3, 4, 5, 6)
        assert archive.getinfo("word/document.xml").comment == b"entry comment"


@pytest.mark.parametrize("default_namespace", [False, True])
def test_local_edits_preserve_unknown_xml_namespaces_and_other_parts(tmp_path, default_namespace):
    source = tmp_path / "source.docx"
    original = package(source, default_namespace=default_namespace)
    doc = aw.DocxDocument(source)
    paragraph = doc.body.paragraphs[0]
    assert paragraph.runs[0].font.bold is True
    assert paragraph.runs[1].font.bold is None
    paragraph.runs[0].text = " 新客户 & <公司> "
    paragraph.runs[0].font.bold = False
    paragraph.runs[0].font.italic = True
    paragraph.paragraph_format.alignment = "center"
    paragraph.paragraph_format.style_id = "Title"
    actual = payloads(doc.to_bytes())
    assert {name: data for name, data in actual.items() if name != "word/document.xml"} == {
        name: data for name, data in original.items() if name != "word/document.xml"}
    xml = actual["word/document.xml"]
    assert b'xmlns:ext="urn:retained-extension"' in xml
    assert b'mc:Ignorable="ext"' in xml
    assert b'<!--retain comment-->' in xml
    assert b'<?retain instruction?>' in xml
    root = fromstring(xml)
    assert root.find(f'.//{{{EXT}}}opaque').get(f'{{{EXT}}}value') == "unchanged"
    assert root.find(f'.//{{{W}}}t').get('{http://www.w3.org/XML/1998/namespace}space') == "preserve"
    reopened = aw.DocxDocument(BytesIO(doc.to_bytes()))
    p = reopened.body.paragraphs[0]
    assert p.runs[0].text == " 新客户 & <公司> "
    assert p.runs[0].font.bold is False
    assert p.runs[0].font.italic is True
    assert p.paragraph_format.alignment == "center"
    assert p.paragraph_format.style_id == "Title"


def test_direct_format_none_restores_inheritance_and_keeps_order(tmp_path):
    source = tmp_path / "source.docx"
    package(source, '<w:p><w:pPr><w:jc w:val="left"/><w:rPr/></w:pPr><w:r><w:rPr><w:i/><w:sz w:val="24"/></w:rPr><w:t>A</w:t></w:r></w:p>')
    doc = aw.DocxDocument(source)
    p = doc.body.paragraphs[0]
    p.runs[0].font.bold = True
    p.paragraph_format.style_id = "Title"
    root = fromstring(payloads(doc.to_bytes())["word/document.xml"])
    assert [node.tag.split('}')[-1] for node in root.find(f'.//{{{W}}}rPr')] == []  # paragraph-mark font
    run_properties = root.find(f'.//{{{W}}}r/{{{W}}}rPr')
    assert [node.tag.split('}')[-1] for node in run_properties] == ["b", "i", "sz"]
    assert [node.tag.split('}')[-1] for node in root.find(f'.//{{{W}}}pPr')] == ["pStyle", "jc", "rPr"]
    p.runs[0].font.bold = None
    p.paragraph_format.alignment = None
    assert p.runs[0].font.bold is None
    assert p.paragraph_format.alignment is None
    with pytest.raises(ValueError):
        p.paragraph_format.style_id = "missing"


@pytest.mark.parametrize("runs,old,new,expected,count", [
    (["客户", "公司"], "客户公司", " 新公司 ", " 新公司 ", 1),
    (["ababa"], "ab", "XYZ", "XYZXYZa", 2),
    (["a", "ba", "ba", "b"], "ab", "X", "XXX", 3),
    (["a", "", "b"], "ab", "", "", 1),
    (["aba", "ba"], "ab", "abcd", "abcdabcda", 2),
    (["a", "b"], "missing", "x", "ab", 0),
])
def test_cross_run_replacement_retains_first_format_and_original_run_nodes(tmp_path, runs, old, new, expected, count):
    source = tmp_path / "source.docx"
    body = '<w:p>' + ''.join(f'<w:r><w:rPr><w:b/></w:rPr><w:t>{text}</w:t></w:r>' for text in runs) + '</w:p>'
    package(source, body)
    doc = aw.DocxDocument(source)
    p = doc.body.paragraphs[0]
    original_runs = p.runs
    assert p.replace_text(old, new) == count
    assert p.text == expected
    assert p.runs == original_runs
    assert all(run.font.bold for run in p.runs)
    assert aw.DocxDocument(BytesIO(doc.to_bytes())).body.paragraphs[0].text == expected


def test_ordered_structure_and_parent_ownership(tmp_path):
    source = tmp_path / "source.docx"
    package(source, '<w:p><w:r><w:t>first</w:t></w:r></w:p><w:sectPr/>')
    doc = aw.DocxDocument(source)
    first = doc.body.paragraphs[0]
    table = doc.create_table(1, 2)
    doc.body.append_child(table)
    last = doc.create_paragraph("last")
    doc.body.append_child(last)
    assert [type(node).__name__ for node in doc.body.child_nodes] == ["Paragraph", "Table", "Paragraph", "UnknownNode"]
    cell = table.rows[0].cells[0]
    nested = doc.create_table(1, 1)
    cell.append_child(nested)
    cell.insert_before(doc.create_paragraph("before nested"), nested)
    assert [type(node).__name__ for node in cell.child_nodes] == ["Paragraph", "Table", "Paragraph"]
    assert nested.parent_node is cell
    assert cell.owner_document is doc
    with pytest.raises(ValueError):
        nested.rows[0].cells[0].append_child(table)
    with pytest.raises(ValueError):
        first.append_child(table)
    with pytest.raises(ValueError):
        cell.paragraphs[-1].remove()
    with pytest.raises(NotImplementedError):
        table.rows[0].remove()
    copied = first.clone()
    doc.body.insert_before(copied, first)
    assert copied is not first and copied.text == first.text
    copied.remove()
    doc.body.insert_before(last, first)
    assert doc.body.paragraphs[0] is last
    assert last.parent_node is doc.body
    assert len(doc.get_child_nodes(aw.NodeType.TABLE, deep=True)) == 2
    assert isinstance(doc.body.child_nodes[-1], UnknownNode)
    root = fromstring(payloads(doc.to_bytes())["word/document.xml"])
    assert list(root.find(f'{{{W}}}body'))[-1].tag == f'{{{W}}}sectPr'


@pytest.mark.parametrize("kind", ["paragraph", "run", "table", "row", "cell"])
def test_clone_depth_preserves_formatting_and_owner(tmp_path, kind):
    source = tmp_path / "clone.docx"
    package(source, '<w:p><w:pPr><w:ind w:left="340"/></w:pPr>'
            '<w:r><w:rPr><w:b/></w:rPr><w:t>ALPHA</w:t></w:r></w:p>'
            '<w:tbl><w:tblPr><w:tblW w:w="2400" w:type="dxa"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>'
            '<w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr>'
            '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>'
            '<w:p><w:r><w:t>CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/>')
    doc = aw.DocxDocument(source)
    paragraph, table = doc.body.paragraphs[0], doc.body.tables[0]
    node = {"paragraph": paragraph, "run": paragraph.runs[0], "table": table,
            "row": table.rows[0], "cell": table.rows[0].cells[0]}[kind]
    original = node.xml
    copied = node.clone(False)
    assert copied.parent_node is None and copied.owner_document is doc
    assert copied.part_name == node.part_name and node.xml == original
    root = fromstring(copied.xml)
    properties = {"paragraph": "pPr", "run": "rPr", "table": "tblPr", "row": "trPr", "cell": "tcPr"}
    assert root.find(f'{{{W}}}{properties[kind]}') is not None
    if kind == "run":
        assert copied.text == "ALPHA" and copied.font.bold
        copied.text = "COPY"
        assert node.text == "ALPHA"
    else:
        assert copied.child_nodes == ()
        assert not any(child.tag in {f'{{{W}}}{name}' for name in ("p", "r", "tr", "tc", "tbl")} for child in root)
    if kind == "table":
        assert root.find(f'{{{W}}}tblGrid/{{{W}}}gridCol').get(f'{{{W}}}w') == "2400"
    assert doc.to_bytes() == source.read_bytes()
    assert node.clone(True).xml == node.clone().xml
    if kind in {"paragraph", "run"}:
        parent = doc.body if kind == "paragraph" else paragraph
        parent.append_child(copied)
        reopened = aw.DocxDocument(BytesIO(doc.to_bytes()))
        if kind == "paragraph":
            saved = reopened.body.paragraphs[-1]
            assert saved.text == "" and not saved.runs
            assert fromstring(saved.xml).find(f'{{{W}}}pPr/{{{W}}}ind').get(f'{{{W}}}left') == "340"
        else:
            assert reopened.body.paragraphs[0].runs[-1].text == "COPY"
            assert reopened.body.paragraphs[0].runs[-1].font.bold


@pytest.mark.parametrize("value", [None, 0, 1, "false", []])
def test_clone_rejects_non_boolean_depth_without_mutation(tmp_path, value):
    source = tmp_path / "clone.docx"
    package(source, '<w:p><w:r><w:t>ALPHA</w:t></w:r></w:p>')
    doc = aw.DocxDocument(source)
    with pytest.raises(TypeError, match="boolean"):
        doc.body.paragraphs[0].clone(value)
    assert doc.to_bytes() == source.read_bytes()


@pytest.mark.parametrize("kind", ["paragraph", "run", "table", "cell_paragraph", "header_paragraph"])
def test_first_edit_removal_persists_and_preserves_unaffected_parts(tmp_path, kind):
    source = tmp_path / "source.docx"
    header = f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>first header</w:t></w:r></w:p><w:p><w:r><w:t>second header</w:t></w:r></w:p></w:hdr>'
    original = package(source,
                       '<w:p><w:r><w:t>FIRST</w:t></w:r><w:r><w:t>tail</w:t></w:r></w:p>'
                       '<w:p><w:r><w:t>SECOND</w:t></w:r></w:p>'
                       '<w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc>'
                       '<w:p><w:r><w:t>first cell</w:t></w:r></w:p>'
                       '<w:p><w:r><w:t>second cell</w:t></w:r></w:p>'
                       '</w:tc></w:tr></w:tbl><w:sectPr/>',
                       extras={"word/header1.xml": header.encode()})
    doc = aw.DocxDocument(source)
    part = "word/header1.xml" if kind == "header_paragraph" else "word/document.xml"
    targets = {"paragraph": lambda: doc.body.paragraphs[0],
               "run": lambda: doc.body.paragraphs[0].runs[0],
               "table": lambda: doc.body.tables[0],
               "cell_paragraph": lambda: doc.body.tables[0].rows[0].cells[0].paragraphs[0],
               "header_paragraph": lambda: doc.story(part).paragraphs[0]}
    removed = targets[kind]()
    removed.remove()
    assert removed.parent_node is None and removed.owner_document is doc
    saved = payloads(doc.to_bytes())
    assert {name for name in original if original[name] != saved[name]} == {part}
    expected_removed = {"paragraph": "FIRSTtail", "run": "FIRST", "table": "first cellsecond cell",
                        "cell_paragraph": "first cell", "header_paragraph": "first header"}[kind]
    visible = "".join(node.text or "" for node in fromstring(saved[part]).iter(f'{{{W}}}t'))
    assert expected_removed not in visible
    assert ("second header" if kind == "header_paragraph" else "SECOND") in visible
    reloaded = aw.DocxDocument(BytesIO(doc.to_bytes()))
    assert reloaded.part_xml(part) == doc.part_xml(part)


@pytest.mark.parametrize("inline", [
    '<w:bookmarkStart w:id="1" w:name="b"/>',
    '<w:hyperlink><w:r><w:t>link</w:t></w:r></w:hyperlink>',
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    '<w:r><w:tab/></w:r>',
    '<w:ins w:id="1"><w:r><w:t>revision</w:t></w:r></w:ins>',
])
def test_complex_replacement_and_structural_delete_fail_before_mutation(tmp_path, inline):
    source = tmp_path / "source.docx"
    original = package(source, f'<w:p><w:r><w:t>A</w:t></w:r>{inline}</w:p>')
    doc = aw.DocxDocument(source)
    p = doc.body.paragraphs[0]
    with pytest.raises(NotImplementedError):
        p.replace_text("A", "B")
    with pytest.raises(NotImplementedError):
        p.remove()
    assert payloads(doc.to_bytes()) == original


def test_unsupported_ancestor_is_readable_but_not_editable(tmp_path):
    source = tmp_path / "source.docx"
    package(source, '<w:sdt><w:sdtContent><w:p><w:r><w:t>protected</w:t></w:r></w:p></w:sdtContent></w:sdt>')
    doc = aw.DocxDocument(source)
    run = doc.get_child_nodes(aw.NodeType.RUN, deep=True)[0]
    assert run.text == "protected"
    with pytest.raises(NotImplementedError):
        run.text = "changed"


def test_headers_and_cross_part_document_rejection(tmp_path):
    source = tmp_path / "source.docx"
    original = package(source)
    doc = aw.DocxDocument(source)
    header = doc.story("word/header1.xml")
    header.paragraphs[0].runs[0].text = "新页眉"
    header.append_child(doc.create_paragraph("追加", part_name="word/header1.xml"))
    with pytest.raises(ValueError):
        header.append_child(doc.body.paragraphs[0])
    other = aw.DocxDocument(source)
    with pytest.raises(ValueError):
        doc.body.append_child(other.body.paragraphs[0])
    actual = payloads(doc.to_bytes())
    assert actual["word/document.xml"] == original["word/document.xml"]
    assert aw.DocxDocument(BytesIO(doc.to_bytes())).story("word/header1.xml").paragraphs[0].text == "新页眉"


@pytest.mark.parametrize("bad", ["\x00", "\ud800", "\uffff", "\n", "\t", 42])
def test_invalid_text_does_not_modify_part(tmp_path, bad):
    source = tmp_path / "source.docx"
    original = package(source)
    doc = aw.DocxDocument(source)
    with pytest.raises(ValueError):
        doc.body.paragraphs[0].runs[0].text = bad
    assert payloads(doc.to_bytes()) == original


@pytest.mark.parametrize("extras", [
    {"_xmlsignatures/sig.xml": b"signature"},
    {"_rels/.rels": b'<Relationships><Relationship Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"/></Relationships>'},
    {"word/document.xml": b'<!DOCTYPE x><x/>'},
    {"word/document.xml": b'<document xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main"><body/></document>'},
])
def test_signed_dtd_and_strict_packages_rejected(tmp_path, extras):
    source = tmp_path / "source.docx"
    package(source, extras=extras)
    with pytest.raises((ValueError, DTDForbidden)):
        aw.DocxDocument(source)


def test_atomic_save_retains_output_on_serialization_failure(tmp_path, monkeypatch):
    source, output = tmp_path / "source.docx", tmp_path / "output.docx"
    package(source)
    output.write_bytes(b"existing")
    doc = aw.DocxDocument(source)

    def fail():
        raise RuntimeError("serialization failed")

    monkeypatch.setattr(doc._package, "to_bytes", fail)
    with pytest.raises(RuntimeError):
        doc.save(output)
    assert output.read_bytes() == b"existing"
    assert sorted(p.name for p in tmp_path.iterdir()) == ["output.docx", "source.docx"]


def test_ldm_snapshot_contains_edits_but_is_independent(tmp_path):
    source = tmp_path / "source.docx"
    package(source, '<w:p><w:r><w:t>original</w:t></w:r></w:p>')
    doc = aw.DocxDocument(source)
    doc.body.paragraphs[0].runs[0].text = "edited"
    snapshot = doc.to_light_document()
    assert snapshot.all_paragraphs[0].text == "edited"
    snapshot.all_paragraphs[0].runs[0].text = "snapshot only"
    assert doc.body.paragraphs[0].text == "edited"
    assert aw.Document(BytesIO(doc.to_bytes())).get_text() == "edited"


@pytest.mark.parametrize("marker", [
    '<w:bookmarkStart w:id="1" w:name="b"/>',
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    '<w:pPr><w:rPr><w:rPrChange w:id="1"/></w:rPr></w:pPr>',
])
def test_structure_cannot_cross_ranges_outside_edited_subtree(tmp_path, marker):
    source = tmp_path / "source.docx"
    original = package(source, f'<w:p>{marker}</w:p><w:p><w:r><w:t>plain</w:t></w:r></w:p>')
    doc = aw.DocxDocument(source)
    plain = doc.body.paragraphs[1]
    with pytest.raises(NotImplementedError):
        plain.remove()
    with pytest.raises(NotImplementedError):
        plain.clone()
    with pytest.raises(NotImplementedError):
        doc.body.append_child(doc.create_paragraph("new"))
    assert payloads(doc.to_bytes()) == original


def test_multiple_text_nodes_and_new_default_namespace_nodes(tmp_path):
    source = tmp_path / "source.docx"
    package(source, '<w:p><w:r><w:t>ab</w:t><w:t>ab</w:t></w:r></w:p><w:sectPr/>', default_namespace=True)
    doc = aw.DocxDocument(source)
    p = doc.body.paragraphs[0]
    assert p.replace_text("ba", "XYZ") == 1
    assert p.text == "aXYZb"
    p.runs[0].text = "reset"
    new = doc.create_paragraph("created")
    new.runs[0].font.italic = True
    doc.body.append_child(new)
    doc.body.append_child(doc.create_table(1, 1))
    reopened = aw.DocxDocument(BytesIO(doc.to_bytes()))
    assert [p.text for p in reopened.body.paragraphs] == ["reset", "created"]
    assert reopened.body.paragraphs[1].runs[0].font.italic is True
    assert len(reopened.body.tables[0].rows[0].cells) == 1


@pytest.mark.parametrize("rows,columns", [(0, 1), (1, 0), (1, 1025), (True, 1), (1, 2.5)])
def test_invalid_table_dimensions_are_rejected(tmp_path, rows, columns):
    source = tmp_path / "source.docx"
    original = package(source)
    doc = aw.DocxDocument(source)
    with pytest.raises(ValueError):
        doc.create_table(rows, columns)
    assert payloads(doc.to_bytes()) == original


def test_in_place_save_preserves_other_parts(tmp_path):
    source = tmp_path / "source.docx"
    original = package(source)
    doc = aw.DocxDocument(source)
    doc.body.paragraphs[0].runs[0].text = "changed"
    doc.save(source)
    actual = payloads(source.read_bytes())
    assert actual["word/styles.xml"] == original["word/styles.xml"]
    assert aw.DocxDocument(source).body.paragraphs[0].text == "changed公司"


def test_text_edit_retains_comments_processing_instructions_and_attributes(tmp_path):
    source = tmp_path / "source.docx"
    package(source, '<w:p><w:r><w:t ext:flag="retain">A<!--inside--><?inside keep?><![CDATA[B]]></w:t></w:r></w:p>')
    doc = aw.DocxDocument(source)
    run = doc.body.paragraphs[0].runs[0]
    assert run.text == "AB"
    run.text = "new ]]> text"
    xml = payloads(doc.to_bytes())["word/document.xml"]
    assert b'<!--inside-->' in xml and b'<?inside keep?>' in xml
    assert fromstring(xml).find(f'.//{{{W}}}t').get(f'{{{EXT}}}flag') == "retain"
    assert aw.DocxDocument(BytesIO(doc.to_bytes())).body.paragraphs[0].text == "new ]]> text"
