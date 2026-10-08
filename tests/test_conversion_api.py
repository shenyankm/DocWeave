"""Public in-memory output, structured diagnostics and format boundaries."""

from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict
from io import BytesIO
import json
import warnings
from zipfile import ZipFile

import pytest
from pypdf import PdfReader

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer import DocxWriterLossyWarning, LdmDocxWriter
from aspose.words_foss.diagnostics import collect_diagnostics, warn
from aspose.words_foss.pdf_writer import PdfMissingGlyphWarning
from aspose.words_foss.rtf_reader import RtfFileReader

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def package(body):
    stream = BytesIO()
    with ZipFile(stream, "w") as archive:
        archive.writestr("word/document.xml", f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}</w:body></w:document>')
    return stream.getvalue()


def document(text="中文 **粗体** [链接](https://example.com)"):
    return aw.Document(BytesIO(text.encode()), aw.MarkdownLoadOptions())


@pytest.mark.parametrize("mode", ["path", "stream", "bytes", "reader_path", "reader_stream", "reader_bytes"])
def test_standard_text_rtf_has_actionable_error(tmp_path, mode):
    raw = b"{\\rtf1\\ansi ordinary text}"
    path = tmp_path / "standard.rtf"
    path.write_bytes(raw)
    reader = RtfFileReader()
    with pytest.raises(ValueError, match="Standard text RTF.*libreoffice"):
        if mode == "path":
            aw.Document(path)
        elif mode == "stream":
            aw.Document(BytesIO(raw))
        elif mode == "bytes":
            with pytest.warns(DeprecationWarning):
                aw.Document(data=raw)
        else:
            getattr(reader, "load_file" if mode == "reader_path" else "load_" + mode.removeprefix("reader_"))(
                path if mode == "reader_path" else BytesIO(raw) if mode == "reader_stream" else raw)


@pytest.mark.parametrize("tag", ["footnoteReference", "fldSimple", "ins", "del", "sdt", "oMath"])
def test_load_loss_is_recorded_even_when_warning_is_ignored(tag):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", aw.loading.DocumentLoadWarning)
        doc = aw.Document(BytesIO(package(f"<w:p><w:{tag}/></w:p>")))
    assert len(doc.diagnostics) == 1
    item = asdict(doc.diagnostics[0])
    assert item["code"] == "load.content_loss" and tag in item["message"]
    assert item["location"] and item["severity"] == "warning"


@pytest.mark.parametrize("kind", ["headerReference", "footerReference"])
@pytest.mark.parametrize("variant", ["different_sections", "first", "even"])
def test_header_footer_source_loss_is_reported(kind, variant):
    if variant == "different_sections":
        body = (f'<w:p><w:pPr><w:sectPr><w:{kind} w:type="default" r:id="r1"/></w:sectPr></w:pPr></w:p>'
                f'<w:sectPr><w:{kind} w:type="default" r:id="r2"/></w:sectPr>')
    else:
        body = f'<w:sectPr><w:{kind} w:type="{variant}" r:id="r1"/></w:sectPr>'
    with pytest.warns(aw.loading.DocumentLoadWarning, match="flattened"):
        doc = aw.Document(BytesIO(package(body)))
    assert doc.diagnostics[0].code == "load.header_footer_flattened"


def test_distinct_model_headers_and_unknown_nodes_warn_before_save():
    doc = document()
    doc.sections[0].headers_footers = [ldm.HeaderFooter(header_footer_type=0, children=[ldm.Paragraph(children=[ldm.Run(text="A")])])]
    doc.sections.append(ldm.Section(headers_footers=[ldm.HeaderFooter(header_footer_type=0, children=[ldm.Paragraph(children=[ldm.Run(text="B")])])]))
    with pytest.warns(DocxWriterLossyWarning, match="per-section"):
        doc.to_bytes(aw.SaveFormat.DOCX)
    assert doc.diagnostics[-1].code == "docx.header_footer_flattened"
    doc.sections[0].body.children.append(ldm.UnknownNode(_type="not_supported"))
    with pytest.warns(DocxWriterLossyWarning):
        doc.to_bytes(aw.SaveFormat.DOCX)
    assert any(item.code == "docx.unknown_node" for item in doc.diagnostics)


def test_complex_fields_are_retained_not_misreported():
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[ldm.Paragraph(children=[
        ldm.FieldStart(), ldm.Run(text=" DATE "), ldm.FieldSeparator(),
        ldm.Run(text="2026-01-01"), ldm.FieldEnd(),
    ])]))])
    raw = LdmDocxWriter().write_to_bytes(model)
    with ZipFile(BytesIO(raw)) as archive:
        xml = archive.read("word/document.xml").decode()
    assert "instrText" in xml and " DATE " in xml and "2026-01-01" in xml
    assert all(f'w:fldCharType="{part}"' in xml for part in ("begin", "separate", "end"))
    doc = aw.Document(BytesIO(raw))
    assert not doc.diagnostics
    assert "instrText" in ZipFile(BytesIO(doc.to_bytes("docx"))).read("word/document.xml").decode()


@pytest.mark.parametrize("fmt", [aw.SaveFormat.DOCX, aw.SaveFormat.PDF, aw.SaveFormat.TEXT, aw.SaveFormat.MARKDOWN])
def test_memory_output_matches_path_content(tmp_path, fmt):
    doc = document()
    raw = doc.to_bytes(fmt)
    output = tmp_path / f"report.{fmt.name.lower()}"
    doc.save(output, fmt)
    if fmt == aw.SaveFormat.DOCX:
        with ZipFile(BytesIO(raw)) as memory, ZipFile(output) as disk:
            assert {n: memory.read(n) for n in memory.namelist()} == {n: disk.read(n) for n in disk.namelist()}
        assert "中文" in aw.Document(BytesIO(raw)).get_text()
    elif fmt == aw.SaveFormat.PDF:
        assert "中文" in "".join(p.extract_text() for p in PdfReader(BytesIO(raw)).pages)
        assert "".join(p.extract_text() for p in PdfReader(BytesIO(raw)).pages) == "".join(p.extract_text() for p in PdfReader(output).pages)
    else:
        assert raw == output.read_bytes()
    assert sorted(p.name for p in tmp_path.iterdir()) == [output.name]


def test_memory_options_and_validation(tmp_path):
    opts = aw.saving.MarkdownSaveOptions()
    opts.encoding = "utf-8-sig"
    assert document().to_bytes(opts).startswith(b"\xef\xbb\xbf")
    opts.images_folder = str(tmp_path / "images")
    with pytest.raises(ValueError, match="require save"):
        document().to_bytes(opts)
    assert not list(tmp_path.iterdir())
    for fmt in (None, "unknown", aw.SaveFormat.DOC):
        with pytest.raises(ValueError, match="Unsupported"):
            document().to_bytes(fmt)


def test_error_filtered_warning_is_recorded_and_output_preserved(tmp_path):
    doc = document("missing " + chr(0x1FAE0))
    output = tmp_path / "report.pdf"
    output.write_bytes(b"old")
    with warnings.catch_warnings():
        warnings.simplefilter("error", PdfMissingGlyphWarning)
        with pytest.raises(PdfMissingGlyphWarning):
            doc.save(output)
    assert output.read_bytes() == b"old"
    assert doc.diagnostics[-1].code == "pdf.missing_glyph"


def test_diagnostic_collection_is_context_local():
    def collect(number):
        target = []
        with collect_diagnostics(target), warnings.catch_warnings():
            warnings.simplefilter("ignore", aw.loading.DocumentLoadWarning)
            warn(str(number), aw.loading.DocumentLoadWarning)
        return [item.message for item in target]
    with ThreadPoolExecutor(max_workers=4) as pool:
        assert list(pool.map(collect, range(12))) == [[str(n)] for n in range(12)]


def test_structure_export_marks_links_and_excludes_hidden_text():
    doc = document("[公开链接](https://example.com)")
    doc.sections[0].body.children[0]._children.append(ldm.Run(text="hidden", font=ldm.Font(hidden=True)))
    block = doc.to_dict()["blocks"][0]
    assert block["text"] == "公开链接"
    assert block["runs"][0]["link"] == "https://example.com"
    assert "hidden" not in json.dumps(block)


def test_structure_export_keeps_body_order_and_locations():
    doc = document("# 标题\n\n正文\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\n结尾")
    result = json.loads(json.dumps(doc.to_dict(), ensure_ascii=False))
    assert result["schema_version"] == 1
    assert [block["type"] for block in result["blocks"]] == ["paragraph", "paragraph", "paragraph", "table", "paragraph"]
    assert result["blocks"][0]["heading_level"] == 1
    assert result["blocks"][3]["rows"][1]["cells"][0]["paragraphs"][0]["text"] == "1"
    assert result["blocks"][-1]["location"] == "sections[0].body.children[4]"
