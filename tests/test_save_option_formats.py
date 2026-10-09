"""Save-option format requests must agree with the selected writer."""

from io import BytesIO
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss.docx_writer import LdmDocxWriter


def document():
    return aw.Document(BytesIO(b"FORMAT_SENTINEL"), aw.MarkdownLoadOptions())


@pytest.mark.parametrize("kind", ["docx", "markdown"])
@pytest.mark.parametrize("value", ["pdf", aw.SaveFormat.TEXT, 999, None, ".docx"])
@pytest.mark.parametrize("path_output", [False, True])
def test_mismatched_option_format_rejected_before_output(tmp_path, kind, value, path_output):
    options = aw.saving.OoxmlSaveOptions() if kind == "docx" else aw.saving.MarkdownSaveOptions()
    options.save_format = value
    doc = document()
    output = tmp_path / "existing.output"
    output.write_bytes(b"ORIGINAL")
    with pytest.raises(ValueError, match="save_format"):
        if path_output:
            doc.save(output, options)
        else:
            doc.to_bytes(options)
    assert output.read_bytes() == b"ORIGINAL"
    assert list(tmp_path.iterdir()) == [output]


@pytest.mark.parametrize("value", ["pdf", aw.SaveFormat.TEXT, 999])
@pytest.mark.parametrize("path_output", [False, True])
def test_direct_docx_writer_validates_mutated_options_and_recovers(tmp_path, value, path_output):
    options = aw.saving.OoxmlSaveOptions()
    writer = LdmDocxWriter(options)
    options.save_format = value
    output = tmp_path / "existing.docx"
    output.write_bytes(b"ORIGINAL")
    with pytest.raises(ValueError, match="save_format"):
        if path_output:
            writer.write(document().light_document_model, output)
        else:
            writer.write_to_bytes(document().light_document_model)
    assert output.read_bytes() == b"ORIGINAL"
    options.save_format = aw.SaveFormat.DOCX
    assert aw.Document(BytesIO(writer.write_to_bytes(document().light_document_model))).get_text().strip() == "FORMAT_SENTINEL"


def test_docx_constructor_none_keeps_default_format():
    raw = document().to_bytes(aw.saving.OoxmlSaveOptions(None))
    assert aw.Document(BytesIO(raw)).get_text().strip() == "FORMAT_SENTINEL"


@pytest.mark.parametrize("value", ["pdf", aw.SaveFormat.TEXT, 999])
def test_docx_constructor_format_is_enforced(value):
    with pytest.raises(ValueError, match="save_format"):
        document().to_bytes(aw.saving.OoxmlSaveOptions(value))


@pytest.mark.parametrize("kind,values", [
    ("docx", [aw.SaveFormat.DOCX, int(aw.SaveFormat.DOCX), "docx", "DOCX"]),
    ("markdown", [aw.SaveFormat.MARKDOWN, int(aw.SaveFormat.MARKDOWN), "markdown", "MD"]),
])
def test_supported_option_formats_keep_content(kind, values):
    doc = document()
    for value in values:
        options = aw.saving.OoxmlSaveOptions(value) if kind == "docx" else aw.saving.MarkdownSaveOptions()
        options.save_format = value
        raw = doc.to_bytes(options)
        if kind == "docx":
            with ZipFile(BytesIO(raw)) as archive:
                assert b"FORMAT_SENTINEL" in archive.read("word/document.xml")
        else:
            assert raw.decode().strip() == "FORMAT_SENTINEL"
