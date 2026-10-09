"""Valid but unimplemented save requests must produce actionable diagnostics."""

import warnings
from io import BytesIO
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss.diagnostics import ConversionWarning, collect_diagnostics
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.md_writer import LdmMarkdownWriter
from aspose.words_foss.models import ConversionOptions


def requested_options(kind):
    if kind == "zip64":
        options = aw.saving.OoxmlSaveOptions()
        options.zip_64_mode = aw.saving.Zip64Mode.ALWAYS
    else:
        options = aw.saving.MarkdownSaveOptions()
        if kind == "html":
            options.export_as_html = aw.saving.MarkdownExportAsHtml.NON_COMPATIBLE_TABLES
        else:
            options.image_resolution = 300
    return options


@pytest.mark.parametrize("kind,option", [("zip64", "zip_64_mode"), ("html", "export_as_html"),
                                        ("resolution", "image_resolution")])
@pytest.mark.parametrize("path_output", [False, True])
def test_unsupported_request_is_recorded_even_when_warnings_ignored(tmp_path, kind, option, path_output):
    doc = aw.Document(BytesIO(b"CONTENT_SENTINEL"), aw.MarkdownLoadOptions())
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", ConversionWarning)
        if path_output:
            target = tmp_path / "result.output"
            doc.save(target, requested_options(kind))
            raw = target.read_bytes()
        else:
            raw = doc.to_bytes(requested_options(kind))
    assert raw
    assert len(doc.diagnostics) == 1
    assert doc.diagnostics[0].code == ("docx" if kind == "zip64" else "markdown") + ".unsupported_option"
    assert option in doc.diagnostics[0].message
    if kind == "zip64":
        assert aw.Document(BytesIO(raw)).get_text().strip() == "CONTENT_SENTINEL"
        with ZipFile(BytesIO(raw)) as archive:
            assert all(info.extract_version < 45 for info in archive.infolist())
    else:
        assert raw.decode().strip() == "CONTENT_SENTINEL"


@pytest.mark.parametrize("kind", ["zip64", "html", "resolution"])
def test_rejected_warning_keeps_original_output(tmp_path, kind):
    doc = aw.Document(BytesIO(b"CONTENT_SENTINEL"), aw.MarkdownLoadOptions())
    target = tmp_path / "original.output"
    target.write_bytes(b"ORIGINAL")
    with warnings.catch_warnings():
        warnings.simplefilter("error", ConversionWarning)
        with pytest.raises(ConversionWarning):
            doc.save(target, requested_options(kind))
    assert target.read_bytes() == b"ORIGINAL"
    assert list(tmp_path.iterdir()) == [target]


@pytest.mark.parametrize("kind", ["docx", "markdown"])
def test_default_options_do_not_warn(kind):
    doc = aw.Document(BytesIO(b"CONTENT_SENTINEL"), aw.MarkdownLoadOptions())
    options = aw.saving.OoxmlSaveOptions() if kind == "docx" else aw.saving.MarkdownSaveOptions()
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        doc.to_bytes(options)
    assert not caught and not doc.diagnostics


@pytest.mark.parametrize("kind", ["zip64", "html"])
def test_direct_writers_share_unsupported_request_diagnostics(kind):
    model = aw.Document(BytesIO(b"CONTENT_SENTINEL"), aw.MarkdownLoadOptions()).light_document_model
    diagnostics = []
    with collect_diagnostics(diagnostics), warnings.catch_warnings():
        warnings.simplefilter("ignore", ConversionWarning)
        if kind == "zip64":
            LdmDocxWriter(requested_options(kind)).write_to_bytes(model)
        else:
            LdmMarkdownWriter(ConversionOptions(
                export_as_html=aw.saving.MarkdownExportAsHtml.NON_COMPATIBLE_TABLES)).write(model)
    assert len(diagnostics) == 1 and diagnostics[0].code.endswith(".unsupported_option")
