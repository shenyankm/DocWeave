"""Save options must change the actual PDF streams and catalog."""

from io import BytesIO
import warnings

from pypdf import PdfReader
import pytest

import aspose.words_foss as aw
from aspose.words_foss.pdf_writer import PdfUnsupportedOptionWarning
from aspose.words_foss.saving import PdfPageMode, PdfSaveOptions, PdfTextCompression


def document():
    options = aw.LoadOptions()
    options.load_format = aw.LoadFormat.TEXT
    return aw.Document(BytesIO(("中文 PDF compression test. " * 250).encode()), options)


def test_compression_changes_streams_without_changing_content():
    outputs = []
    readers = []
    for compression in PdfTextCompression:
        options = PdfSaveOptions()
        options.text_compression = compression
        with warnings.catch_warnings(record=True) as caught:
            output = document().to_bytes(options)
        assert not any(issubclass(item.category, PdfUnsupportedOptionWarning) for item in caught)
        outputs.append(output)
        readers.append(PdfReader(BytesIO(output)))
        for page in readers[-1].pages:
            stream = page["/Contents"].get_object()
            assert stream.get("/Filter") == ("/FlateDecode" if compression == PdfTextCompression.FLATE else None)
    assert len(outputs[1]) < len(outputs[0])
    assert len(readers[0].pages) == len(readers[1].pages)
    for plain, compressed in zip(readers[0].pages, readers[1].pages):
        assert plain["/Contents"].get_object().get_data() == compressed["/Contents"].get_object().get_data()
        assert plain.extract_text() == compressed.extract_text()


@pytest.mark.parametrize("mode,expected", [
    (PdfPageMode.USE_NONE, "/UseNone"),
    (PdfPageMode.USE_OUTLINES, "/UseOutlines"),
    (PdfPageMode.USE_THUMBS, "/UseThumbs"),
    (PdfPageMode.FULL_SCREEN, "/FullScreen"),
    (PdfPageMode.USE_OC, "/UseOC"),
    (PdfPageMode.USE_ATTACHMENTS, "/UseAttachments"),
    ("use_thumbs", "/UseThumbs"),
    (3, "/FullScreen"),
])
def test_page_mode_is_written_to_catalog(mode, expected):
    options = PdfSaveOptions()
    options.page_mode = mode
    with warnings.catch_warnings(record=True) as caught:
        output = document().to_bytes(options)
    assert not any(issubclass(item.category, PdfUnsupportedOptionWarning) for item in caught)
    assert PdfReader(BytesIO(output)).trailer["/Root"]["/PageMode"] == expected


def test_default_stream_and_viewer_options():
    reader = PdfReader(BytesIO(document().to_bytes(PdfSaveOptions())))
    assert reader.trailer["/Root"]["/PageMode"] == "/UseOutlines"
    assert "/Outlines" not in reader.trailer["/Root"]
    assert reader.pages[0]["/Contents"].get_object()["/Filter"] == "/FlateDecode"


@pytest.mark.parametrize("option", ["text_compression", "page_mode"])
@pytest.mark.parametrize("value", [99, "invalid"])
def test_invalid_option_preserves_existing_output(tmp_path, option, value):
    target = tmp_path / "existing.pdf"
    target.write_bytes(b"existing output")
    options = PdfSaveOptions()
    setattr(options, option, value)
    with pytest.raises(ValueError):
        document().save(target, options)
    assert target.read_bytes() == b"existing output"
