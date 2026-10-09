"""Viewer zoom percentages must not create invalid PDF destinations."""

from io import BytesIO

import pytest
from docx import Document as NativeDocument
from pypdf import PdfReader

import aspose.words_foss as aw
from aspose.words_foss.pdf_writer import LdmPdfWriter

TEXT = "LITERAL /OpenAction [999 0 R /Fit] END"


def document():
    return aw.Document(BytesIO(TEXT.encode()), aw.MarkdownLoadOptions())


@pytest.mark.parametrize("value", [-1, True, False, None, "fullpage", "default", "100", 1j,
                                   float("nan"), float("inf"), float("-inf"), 10**400, 10**100, 1e-100])
@pytest.mark.parametrize("entry", ["memory", "path", "writer"])
def test_invalid_active_zoom_is_rejected_before_output(tmp_path, value, entry):
    doc = document()
    options = aw.saving.PdfSaveOptions()
    options.zoom_behavior = aw.saving.PdfZoomBehavior.ZOOM_FACTOR
    writer = LdmPdfWriter(options)
    options.zoom_factor = value
    target = tmp_path / "original.pdf"
    target.write_bytes(b"ORIGINAL")
    with pytest.raises(ValueError, match="zoom_factor"):
        if entry == "memory":
            doc.to_bytes(options)
        elif entry == "path":
            doc.save(target, options)
        else:
            writer.write_to_bytes(doc.light_document_model)
    assert target.read_bytes() == b"ORIGINAL"
    assert list(tmp_path.iterdir()) == [target]


@pytest.mark.parametrize("value", [1, 100, 125.5, 1000, 2e-9, 1e30, 1e48, 1e-43])
def test_valid_percentages_are_written_to_first_page_destination(value):
    options = aw.saving.PdfSaveOptions()
    options.zoom_behavior = aw.saving.PdfZoomBehavior.ZOOM_FACTOR
    options.zoom_factor = value
    reader = PdfReader(BytesIO(document().to_bytes(options)), strict=True)
    action = reader.trailer["/Root"]["/OpenAction"]
    assert action[0].idnum == reader.pages[0].indirect_reference.idnum
    assert action[1] == "/XYZ" and float(action[-1]) == pytest.approx(value / 100, rel=1e-12, abs=0)


@pytest.mark.parametrize("compression", list(aw.saving.PdfTextCompression))
@pytest.mark.parametrize("mode,expected", [
    (aw.saving.PdfZoomBehavior.NONE, None),
    (aw.saving.PdfZoomBehavior.ZOOM_FACTOR, "/FitH"),
    (aw.saving.PdfZoomBehavior.FIT_PAGE, "/Fit"),
    (aw.saving.PdfZoomBehavior.FIT_WIDTH, "/FitH"),
    (aw.saving.PdfZoomBehavior.FIT_HEIGHT, "/FitV"),
    (aw.saving.PdfZoomBehavior.FIT_BOX, "/FitB"),
])
def test_all_modes_keep_content_and_valid_catalog(mode, expected, compression):
    options = aw.saving.PdfSaveOptions()
    options.zoom_behavior = mode
    options.text_compression = compression
    reader = PdfReader(BytesIO(document().to_bytes(options)), strict=True)
    action = reader.trailer["/Root"].get("/OpenAction")
    if expected is None:
        assert action is None
    else:
        assert action[1] == expected
        assert action[0].idnum == reader.pages[0].indirect_reference.idnum
    assert reader.pages[0].extract_text().strip() == TEXT


def test_inactive_zoom_factor_is_not_interpreted():
    options = aw.saving.PdfSaveOptions()
    options.zoom_factor = "fullpage"
    reader = PdfReader(BytesIO(document().to_bytes(options)), strict=True)
    assert "/OpenAction" not in reader.trailer["/Root"]


def test_real_docx_zoom_preserves_content_and_outline():
    native = NativeDocument()
    native.add_heading("Zoom heading", level=1)
    native.add_paragraph(TEXT)
    source = BytesIO()
    native.save(source)
    source.seek(0)
    options = aw.saving.PdfSaveOptions()
    options.zoom_behavior = aw.saving.PdfZoomBehavior.ZOOM_FACTOR
    options.zoom_factor = 125.5
    options.outline_options.headings_outline_levels = 1
    reader = PdfReader(BytesIO(aw.Document(source).to_bytes(options)), strict=True)
    assert float(reader.trailer["/Root"]["/OpenAction"][-1]) == 1.255
    assert reader.outline[0]["/Title"] == "Zoom heading"
    assert TEXT in reader.pages[0].extract_text()
