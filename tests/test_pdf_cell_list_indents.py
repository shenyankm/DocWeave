"""Hanging indents must affect cell-list measurement and actual placement."""

from io import BytesIO

import pymupdf
from pypdf import PdfReader
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.saving import PdfSaveOptions
from tests.test_pdf_cell_lists import cell_list_model
from tests.test_pdf_structure_pages import assert_page_tags


def indent_model(long=False, explicit=True, alignment=0):
    doc = cell_list_model(explicit=explicit)
    cell = doc.sections[0].body.children[0].rows[0].cells[0]
    for index, para in enumerate(cell.paragraphs[:3]):
        para.paragraph_format.left_indent = 24 + 12 * para.list_format.list_level_number
        para.paragraph_format.first_line_indent = -12
        para.paragraph_format.right_indent = 6
        para.paragraph_format.alignment = alignment
        text = f"BODY{index} " + (
            " ".join(f"WORD{index}_{i:02}" for i in range(30))
            if not long
            else "\n".join(f"LINE{index}_{i:02}" for i in range(32))
        )
        para._children = [ldm.Run(text=text, font=ldm.Font(size=9))]
    return doc


@pytest.mark.parametrize("long", [False, True])
@pytest.mark.parametrize("shaping", [False, True])
def test_explicit_hanging_indent_has_aligned_continuations(long, shaping):
    doc = indent_model(long)
    snapshot = doc.model_dump()
    options = PdfSaveOptions()
    options.export_document_structure = True
    options.text_shaping = shaping
    raw = LdmPdfWriter(options).write_to_bytes(doc)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        words = [
            (page.number, *word) for page in pdf for word in page.get_text("words")
        ]
        starts = []
        for index in range(3):
            body = next(word for word in words if word[5] == f"BODY{index}")
            marker = next(word for word in words if word[5] == f"{index + 1}.")
            # Actual positions are in pt: the source hanging offset is 12 pt.
            assert body[1] - marker[1] == pytest.approx(12, abs=0.05)
            continuations = [
                word
                for word in words
                if word[5].startswith(f"LINE{index}_")
                or word[5].startswith(f"WORD{index}_")
            ]
            row_starts = {}
            for word in continuations:
                row_starts.setdefault((word[0], round(word[2], 1)), word[1])
            for (page_number, y), x in row_starts.items():
                if (page_number, y) != (body[0], round(body[2], 1)):
                    assert x == pytest.approx(body[1], abs=0.05)
            assert continuations and len(row_starts) > 1
            starts.append(body[1])
        assert starts[1] - starts[0] == pytest.approx(12, abs=0.05)
        assert starts[2] == pytest.approx(starts[0], abs=0.05)
    control_options = PdfSaveOptions()
    control_options.text_shaping = shaping
    with (
        pymupdf.open(stream=raw, filetype="pdf") as actual,
        pymupdf.open(
            stream=LdmPdfWriter(control_options).write_to_bytes(doc), filetype="pdf"
        ) as control,
    ):
        assert len(actual) == len(control)
        for page, reference in zip(actual, control, strict=True):
            assert page.get_text("words") == reference.get_text("words")
            assert page.get_pixmap().samples == reference.get_pixmap().samples
    if long:
        assert_page_tags(raw)
    assert doc.model_dump() == snapshot


@pytest.mark.parametrize("alignment", [1, 2])
def test_aligned_lists_keep_markers_separate_and_respect_right_indent(alignment):
    from aspose.words_foss.pdf_writer.constants import (
        DEFAULT_CELL_PAD_LEFT_MM,
        PT_TO_MM,
    )

    doc = indent_model(alignment=alignment)
    raw = LdmPdfWriter().write_to_bytes(doc)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        words = [word for page in pdf for word in page.get_text("words")]
        for index in range(3):
            body = next(word for word in words if word[4] == f"BODY{index}")
            marker = next(word for word in words if word[4] == f"{index + 1}.")
            assert body[0] - marker[0] == pytest.approx(12, abs=0.05)
            assert marker[2] < body[0]
        cell_words = [word for word in words if word[4].startswith(("BODY", "WORD"))]
        limit = 175 - DEFAULT_CELL_PAD_LEFT_MM / PT_TO_MM - 6
        assert all(word[2] <= limit + 0.05 for word in cell_words)


def test_default_cell_list_levels_have_visible_indentation():
    doc = cell_list_model()
    raw = LdmPdfWriter().write_to_bytes(doc)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        words = pdf[0].get_text("words")
        starts = [
            next(word[0] for word in words if word[4] == f"CELLITEM{i}")
            for i in range(3)
        ]
        assert starts[1] - starts[0] == pytest.approx(5 * 72 / 25.4, abs=0.05)
        assert starts[2] == pytest.approx(starts[0], abs=0.05)


def test_public_docx_hanging_indent_reaches_pdf():
    import aspose.words_foss as aw
    from aspose.words_foss.docx_writer import LdmDocxWriter

    doc = aw.Document(BytesIO(LdmDocxWriter().write_to_bytes(indent_model())))
    options = PdfSaveOptions()
    options.export_document_structure = True
    raw = doc.to_bytes(options)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        words = [word for page in pdf for word in page.get_text("words")]
        body = next(word for word in words if word[4] == "BODY0")
        marker = next(word for word in words if word[4] == "1.")
        nested = next(word for word in words if word[4] == "BODY1")
        assert body[0] - marker[0] == pytest.approx(12, abs=0.05)
        assert nested[0] - body[0] == pytest.approx(12, abs=0.05)
        text = "".join(page.get_text() for page in pdf)
        assert all(text.count(f"BODY{i}") == 1 for i in range(3))
    assert PdfReader(BytesIO(raw)).trailer["/Root"]["/StructTreeRoot"]


@pytest.mark.parametrize("shaping", [False, True])
def test_list_image_fits_indented_text_area(shaping):
    from PIL import Image
    from aspose.words_foss.pdf_writer.constants import (
        DEFAULT_CELL_PAD_LEFT_MM,
        PT_TO_MM,
    )

    doc = indent_model()
    paragraph = doc.sections[0].body.children[0].rows[0].cells[0].paragraphs[0]
    stream = BytesIO()
    Image.new("RGB", (300, 20), "blue").save(stream, format="PNG")
    paragraph._children.append(
        ldm.Shape(
            has_image=True,
            width=300,
            height=20,
            alternative_text="BLUE",
            image_data=ldm.ImageData(image_bytes=stream.getvalue()),
        )
    )
    options = PdfSaveOptions()
    options.text_shaping = shaping
    raw = LdmPdfWriter(options).write_to_bytes(doc)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        words = [word for page in pdf for word in page.get_text("words")]
        body = next(word for word in words if word[4] == "BODY0")
        images = [image for page in pdf for image in page.get_image_info()]
        assert len(images) == 1
        assert images[0]["bbox"][0] == pytest.approx(body[0], abs=0.05)
        assert (
            images[0]["bbox"][2] <= 175 - DEFAULT_CELL_PAD_LEFT_MM / PT_TO_MM - 6 + 0.05
        )


def test_cell_list_indents_without_text_area_preserve_output(tmp_path):
    doc = indent_model()
    para = doc.sections[0].body.children[0].rows[0].cells[0].paragraphs[0]
    para.paragraph_format.left_indent = 500
    para.paragraph_format.first_line_indent = -500
    output = tmp_path / "keep.pdf"
    output.write_bytes(b"KEEP")
    with pytest.raises(
        ValueError, match="No usable text width after cell list indents"
    ):
        LdmPdfWriter().write(doc, output)
    assert output.read_bytes() == b"KEEP"
