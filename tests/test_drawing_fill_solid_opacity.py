"""Owned transparent solid fills with and without DrawingML style references."""
from io import BytesIO
from xml.etree import ElementTree as ET

import pytest

from aspose.words_foss import Document, SaveFormat, light_document_model as ldm
from aspose.words_foss._drawing_fill import resolve_drawing_fill
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer import LdmPdfWriter, PdfContentLossWarning
from tests.test_drawing_fill_sources import A, owned_rectangle, shape
from tests.test_drawing_fill_opacity_order import CASES, ordered_gradient


def solid_declaration(transforms):
    gradient = ET.fromstring(ordered_gradient(transforms))
    fill = ET.Element(f'{{{A}}}solidFill')
    fill.append(gradient.find(f'{{{A}}}gsLst')[0][0])
    return ET.tostring(fill, encoding='unicode')


@pytest.mark.parametrize('transforms, byte', CASES)
@pytest.mark.parametrize('style', [False, True])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_solid_opacity_keeps_declaration_and_actual_paint(transforms, byte, style, format, tmp_path):
    model = Document(BytesIO(owned_rectangle(solid_declaration(transforms), style))).light_document_model
    source = shape(model).source_drawing_fill
    model = ldm.Document.model_validate_json(model.model_dump_json())
    output = tmp_path / 'solid.docx'
    Document(BytesIO(LdmDocxWriter().write_to_bytes(model))).save(output, format)
    cold = Document(output).light_document_model
    assert shape(cold).source_drawing_fill == source
    solid, gradient = resolve_drawing_fill(source, None)
    assert solid == (255, 0, 0, byte / 255) and gradient is None
    before = cold.model_dump_json()
    pymupdf = pytest.importorskip('pymupdf')
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(cold), filetype='pdf') as rendered:
        pixel = rendered[0].get_pixmap(matrix=pymupdf.Matrix(2, 2)).pixel(
            round(80 * 144 / 25.4), round(70 * 144 / 25.4))
        assert pixel == pytest.approx((255, 255 - byte, 255 - byte), abs=2, rel=0)
    assert cold.model_dump_json() == before


@pytest.mark.parametrize('rotation', [0, 30])
def test_solid_opacity_does_not_change_stroke_or_following_content(rotation):
    from types import SimpleNamespace
    from fpdf import FPDF
    from aspose.words_foss.pdf_writer.shape_renderer import ShapeRenderer
    model = Document(BytesIO(owned_rectangle(solid_declaration([('alpha', 0)]), False))).light_document_model
    rectangle = shape(model)
    rectangle.source_drawing_fill = rectangle.source_drawing_fill.model_copy(update={'rotation': rotation * 60000})
    rectangle.stroke = ldm.Border(color='#0000FF', line_width=2)
    pdf = FPDF(format='Letter')
    pdf.add_page()
    ShapeRenderer(SimpleNamespace(_doc=model, _page_width=215.9, _page_height=279.4)).render_positioned_shape(pdf, rectangle)
    pdf.set_fill_color(0, 0, 0)
    pdf.rect(10, 10, 5, 5, 'F')
    pdf.set_font('Helvetica', size=10)
    pdf.text(10, 20, 'AFTER')
    pymupdf = pytest.importorskip('pymupdf')
    with pymupdf.open(stream=bytes(pdf.output()), filetype='pdf') as rendered:
        page = rendered[0]
        borders = [draw for draw in page.get_drawings() if draw['color'] == (0, 0, 1)]
        assert borders and all(draw['stroke_opacity'] == 1 for draw in borders)
        assert page.get_pixmap(matrix=pymupdf.Matrix(2, 2)).pixel(
            round(12 * 144 / 25.4), round(12 * 144 / 25.4)) == (0, 0, 0)
        assert page.get_texttrace()[-1]['opacity'] == 1 and page.get_texttrace()[-1]['dir'] == (1, 0)


def test_unsupported_solid_transform_warns_without_changing_source():
    root = ET.fromstring(solid_declaration([]))
    ET.SubElement(root[0], f'{{{A}}}gamma')
    model = Document(BytesIO(owned_rectangle(ET.tostring(root, encoding='unicode'), False))).light_document_model
    before = model.model_dump_json()
    with pytest.warns(PdfContentLossWarning, match='colour transform.*gamma'):
        assert LdmPdfWriter().write_to_bytes(model).startswith(b'%PDF')
    assert model.model_dump_json() == before
