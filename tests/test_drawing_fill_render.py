"""Generated rectangle paints and explicit unsupported DrawingML boundaries."""
from io import BytesIO
from xml.etree import ElementTree as ET

import pytest
from pypdf import PdfReader

from aspose.words_foss import Document, light_document_model as ldm
from aspose.words_foss._drawing_fill import resolve_drawing_fill
from aspose.words_foss._theme import default_theme_bytes
from aspose.words_foss.pdf_writer import LdmPdfWriter, PdfContentLossWarning
from aspose.words_foss.pdf_writer.constants import PT_TO_MM
from tests.test_drawing_fill_sources import A, owned_rectangle, shape


def ramp(alphas=(100000,) * 3, *, follow=True, colours=('FF0000',) * 3):
    stops = ''.join(f'<a:gs pos="{pos}"><a:srgbClr val="{colour}"><a:alpha val="{alpha}"/></a:srgbClr></a:gs>'
                    for pos, alpha, colour in zip((0, 50000, 100000), alphas, colours))
    return f'<a:gradFill rotWithShape="{int(follow)}"><a:gsLst>{stops}</a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>'


@pytest.mark.parametrize('rotation', [0, 30, 90])
@pytest.mark.parametrize('follow', [False, True])
@pytest.mark.parametrize('alphas', [(50000,) * 3, (0, 50000, 100000)])
def test_transparent_rectangle_keeps_source_and_graphics_state(rotation, follow, alphas):
    model = Document(BytesIO(owned_rectangle(ramp(alphas, follow=follow), False))).light_document_model
    rectangle = shape(model)
    rectangle.source_drawing_fill = rectangle.source_drawing_fill.model_copy(update={
        'rotation': rotation * 60000, 'flip_horizontal': False})
    before = model.model_dump_json()
    data = LdmPdfWriter().write_to_bytes(model)
    pdf = PdfReader(BytesIO(data))
    assert any('/SMask' in item.get_object() for item in pdf.pages[0]['/Resources']['/ExtGState'].values())
    assert model.model_dump_json() == before
    pymupdf = pytest.importorskip('pymupdf')
    with pymupdf.open(stream=data, filetype='pdf') as rendered:
        pixel = rendered[0].get_pixmap(matrix=pymupdf.Matrix(2, 2)).pixel(
            round(80 * 144 / 25.4), round(70 * 144 / 25.4))
        assert pixel == pytest.approx((255, 127, 127), abs=2, rel=0)
    from types import SimpleNamespace
    from fpdf import FPDF
    from aspose.words_foss.pdf_writer.shape_renderer import ShapeRenderer
    pdf = FPDF(format='Letter')
    pdf.add_page()
    rectangle.stroke = ldm.Border(color='#0000FF', line_width=2)
    writer = SimpleNamespace(_doc=model, _page_width=215.9, _page_height=279.4)
    ShapeRenderer(writer).render_positioned_shape(pdf, rectangle)
    pdf.set_fill_color(0, 0, 0)
    pdf.rect(10, 10, 5, 5, 'F')
    pdf.set_font('Helvetica', size=10)
    pdf.text(10, 20, 'AFTER')
    with pymupdf.open(stream=bytes(pdf.output()), filetype='pdf') as rendered:
        page = rendered[0]
        assert page.get_pixmap(matrix=pymupdf.Matrix(2, 2)).pixel(
            round(12 * 144 / 25.4), round(12 * 144 / 25.4)) == (0, 0, 0)
        assert all(draw['stroke_opacity'] == 1 for draw in page.get_drawings()
                   if draw['color'] == (0, 0, 1))
        assert page.get_texttrace()[-1]['opacity'] == 1
        assert page.get_texttrace()[-1]['dir'] == (1, 0)


@pytest.mark.parametrize('height', [792, 841.9, 888.8])
@pytest.mark.parametrize('alphas', [(100000,) * 3, (0, 50000, 100000)])
def test_pattern_page_height_uses_exact_point_conversion(height, alphas):
    model = Document(BytesIO(owned_rectangle(ramp(alphas, colours=('000000', '808080', 'FFFFFF')), False))).light_document_model
    rectangle = shape(model)
    rectangle.source_drawing_fill = rectangle.source_drawing_fill.model_copy(update={'rotation': 0})
    rectangle.drawing_position_mm = (60, 170 * PT_TO_MM)
    model = ldm.Document.model_validate_json(model.model_dump_json())
    model.sections[0].page_setup.page_height = height
    pdf = PdfReader(BytesIO(LdmPdfWriter().write_to_bytes(model)))
    assert PT_TO_MM == 25.4 / 72
    for item in pdf.pages[0]['/Resources']['/Pattern'].values():
        matrix = item.get_object()['/Matrix']
        assert float(matrix[5]) + (170 if matrix[4] else 0) == pytest.approx(height, abs=1e-5, rel=0)


@pytest.mark.parametrize('direct', [
    '<a:gradFill><a:gsLst/><a:path path="circle"/></a:gradFill>',
    ramp((0, 50000, 0)),
    '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>',
])
def test_unsupported_paint_warns_and_keeps_original_declaration(direct):
    model = Document(BytesIO(owned_rectangle(direct, False))).light_document_model
    before = model.model_dump_json()
    with pytest.warns(PdfContentLossWarning, match='DrawingML'):
        assert LdmPdfWriter().write_to_bytes(model).startswith(b'%PDF')
    assert model.model_dump_json() == before


def test_missing_theme_projection_uses_complete_licensed_template():
    root = ET.fromstring(default_theme_bytes())
    elements = root.find(f'{{{A}}}themeElements')
    assert [child.tag for child in elements] == [f'{{{A}}}{name}' for name in ('clrScheme', 'fontScheme', 'fmtScheme')]
    model = Document(BytesIO(owned_rectangle('', True))).light_document_model
    assert model.source_theme is None
    source = shape(model).source_drawing_fill
    source = source.model_copy(update={'style_xml': source.style_xml.replace('idx="3"', 'idx="1"')})
    assert resolve_drawing_fill(source, None)[0] == (79, 129, 189)
    assert model.source_theme is None
