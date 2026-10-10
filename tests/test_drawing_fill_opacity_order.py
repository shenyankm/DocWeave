"""Generated ordered opacity transforms and save/read/render regression cases."""
from io import BytesIO
from xml.etree import ElementTree as ET

import pytest

from aspose.words_foss import Document, SaveFormat, light_document_model as ldm
from aspose.words_foss._drawing_fill import resolve_drawing_fill
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer import LdmPdfWriter, PdfContentLossWarning
from tests.test_drawing_fill_sources import A, owned_rectangle, shape
from tests.test_drawing_fill_render import ramp


def ordered_gradient(transforms):
    root = ET.fromstring(f'<root xmlns:a="{A}">{ramp()}</root>')[0]
    for stop in root.find(f'{{{A}}}gsLst'):
        colour = stop[0]
        colour.clear()
        colour.set('val', 'FF0000')
        for tag, value in transforms:
            ET.SubElement(colour, f'{{{A}}}{tag}', val=str(value))
    return ET.tostring(root, encoding='unicode')


CASES = [
    ([('alpha', 30000)], 76),
    ([('alpha', 70000)], 178),
    ([('alpha', 25000), ('alphaMod', 50000), ('alphaOff', 10000)], 58),
    ([('alpha', 25000), ('alphaOff', 10000), ('alphaMod', 50000)], 45),
    ([('alpha', 50000), ('alphaOff', 10000), ('alphaMod', 50000)], 77),
    ([('alpha', 50000), ('alphaMod', 50000), ('alphaOff', 10000)], 90),
    ([('alpha', 0), ('alphaOff', 50000), ('alphaMod', 25000)], 32),
    ([('alpha', 100000), ('alphaMod', 200000), ('alphaOff', -50000)], 127),
]


@pytest.mark.parametrize('transforms, byte', CASES)
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_ordered_opacity_survives_json_save_and_actual_pdf(transforms, byte, format, tmp_path):
    model = Document(BytesIO(owned_rectangle(ordered_gradient(transforms), False))).light_document_model
    source = shape(model).source_drawing_fill
    model = ldm.Document.model_validate_json(model.model_dump_json())
    output = tmp_path / 'ordered.docx'
    Document(BytesIO(LdmDocxWriter().write_to_bytes(model))).save(output, format)
    cold = Document(output).light_document_model
    assert shape(cold).source_drawing_fill == source
    solid, gradient = resolve_drawing_fill(source, None)
    assert solid is None and all(colour[3] == byte / 255 for colour in gradient[2])
    before = cold.model_dump_json()
    pymupdf = pytest.importorskip('pymupdf')
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(cold), filetype='pdf') as pdf:
        actual = pdf[0].get_pixmap(matrix=pymupdf.Matrix(2, 2)).pixel(
            round(80 * 144 / 25.4), round(70 * 144 / 25.4))
        assert actual == pytest.approx((255, 255 - byte, 255 - byte), abs=2, rel=0)
    assert cold.model_dump_json() == before


@pytest.mark.parametrize('transforms', [
    [('alpha', -1)], [('alpha', 100001)], [('alphaMod', -1)], [('alphaOff', -100001)],
])
def test_invalid_opacity_range_is_diagnosed_and_source_retained(transforms):
    model = Document(BytesIO(owned_rectangle(ordered_gradient(transforms), False))).light_document_model
    before = model.model_dump_json()
    with pytest.warns(PdfContentLossWarning, match='outside supported range'):
        assert LdmPdfWriter().write_to_bytes(model).startswith(b'%PDF')
    assert model.model_dump_json() == before
