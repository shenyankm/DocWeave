"""Owned theme-reference colours retain their opacity when used as placeholders."""
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from aspose.words_foss import Document, SaveFormat, light_document_model as ldm
from aspose.words_foss._drawing_fill import resolve_drawing_fill
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss._theme import default_theme_bytes
from tests.test_drawing_fill_sources import A, WPS, R, CT, owned_rectangle, shape as gradient_shape
from tests.test_drawing_fill_solid_opacity import solid_declaration
from tests.test_drawing_fill_render import ramp


def placeholder_docx(reference_transforms, fill_transforms=(), *, placeholder=True, rotation=0):
    with ZipFile(BytesIO(owned_rectangle('', True))) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    root = ET.fromstring(parts['word/document.xml'])
    root.find(f'.//{{{A}}}xfrm').set('rot', str(rotation * 60000))
    ref = root.find(f'.//{{{WPS}}}style/{{{A}}}fillRef')
    ref.clear()
    ref.set('idx', '3')
    colour = ET.SubElement(ref, f'{{{A}}}srgbClr', val='FF0000')
    for tag, value in reference_transforms:
        ET.SubElement(colour, f'{{{A}}}{tag}', val=str(value))
    parts['word/document.xml'] = ET.tostring(root)
    theme = ET.fromstring(default_theme_bytes())
    fills = theme.find(f'{{{A}}}themeElements/{{{A}}}fmtScheme/{{{A}}}fillStyleLst')
    fills.remove(fills[2])
    solid = ET.fromstring(solid_declaration(fill_transforms))
    if placeholder:
        solid[0].tag = f'{{{A}}}schemeClr'
        solid[0].set('val', 'phClr')
    fills.append(solid)
    parts['word/theme/theme1.xml'] = ET.tostring(theme)
    parts['word/_rels/document.xml.rels'] = (f'<Relationships xmlns="{R}"><Relationship Id="theme" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        'Target="theme/theme1.xml"/></Relationships>').encode()
    types = ET.fromstring(parts['[Content_Types].xml'])
    ET.SubElement(types, f'{{{CT}}}Override', PartName='/word/theme/theme1.xml',
                  ContentType='application/vnd.openxmlformats-officedocument.theme+xml')
    parts['[Content_Types].xml'] = ET.tostring(types)
    output = BytesIO()
    with ZipFile(output, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return output.getvalue()


CASES = [
    ([('alpha', 0)], [], True, 0),
    ([('alpha', 50000)], [], True, 128),
    ([('alpha', 100000)], [], True, 255),
    ([('alpha', 50000)], [('alpha', 25000)], True, 64),
    ([('alpha', 50000)], [('alphaMod', 50000)], True, 64),
    ([('alpha', 50000)], [('alphaOff', 10000)], True, 154),
    ([('alpha', 50000)], [('alpha', 50000), ('alphaMod', 50000), ('alphaOff', 10000)], True, 90),
    ([('alpha', 50000), ('alphaOff', 10000), ('alphaMod', 50000)], [], True, 77),
    ([('alpha', 50000)], [], False, 255),
    ([('alpha', 0)], [], False, 255),
    ([('alphaMod', 50000)], [('alphaOff', -10000)], True, 102),
]


@pytest.mark.parametrize('reference,fill,placeholder,byte', CASES)
@pytest.mark.parametrize('rotation', [0, 30])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_placeholder_alpha_json_save_cold_and_actual_paint(reference, fill, placeholder, byte, rotation, format):
    import pymupdf

    owner = Document(BytesIO(placeholder_docx(reference, fill, placeholder=placeholder, rotation=rotation)))
    model = ldm.Document.model_validate_json(owner.light_document_model.model_dump_json())
    source = gradient_shape(model).source_drawing_fill
    rgb, gradient = resolve_drawing_fill(source, model.source_theme.data)
    assert gradient is None
    assert rgb == ((255, 0, 0) if byte == 255 else (255, 0, 0, byte / 255))
    owner = Document(BytesIO(LdmDocxWriter().write_to_bytes(model)))
    cold = Document(BytesIO(owner.to_bytes(format))).light_document_model
    assert gradient_shape(cold).source_drawing_fill == source
    before = cold.model_dump_json()
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(cold), filetype='pdf') as pdf:
        pixel = pdf[0].get_pixmap(matrix=pymupdf.Matrix(2, 2)).pixel(
            round(80 * 144 / 25.4), round(70 * 144 / 25.4))
    assert pixel == pytest.approx((255, 255 - byte, 255 - byte), abs=2)
    assert cold.model_dump_json() == before


def placeholder_gradient_docx(reference_transforms, stop_transform='alphaMod', *, rotation=0):
    with ZipFile(BytesIO(placeholder_docx(reference_transforms, rotation=rotation))) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    theme = ET.fromstring(parts['word/theme/theme1.xml'])
    fills = theme.find(f'{{{A}}}themeElements/{{{A}}}fmtScheme/{{{A}}}fillStyleLst')
    fills.remove(fills[2])
    gradient = ET.fromstring(f'<root xmlns:a="{A}">{ramp()}</root>')[0]
    values = [0, 50000, 100000] if stop_transform == 'alphaMod' else [-10000, 0, 10000]
    for stop, value in zip(gradient.find(f'{{{A}}}gsLst'), values):
        colour = stop[0]
        colour.clear()
        colour.tag = f'{{{A}}}schemeClr'
        colour.set('val', 'phClr')
        ET.SubElement(colour, f'{{{A}}}{stop_transform}', val=str(value))
    fills.append(gradient)
    parts['word/theme/theme1.xml'] = ET.tostring(theme)
    output = BytesIO()
    with ZipFile(output, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return output.getvalue()


@pytest.mark.parametrize('reference,base', [([('alpha', 50000)], 128),
    ([('alpha', 50000), ('alphaOff', 10000), ('alphaMod', 50000)], 77)])
@pytest.mark.parametrize('transform', ['alphaMod', 'alphaOff'])
@pytest.mark.parametrize('rotation', [0, 30])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_placeholder_gradient_inherits_each_stop_opacity(reference, base, transform, rotation, format):
    owner = Document(BytesIO(placeholder_gradient_docx(reference, transform, rotation=rotation)))
    model = ldm.Document.model_validate_json(owner.light_document_model.model_dump_json())
    source = gradient_shape(model).source_drawing_fill
    solid, gradient = resolve_drawing_fill(source, model.source_theme.data)
    assert solid is None
    expected = [0, round(base / 2), base] if transform == 'alphaMod' else [base - 26, base, base + 26]
    assert [round(colour[3] * 255) for colour in gradient[2]] == expected
    owner = Document(BytesIO(LdmDocxWriter().write_to_bytes(model)))
    cold = Document(BytesIO(owner.to_bytes(format))).light_document_model
    assert gradient_shape(cold).source_drawing_fill == source
    assert resolve_drawing_fill(source, cold.source_theme.data) == (solid, gradient)
    before = cold.model_dump_json()
    assert LdmPdfWriter().write_to_bytes(cold).startswith(b'%PDF-')
    assert cold.model_dump_json() == before
