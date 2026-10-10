"""Owned, generated DrawingML declarations; no external baseline files."""
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from aspose.words_foss import Document, SaveFormat, light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer import LdmPdfWriter, PdfContentLossWarning

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'


def owned_rectangle(direct, style=True):
    reference = ('<wps:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>'
                 '<a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>'
                 '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
                 '<a:fontRef idx="minor"><a:schemeClr val="accent1"/></a:fontRef></wps:style>'
                 if style else '')
    parts = {
        '[Content_Types].xml': f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        '_rels/.rels': f'<Relationships xmlns="{R}"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        'word/document.xml': f'''<w:document xmlns:w="{W}" xmlns:a="{A}" xmlns:wp="{WP}" xmlns:wps="{WPS}"><w:body><w:p><w:r><w:drawing>
          <wp:anchor simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="page"><wp:posOffset>2160000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>2160000</wp:posOffset></wp:positionV>
            <wp:extent cx="1440000" cy="720000"/><wp:wrapNone/><wp:docPr id="1" name="Owned"/>
            <a:graphic><a:graphicData uri="{WPS}"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm rot="1800000" flipH="1"><a:off x="0" y="0"/><a:ext cx="1440000" cy="720000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{direct}<a:ln><a:noFill/></a:ln></wps:spPr>{reference}<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>
          </wp:anchor></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>''',
    }
    result = BytesIO()
    with ZipFile(result, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return result.getvalue()


FILLS = [
    '<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>',
    '<a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>',
    '<a:noFill/>',
]


def shape(model):
    nodes = model.get_child_nodes(ldm.NodeType.SHAPE, True)
    assert len(nodes) == 1
    return nodes[0]


@pytest.mark.parametrize('direct', FILLS)
@pytest.mark.parametrize('style', [False, True])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_fill_sources_survive_edit_json_and_package_save(direct, style, format, tmp_path):
    model = Document(BytesIO(owned_rectangle(direct, style))).light_document_model
    original = shape(model).source_drawing_fill
    assert bool(original.style_xml) is style
    assert ET.fromstring(original.direct_xml).tag == ET.fromstring(f'<root xmlns:a="{A}">{direct}</root>')[0].tag
    assert original.rotation == 1800000 and original.flip_horizontal
    updated = original.model_copy(update={'flip_vertical': True})
    shape(model).source_drawing_fill = updated
    assert not original.flip_vertical
    restored = ldm.Document.model_validate_json(model.model_dump_json())
    assert shape(restored)._is_positioned and shape(restored).drawing_position_mm == (60, 60)
    output = tmp_path / 'rectangle.docx'
    Document(BytesIO(LdmDocxWriter().write_to_bytes(restored))).save(output, format)
    cold = Document(output).light_document_model
    assert shape(cold).source_drawing_fill == updated
    assert (shape(cold).width, shape(cold).height) == (40, 20)


def test_reference_only_rectangle_is_preserved_and_pdf_loss_is_explicit():
    model = Document(BytesIO(owned_rectangle('', True))).light_document_model
    assert shape(model).source_drawing_fill.direct_xml == ''
    with pytest.warns(PdfContentLossWarning, match='DrawingML source fills'):
        assert LdmPdfWriter().write_to_bytes(model).startswith(b'%PDF')


def test_direct_colour_edit_replaces_retained_declaration():
    model = Document(BytesIO(owned_rectangle(FILLS[0], False))).light_document_model
    shape(model).fill_color = '#00FF00'
    assert shape(model).source_drawing_fill is None
    cold = Document(BytesIO(LdmDocxWriter().write_to_bytes(model))).light_document_model
    assert shape(cold).source_drawing_fill is not None
    assert '00FF00' in shape(cold).source_drawing_fill.direct_xml


@pytest.mark.parametrize('xml', ['<wrong/>', '<a:blipFill xmlns:a="' + A + '"/>', '<!DOCTYPE x [<!ENTITY a "x">]><x>&a;</x>'])
def test_invalid_source_xml_is_rejected(xml):
    with pytest.raises((ValueError, ET.ParseError)):
        ldm.SourceDrawingFill(direct_xml=xml)
