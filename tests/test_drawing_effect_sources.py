"""Owned shape effects survive independently of fill edits; PDF reports loss."""
from io import BytesIO
from zipfile import ZipFile
import warnings

import pytest
from pydantic import ValidationError
from defusedxml.common import DefusedXmlException

from aspose.words_foss import Document, DocxDocument, SaveFormat, light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer import LdmPdfWriter, PdfContentLossWarning

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'


def effect_docx(effect, fill='<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'):
    w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    r = 'http://schemas.openxmlformats.org/package/2006/relationships'
    ct = 'http://schemas.openxmlformats.org/package/2006/content-types'
    parts = {
        '[Content_Types].xml': f'<Types xmlns="{ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        '_rels/.rels': f'<Relationships xmlns="{r}"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        'word/document.xml': f'''<w:document xmlns:w="{w}" xmlns:a="{A}" xmlns:wp="{wp}" xmlns:wps="{WPS}"><w:body><w:p><w:r><w:drawing>
        <wp:anchor simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="page"><wp:posOffset>2160000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>2160000</wp:posOffset></wp:positionV>
        <wp:extent cx="1440000" cy="720000"/><wp:wrapNone/><wp:docPr id="1" name="OwnedEffect"/><a:graphic><a:graphicData uri="{WPS}">
        <wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1440000" cy="720000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill}<a:ln w="25400"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>{effect}</wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>OPAQUE</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>
        </a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p><w:p><w:r><w:t>AFTER</w:t></w:r></w:p><w:sectPr/></w:body></w:document>''',
    }
    result = BytesIO()
    with ZipFile(result, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return result.getvalue()


EFFECTS = [
    '<a:effectDag name="Owned" type="tree"><a:alphaInv/></a:effectDag>',
    '<a:effectDag type="sib"><a:cont type="tree" name="Nested"><a:alphaInv><a:srgbClr val="00FF00"/></a:alphaInv></a:cont></a:effectDag>',
    '<a:effectLst><a:blur rad="127000" grow="0"/></a:effectLst>',
]


def shape(model):
    nodes = model.get_child_nodes(ldm.NodeType.SHAPE, True)
    assert len(nodes) == 1
    return nodes[0]


@pytest.mark.parametrize('effect', EFFECTS)
@pytest.mark.parametrize('api', ['source', 'dom'])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_effects_survive_source_json_and_public_saves(effect, api, format, tmp_path):
    data = effect_docx(effect)
    original = Document(BytesIO(data)).light_document_model
    source = shape(original).source_drawing_effects
    assert source is not None
    restored = ldm.Document.model_validate_json(original.model_dump_json())
    assert shape(restored).source_drawing_effects == source
    assert isinstance(shape(restored).text_box['paragraphs'][0], ldm.Paragraph)
    with pytest.warns(PdfContentLossWarning, match='shape effects'):
        LdmPdfWriter().write_to_bytes(restored)
    data = LdmDocxWriter().write_to_bytes(restored) if api == 'source' else data
    owner = Document(BytesIO(data)) if api == 'source' else DocxDocument(BytesIO(data))
    output = tmp_path / 'effects.docx'
    if api == 'source':
        owner.save(output, format)
    elif format == SaveFormat.DOCX:
        owner.save(output)
    else:
        owner.save_flat_opc(output)
    cold = Document(output).light_document_model
    assert shape(cold).source_drawing_effects == source
    with pytest.warns(PdfContentLossWarning, match='shape effects'):
        assert LdmPdfWriter().write_to_bytes(cold).startswith(b'%PDF')


@pytest.mark.parametrize('effect', EFFECTS)
def test_fill_edit_does_not_delete_independent_effects(effect):
    model = Document(BytesIO(effect_docx(effect))).light_document_model
    target = shape(model)
    source = target.source_drawing_effects
    target.fill_color = '#00FF00'
    assert target.source_drawing_fill is None
    assert target.source_drawing_effects is source
    cold = Document(BytesIO(LdmDocxWriter().write_to_bytes(model))).light_document_model
    assert shape(cold).source_drawing_effects == source


def test_effect_only_shape_is_not_dropped_and_empty_container_does_not_warn():
    model = Document(BytesIO(effect_docx('<a:effectDag name="Empty" type="tree"/>', fill=''))).light_document_model
    shape(model).stroke = None
    shape(model).text_box = None
    before = shape(model).source_drawing_effects
    cold = Document(BytesIO(LdmDocxWriter().write_to_bytes(model))).light_document_model
    assert shape(cold).source_drawing_effects == before
    assert shape(cold).drawing_position_mm == (60, 60)
    with warnings.catch_warnings():
        warnings.simplefilter('error', PdfContentLossWarning)
        LdmPdfWriter().write_to_bytes(cold)


def test_effect_loss_can_abort_public_pdf_save_without_replacing_output(tmp_path):
    owner = Document(BytesIO(effect_docx(EFFECTS[0])))
    output = tmp_path / 'existing.pdf'
    output.write_bytes(b'existing')
    before = owner.light_document_model.model_dump_json()
    with warnings.catch_warnings():
        warnings.simplefilter('error', PdfContentLossWarning)
        with pytest.raises(PdfContentLossWarning, match='shape effects'):
            owner.save(output, SaveFormat.PDF)
    assert output.read_bytes() == b'existing'
    assert owner.light_document_model.model_dump_json() == before


@pytest.mark.parametrize('xml', ['<alphaInv/>', '<a:solidFill xmlns:a="'+A+'"/>', '<!DOCTYPE x [<!ENTITY x "bad">]><x>&x;</x>'])
def test_source_effect_root_and_entity_validation(xml):
    with pytest.raises((ValidationError, DefusedXmlException)):
        ldm.SourceDrawingEffects(xml=xml)


def test_source_effect_rejects_even_an_entity_free_doctype():
    xml = f'<!DOCTYPE effectDag><a:effectDag xmlns:a="{A}"/>'
    with pytest.raises(ValidationError, match="DTDForbidden"):
        ldm.SourceDrawingEffects(xml=xml)
