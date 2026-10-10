"""DrawingML join source defaults remain distinct from legacy Border defaults."""
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import warnings

import pytest
from pydantic import ValidationError

from aspose.words_foss import Document, SaveFormat, light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer.diagnostics import PdfContentLossWarning
from tests.test_drawing_effect_sources import effect_docx,A,WPS


def join_docx(join, attrs=None):
    with ZipFile(BytesIO(effect_docx(''))) as z:
        parts={n:z.read(n) for n in z.namelist()}
    root=ET.fromstring(parts['word/document.xml']);sp=root.find('.//{'+WPS+'}spPr')
    if join:ET.SubElement(sp.find('{'+A+'}ln'),'{'+A+'}'+join,attrs or {})
    parts['word/document.xml']=ET.tostring(root)
    out=BytesIO()
    with ZipFile(out,'w')as z:
        for n,b in parts.items():z.writestr(n,b)
    return out.getvalue()


def stroke(owner):
    return owner.light_document_model.get_child_nodes(ldm.NodeType.SHAPE,True)[0].stroke


@pytest.mark.parametrize('join,attrs,j', [('',{},1),('round',{},1),('bevel',{},2),('miter',{},0),('miter',{'lim':'100000'},0)])
@pytest.mark.parametrize('format',[SaveFormat.DOCX,SaveFormat.FLAT_OPC])
def test_join_survives_json_edits_save_cold_and_pdf(tmp_path,join,attrs,j,format):
    fitz=pytest.importorskip('pymupdf')
    model=Document(BytesIO(join_docx(join,attrs))).light_document_model
    model=ldm.Document.model_validate_json(model.model_dump_json())
    owner=Document(BytesIO(LdmDocxWriter().write_to_bytes(model)))
    source=stroke(owner).source_drawing_join
    assert source is not None
    if join:
        node=ET.fromstring(source.xml)
        assert node.tag=='{'+A+'}'+join
        assert node.attrib==attrs
    else:assert source.xml==''
    stroke(owner).color='Color [A=255, R=0, G=128, B=192]'
    stroke(owner).line_width=8
    output=tmp_path/'output';owner.save(output,format)
    cold=Document(output)
    assert stroke(cold).source_drawing_join==source
    assert stroke(cold).line_width==8
    before=cold.light_document_model.model_dump_json()
    pdf=cold.to_bytes(SaveFormat.PDF)
    drawings=fitz.open(stream=pdf,filetype='pdf')[0].get_drawings()
    outline=[d for d in drawings if d.get('color') and d['width']==8]
    assert len(outline)==1
    assert outline[0]['lineJoin']==j
    assert cold.light_document_model.model_dump_json()==before


def test_legacy_border_is_not_forced_to_drawingml_default():
    assert ldm.Border().source_drawing_join is None


@pytest.mark.parametrize('join',['unknownJoin','round'])
def test_unknown_or_multiple_join_input_is_rejected(join):
    data=join_docx(join)
    if join=='round':
        with ZipFile(BytesIO(data))as z:parts={n:z.read(n)for n in z.namelist()}
        root=ET.fromstring(parts['word/document.xml']);ET.SubElement(root.find('.//{'+A+'}ln'),'{'+A+'}bevel')
        parts['word/document.xml']=ET.tostring(root);out=BytesIO()
        with ZipFile(out,'w')as z:
            for n,b in parts.items():z.writestr(n,b)
        data=out.getvalue()
    with pytest.raises(ValueError,match='line join declarations'):Document(BytesIO(data))


@pytest.mark.parametrize('lim',['-1','broken','100%'])
def test_unimplemented_or_invalid_limit_aborts_atomically(tmp_path,lim):
    owner=Document(BytesIO(join_docx('miter',{'lim':lim})))
    output=tmp_path/'existing.pdf';output.write_bytes(b'existing')
    with warnings.catch_warnings():
        warnings.simplefilter('error',PdfContentLossWarning)
        with pytest.raises(PdfContentLossWarning,match='miter join limit'):owner.save(output,SaveFormat.PDF)
    assert output.read_bytes()==b'existing'
    assert stroke(owner).source_drawing_join.xml


@pytest.mark.parametrize('xml,error', [('<bad/>',ValidationError),('<!DOCTYPE a:round><a:round xmlns:a="'+A+'"/>',ValidationError)])
def test_source_join_xml_trust_boundary(xml,error):
    with pytest.raises(error):ldm.SourceDrawingLineJoin(xml=xml)


def test_join_context_restores_legacy_sentinel_graphics_state():
    fitz=pytest.importorskip('pymupdf')
    data=join_docx('bevel')
    with ZipFile(BytesIO(data))as z:parts={n:z.read(n)for n in z.namelist()}
    import copy
    root=ET.fromstring(parts['word/document.xml']);body=root.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')
    sentinel=copy.deepcopy(body[0]);wp='{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
    sentinel.find('.//'+wp+'positionV/'+wp+'posOffset').text='3060000'
    body.insert(1,sentinel);parts['word/document.xml']=ET.tostring(root)
    out=BytesIO()
    with ZipFile(out,'w')as z:
        for n,b in parts.items():z.writestr(n,b)
    owner=Document(BytesIO(out.getvalue()));nodes=owner.light_document_model.get_child_nodes(ldm.NodeType.SHAPE,True)
    nodes[1].stroke.source_drawing_join=None
    before=owner.light_document_model.model_dump_json()
    page=fitz.open(stream=owner.to_bytes(SaveFormat.PDF),filetype='pdf')[0]
    paths=[d for d in page.get_drawings()if d.get('color') and d['width']==2]
    assert [d['lineJoin']for d in paths]==[2,0]
    assert owner.light_document_model.model_dump_json()==before


def test_uncalibrated_join_geometry_keeps_loss_diagnostic():
    owner=Document(BytesIO(join_docx('round')))
    target=owner.light_document_model.get_child_nodes(ldm.NodeType.SHAPE,True)[0]
    target.source_drawing_fill=None
    source=target.stroke.source_drawing_join
    with pytest.warns(PdfContentLossWarning,match='geometry not calibrated'):
        owner.to_bytes(SaveFormat.PDF)
    assert target.stroke.source_drawing_join==source


@pytest.mark.parametrize('angle', [-20, 67, 83, 168])
def test_page_space_outline_keeps_rotated_text_and_following_text_context(angle):
    import copy
    fitz = pytest.importorskip('pymupdf')
    owner = Document(BytesIO(join_docx('round')))
    shape = owner.light_document_model.get_child_nodes(ldm.NodeType.SHAPE, True)[0]
    shape.source_drawing_fill = shape.source_drawing_fill.model_copy(update={'rotation': angle * 60000})
    shape.text_box = {'paragraphs': [ldm.Paragraph(children=[
        ldm.Run(text='ROTATED', font=ldm.Font(name='Courier New', size=12))])],
        'insets_mm': [2, 2, 2, 2]}
    control = copy.deepcopy(owner)
    # Both controls use captured geometry; changing joins must not move text.
    stroke(control).source_drawing_join = ldm.SourceDrawingLineJoin(
        xml='<a:miter xmlns:a="' + A + '"/>')
    before = owner.light_document_model.model_dump_json()
    def spans(document):
        page = fitz.open(stream=document.to_bytes(SaveFormat.PDF), filetype='pdf')[0]
        return [(s['text'], s['bbox'], line['dir'])
                for block in page.get_text('dict')['blocks'] if 'lines' in block
                for line in block['lines'] for s in line['spans']]
    observed = spans(owner)
    assert any(text == 'ROTATED' for text, _, _ in observed)
    assert any(text == 'AFTER' and direction == (1, 0) for text, _, direction in observed)
    assert observed == spans(control)
    assert owner.light_document_model.model_dump_json() == before


def test_rotated_gradient_outline_is_preserved_with_explicit_loss_diagnostic():
    owner = Document(BytesIO(join_docx('round')))
    target = owner.light_document_model.get_child_nodes(ldm.NodeType.SHAPE, True)[0]
    gradient = ('<a:gradFill xmlns:a="' + A + '"><a:gsLst>'
        '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>'
        '<a:gs pos="50000"><a:srgbClr val="0000FF"/></a:gs>'
        '<a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs>'
        '</a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>')
    target.source_drawing_fill = target.source_drawing_fill.model_copy(
        update={'rotation': 67 * 60000, 'direct_xml': gradient})
    before = owner.light_document_model.model_dump_json()
    with pytest.warns(PdfContentLossWarning, match='rotated non-solid line join geometry'):
        owner.to_bytes(SaveFormat.PDF)
    assert owner.light_document_model.model_dump_json() == before
