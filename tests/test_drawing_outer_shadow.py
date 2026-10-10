"""Owned zero-blur shadow projection, diagnostics and independent source intent."""
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import warnings

import pytest

from aspose.words_foss import Document, light_document_model as ldm
from aspose.words_foss._drawing_effects import resolve_outer_shadow
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer.diagnostics import PdfContentLossWarning
from aspose.words_foss.saving import DmlEffectsRenderingMode, PdfSaveOptions
from tests.test_drawing_effect_sources import effect_docx, A


def shadow_xml(attrs='', colour='<a:srgbClr val="008000"/>'):
    return f'<a:effectLst xmlns:a="{A}"><a:outerShdw blurRad="0" dist="360000" {attrs}>{colour}</a:outerShdw></a:effectLst>'


def no_outline(data):
    with ZipFile(BytesIO(data)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    root = ET.fromstring(parts['word/document.xml'])
    sp = root.find('.//{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}spPr')
    sp.remove(sp.find('{'+A+'}ln'))
    ET.SubElement(ET.SubElement(sp,'{'+A+'}ln'),'{'+A+'}noFill')
    parts['word/document.xml'] = ET.tostring(root)
    out = BytesIO()
    with ZipFile(out,'w') as archive:
        for name, content in parts.items(): archive.writestr(name,content)
    return out.getvalue()


@pytest.mark.parametrize('mode', [DmlEffectsRenderingMode.SIMPLIFIED, DmlEffectsRenderingMode.FINE])
@pytest.mark.parametrize('direction,offset', [(0,(10,0)),(5400000,(0,10)),(10800000,(-10,0)),(16200000,(0,-10))])
@pytest.mark.parametrize('alpha', [None, 50000])
def test_pdf_outer_shadow_offset_alpha(tmp_path, mode, direction, offset, alpha):
    fitz = pytest.importorskip('pymupdf')
    colour = '<a:srgbClr val="008000">'+(f'<a:alpha val="{alpha}"/>' if alpha else '')+'</a:srgbClr>'
    owner = Document(BytesIO(no_outline(effect_docx(shadow_xml(f'dir="{direction}"',colour)))))
    # JSON restoration must retain an independent effect source before paint.
    model = ldm.Document.model_validate_json(owner.light_document_model.model_dump_json())
    owner = Document(BytesIO(LdmDocxWriter().write_to_bytes(model)))
    before = owner.light_document_model.model_dump_json()
    options = PdfSaveOptions();options.dml_effects_rendering_mode = mode
    output=tmp_path/'shadow.pdf'
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter('always');owner.save(output,options)
    assert not any(isinstance(w.message,PdfContentLossWarning) for w in caught)
    assert owner.light_document_model.model_dump_json() == before
    page = fitz.open(output)[0]
    green = [d for d in page.get_drawings() if d.get('fill') and d['fill'][1] > .49 and d['fill'][0] == 0]
    assert len(green) == 1
    rect=green[0]['rect'];pt=72/25.4
    assert tuple(rect) == pytest.approx(((60+offset[0])*pt,(60+offset[1])*pt,(100+offset[0])*pt,(80+offset[1])*pt),abs=.01)
    assert green[0]['fill_opacity'] == pytest.approx(128/255 if alpha else 1, abs=.0001)


@pytest.mark.parametrize('attrs', ['blurRad="12700"','sx="150000"','ky="60000"','algn="tl"','dist="-1"','dir="21600000"','unknown="1"'])
def test_unsupported_shadow_keeps_source_and_can_abort_atomically(tmp_path,attrs):
    # Override a single default attribute rather than generate duplicate attrs.
    xml=shadow_xml().replace('blurRad="0"','').replace('dist="360000"','')
    xml=xml.replace('<a:outerShdw ',f'<a:outerShdw {attrs} ')
    owner=Document(BytesIO(no_outline(effect_docx(xml))));before=owner.light_document_model.model_dump_json()
    output=tmp_path/'existing.pdf';output.write_bytes(b'existing')
    with warnings.catch_warnings():
        warnings.simplefilter('error',PdfContentLossWarning)
        with pytest.raises(PdfContentLossWarning):owner.save(output,PdfSaveOptions())
    assert output.read_bytes()==b'existing'
    assert owner.light_document_model.model_dump_json()==before


def test_shadow_transform_does_not_mutate_source():
    source=ldm.SourceDrawingEffects(xml=shadow_xml('dir="2700000"'))
    before=source.model_dump_json()
    dx,dy,colour=resolve_outer_shadow(source,None)
    assert (dx,dy)==pytest.approx((2**.5*5,2**.5*5))
    assert colour==(0,128,0)
    assert source.model_dump_json()==before


@pytest.mark.parametrize('fill,outline', [('<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>',False), ('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',True)])
def test_shadow_transparent_fill_or_outline_is_still_diagnosed(fill,outline):
    data=effect_docx(shadow_xml(),fill)
    if not outline:data=no_outline(data)
    with pytest.warns(PdfContentLossWarning,match='outer shadow shape or fill'):
        Document(BytesIO(data)).to_bytes(PdfSaveOptions())


@pytest.mark.parametrize('colour', ['', '<a:srgbClr val="008000"/><a:schemeClr val="accent1"/>'])
def test_shadow_invalid_colour_count_is_not_silently_discarded(colour):
    source=ldm.SourceDrawingEffects(xml=shadow_xml(colour=colour))
    with pytest.raises(NotImplementedError,match='colour declarations'):
        resolve_outer_shadow(source,None)
