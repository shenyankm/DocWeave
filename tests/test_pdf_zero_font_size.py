"""Explicit zero stays sparse zero in the model and paints at half a point."""
import pymupdf
import pytest
from aspose.words_foss import light_document_model as l
from aspose.words_foss.pdf_writer import LdmPdfWriter

@pytest.mark.parametrize('context',['body','Code','Quote','table'])
def test_explicit_zero_paints_at_half_point_without_mutating_source(context):
    font=l.Font(size=0,size_explicit=True)
    paragraph=l.Paragraph(children=[l.Run(text='OWNED_ZERO',font=font)],
        paragraph_format=l.ParagraphFormat(style_name=context if context in ('Code','Quote') else ''))
    node=l.Table(rows=[l.Row(cells=[l.Cell(paragraphs=[paragraph])])]) if context=='table' else paragraph
    doc=l.Document(sections=[l.Section(body=l.Body(children=[node]))])
    before=doc.model_dump_json()
    assert font.size==0
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(doc),filetype='pdf') as pdf:
        spans=[s for page in pdf for b in page.get_text('dict')['blocks'] for line in b.get('lines',[]) for s in line['spans'] if 'OWNED_ZERO' in s['text']]
        assert spans
        assert [s['size'] for s in spans]==pytest.approx([.5]*len(spans),abs=.011)
    assert doc.model_dump_json()==before


def test_unset_zero_preserves_code_fallback():
    assert l.Font(size=0).render_size==0
    assert l.Font(size=0,size_explicit=True).render_size==.5


def test_quote_mixed_zero_and_positive_sizes_keep_run_boundaries():
    runs=[l.Run(text='OWNED_ZERO',font=l.Font(size=0,size_explicit=True)),
          l.Run(text='OWNED_BLUE',font=l.Font(size=12,bold=True,color='0000FF')),
          l.Run(text='OWNED_RED',font=l.Font(size=24,italic=True,color='FF0000'))]
    doc=l.Document(sections=[l.Section(body=l.Body(children=[l.Paragraph(children=runs,
        paragraph_format=l.ParagraphFormat(style_name='Quote'))]))])
    before=doc.model_dump_json()
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(doc),filetype='pdf') as pdf:
        spans={''.join(chr(c[0]) for c in s['chars']):s for page in pdf for s in page.get_texttrace()}
        for text,size in [('OWNED_ZERO',.5),('OWNED_BLUE',12),('OWNED_RED',24)]:
            assert spans[text]['size']==pytest.approx(size,abs=.011)
        assert spans['OWNED_BLUE']['color']==pytest.approx((0,0,1),abs=.001)
        assert spans['OWNED_RED']['color']==pytest.approx((1,0,0),abs=.001)
    assert doc.model_dump_json()==before


@pytest.mark.parametrize('input_size,expected',[(0,0),(.01,0),(.1,0),(.24,0),(.25,.5),(.26,.5),(.5,.5),(1,1),(2.25,2.5),(12.25,12.5),(12.75,13)])
def test_public_numeric_size_rounds_half_up_across_input_routes(input_size,expected):
    constructor=l.Font(size=input_size,size_explicit=True)
    assigned=l.Font();assigned.size=input_size
    copied=l.Font().model_copy(update={'size':input_size})
    for font in (constructor,assigned,copied,l.Font.model_validate_json(assigned.model_dump_json())):
        assert font.size==expected
        assert font.render_size==(expected or .5)


@pytest.mark.parametrize('normal,whole,conditional,expected', [
    (None,None,0,0), (12,None,0,11), (12,0,None,0),
    (12,0,0,0), (12,20,0,20), (12,0,22,0),
])
def test_table_zero_size_keeps_normal_sentinel_and_sparse_source(
        tmp_path, normal, whole, conditional, expected):
    from io import BytesIO
    from zipfile import ZipFile
    from xml.etree import ElementTree as E
    import aspose.words_foss as aw
    from aspose.words_foss.docx_writer import LdmDocxWriter
    from aspose.words_foss.saving import OoxmlSaveOptions
    from .test_table_font_size_live import owned_table, target, W

    path=owned_table(tmp_path/'zero.docx',normal)
    with ZipFile(path) as archive:
        parts={name:archive.read(name) for name in archive.namelist()}
    styles=E.fromstring(parts['word/styles.xml'])
    table=next(style for style in styles if style.get('{'+W+'}type')=='table')
    region=table.find('{'+W+'}tblStylePr')
    if conditional is None:
        table.remove(region)
    else:
        region.find('{'+W+'}rPr/{'+W+'}sz').set('{'+W+'}val',str(conditional*2))
    if whole is not None:
        E.SubElement(E.SubElement(table,'{'+W+'}rPr'),'{'+W+'}sz').set('{'+W+'}val',str(whole*2))
    parts['word/styles.xml']=E.tostring(styles)
    with ZipFile(path,'w') as archive:
        for name,data in parts.items():archive.writestr(name,data)
    model=aw.Document(path).light_document_model
    assert target(model).font.size==expected
    assert target(model).font.size_explicit is False
    assert aw.DocxDocument(path).body.tables[0].rows[0].cells[0].paragraphs[0].runs[0].effective_font.size==expected
    cold_expected = 12 if normal == 12 else expected
    for fmt in (aw.SaveFormat.DOCX,aw.SaveFormat.FLAT_OPC):
        current=l.Document.model_validate_json(model.model_dump_json())
        for _ in range(2):
            data=LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(current)
            current=aw.Document(BytesIO(data)).light_document_model
            assert target(current).font.size==cold_expected
            assert target(current).font.size_explicit is False
            assert aw.DocxDocument(BytesIO(data)).body.tables[0].rows[0].cells[0].paragraphs[0].runs[0].effective_font.size==cold_expected
