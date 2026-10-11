"""Section ownership, variants and part-local resources survive OOXML conversion."""
from io import BytesIO
import pytest
import dockiln as dk
from dockiln import light_document_model as ldm


def p(text):
    return ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(bold=True))])


@pytest.mark.parametrize('fmt', ['DOCX', 'FLAT_OPC'])
@pytest.mark.parametrize('via_json', [False, True])
def test_each_section_and_variant_remains_distinct(tmp_path, fmt, via_json):
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[p('BODY 0')])),
        ldm.Section(body=ldm.Body(children=[p('BODY 1')]), headers_footers=[
            ldm.HeaderFooter(header_footer_type=k, children=[p(f'OWNED SECTION 1 KIND {k}')])
            for k in range(6)]),
        ldm.Section(body=ldm.Body(children=[p('BODY 2')]), headers_footers=[
            ldm.HeaderFooter(header_footer_type=k, children=[p(f'OWNED SECTION 2 KIND {k}')])
            for k in range(6)]),
        ldm.Section(body=ldm.Body(children=[p('BODY 3')]))])
    model.sections[1].page_setup.different_first_page_header_footer = True
    model.sections[1].page_setup.odd_and_even_pages_header_footer = True
    if via_json:
        model = ldm.Document.model_validate_json(model.model_dump_json())
    doc = dk.Document(); doc._document = model
    for step in (1, 2):
        path = tmp_path / f'{fmt}-{step}.bin';doc.save(path, getattr(dk.SaveFormat, fmt))
        doc = dk.Document(path);cold = doc.light_document_model
        assert cold.sections[0].headers_footers == cold.sections[3].headers_footers == []
        assert cold.sections[1].page_setup.different_first_page_header_footer is True
        assert all(s.page_setup.odd_and_even_pages_header_footer for s in cold.sections)
        for sec in (1, 2):
            assert {hf.header_footer_type: [p.text for p in hf.paragraphs]
                    for hf in cold.sections[sec].headers_footers} == {
                        k: [f'OWNED SECTION {sec} KIND {k}'] for k in range(6)}
            assert all(hf.paragraphs[0]._children[0].font.bold for hf in cold.sections[sec].headers_footers)
        assert not any(d.code.endswith('header_footer_flattened') for d in doc.diagnostics)


@pytest.mark.parametrize('fmt', ['DOCX', 'FLAT_OPC'])
def test_explicit_empty_story_breaks_previous_header_inheritance(tmp_path, fmt):
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[p('BODY 0')]),
        headers_footers=[ldm.HeaderFooter(header_footer_type=0, children=[p('VISIBLE')])]),
        ldm.Section(body=ldm.Body(children=[p('BODY 1')]),
                    headers_footers=[ldm.HeaderFooter(header_footer_type=0)]),
        ldm.Section(body=ldm.Body(children=[p('BODY 2')]))])
    doc = dk.Document();doc._document = model
    for step in (1, 2):
        path = tmp_path / f'{fmt}-{step}.bin';doc.save(path,getattr(dk.SaveFormat, fmt));doc = dk.Document(path)
        cold=doc.light_document_model
        assert cold.sections[0].headers_footers[0].paragraphs[0].text == 'VISIBLE'
        assert cold.sections[1].headers_footers[0].paragraphs[0].text == ''
        assert cold.sections[2].headers_footers == []


@pytest.mark.parametrize('kinds', [[0, 0], [None], [-1], [6]])
def test_invalid_story_kind_is_rejected_before_output(kinds):
    model=ldm.Document(sections=[ldm.Section(headers_footers=[
        ldm.HeaderFooter(header_footer_type=k) for k in kinds])])
    doc=dk.Document();doc._document=model
    stream=BytesIO(b'KEEP')
    with pytest.raises(ValueError):doc.save(stream,dk.SaveFormat.DOCX)
    assert stream.getvalue()==b'KEEP'


@pytest.mark.parametrize('fmt', ['DOCX', 'FLAT_OPC'])
@pytest.mark.parametrize('via_json', [False, True])
def test_part_local_images_and_mixed_table_order_are_preserved(tmp_path, fmt, via_json):
    from PIL import Image
    payloads = []
    sections = []
    for sec in (0, 1):
        stories = []
        for kind in (0, 1, 2, 3, 4, 5):
            image = BytesIO();Image.new('RGB', (3, 2), (sec * 100, kind * 30, 80)).save(image, 'PNG')
            payload = image.getvalue();payloads.append(payload)
            inline = ldm.Shape(is_inline=True, has_image=True, width=3, height=2,
                image_data=ldm.ImageData(image_type=ldm.ImageData.from_mime('image/png'),image_bytes=payload))
            stories.append(ldm.HeaderFooter(header_footer_type=kind, children=[
                p(f'BEFORE {sec} {kind}'), ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(children=[
                    ldm.Paragraph(children=[ldm.Run(text=f'CELL {sec} {kind}'), inline,
                        ldm.Run(text=f'[LINK {sec} {kind}](https://owned.invalid/{sec}/{kind})', is_hyperlink=True)])])])]),
                p(f'AFTER {sec} {kind}')]))
        sections.append(ldm.Section(body=ldm.Body(children=[p(f'BODY {sec}')]),headers_footers=stories))
    model=ldm.Document(sections=sections)
    if via_json:model=ldm.Document.model_validate_json(model.model_dump_json())
    doc=dk.Document();doc._document=model
    for step in (1, 2):
        path=tmp_path/f'{fmt}-{step}.bin';doc.save(path,getattr(dk.SaveFormat, fmt));doc=dk.Document(path)
        actual=[]
        for sec, section in enumerate(doc.light_document_model.sections):
            for hf in section.headers_footers:
                kind=hf.header_footer_type
                assert [type(c).__name__ for c in hf.children]==['Paragraph','Table','Paragraph']
                assert hf.children[0].text==f'BEFORE {sec} {kind}'
                assert hf.children[2].text==f'AFTER {sec} {kind}'
                cell=hf.children[1].rows[0].cells[0].paragraphs[0]
                assert cell.text==f'CELL {sec} {kind}[LINK {sec} {kind}](https://owned.invalid/{sec}/{kind})'
                actual.append(cell._children[1].image_data.image_bytes)
        assert actual==payloads


def test_pdf_shared_header_diagnostic_covers_story_that_starts_in_later_section():
    from dockiln.diagnostics import header_footer_losses
    doc=ldm.Document(sections=[ldm.Section(),ldm.Section(headers_footers=[
        ldm.HeaderFooter(header_footer_type=0, children=[p("LATER")])])])
    assert "header_footer_flattened" in dict(header_footer_losses(doc))


@pytest.mark.parametrize('fmt', ['DOCX', 'FLAT_OPC'])
@pytest.mark.parametrize('via_json', [False, True])
def test_header_resources_and_comment_thread_edit_coexist(tmp_path, fmt, via_json):
    from PIL import Image
    from tests.test_comment_threads import owned_thread
    model=dk.Document(BytesIO(owned_thread())).light_document_model
    stories=[s for s in model.source_stories if s.kind=='comment']
    stories[0].comment_metadata.done=False
    stories[1].comment_metadata.done=True
    stories[1].children[-1].runs[-1].text='OWNED EDITED REPLY'
    stories[1].children.append(p('OWNED SECOND PARAGRAPH'))
    payload=BytesIO();Image.new('RGB',(3,2),'navy').save(payload,'PNG');image=payload.getvalue()
    for kind in range(6):
        model.sections[0].headers_footers.append(ldm.HeaderFooter(header_footer_type=kind,children=[
            ldm.Paragraph(children=[ldm.Run(text=f'OWNED HF {kind}'),
                ldm.Shape(is_inline=True,has_image=True,width=3,height=2,
                    image_data=ldm.ImageData(image_type=ldm.ImageData.from_mime('image/png'),image_bytes=image)),
                ldm.Run(text=f'[LINK {kind}](https://owned.invalid/hf/{kind})',is_hyperlink=True)])]))
    if via_json:model=ldm.Document.model_validate_json(model.model_dump_json())
    doc=dk.Document();doc._document=model
    for step in (1,2):
        path=tmp_path/f'{fmt}-{step}.bin';doc.save(path,getattr(dk.SaveFormat,fmt));doc=dk.Document(path)
        model=doc.light_document_model
        replies=[s for s in model.source_stories if s.kind=='comment']
        assert len(replies)==3
        assert replies[0].comment_metadata.done is False
        assert replies[1].comment_metadata.done is True
        assert replies[1].comment_metadata.ancestor_identifier==replies[0].identifier
        assert [p.text for p in replies[1].children]==['OWNED EDITED REPLY','OWNED SECOND PARAGRAPH']
        assert {hf.header_footer_type for hf in model.sections[0].headers_footers}==set(range(6))
        for hf in model.sections[0].headers_footers:
            kind=hf.header_footer_type;children=hf.paragraphs[0]._children
            assert children[0].text==f'OWNED HF {kind}'
            assert children[1].image_data.image_bytes==image
            assert children[2].text==f'[LINK {kind}](https://owned.invalid/hf/{kind})'
