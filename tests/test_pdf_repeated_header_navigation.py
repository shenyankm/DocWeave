"""Repeated table headers draw again without changing document navigation."""

from io import BytesIO

import pymupdf
from pypdf import PdfReader, PdfWriter
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.saving import PdfSaveOptions


def repeated_header_model(kind, nested=False):
    target = ldm.Paragraph(children=[ldm.BookmarkStart(name='Target'),
        ldm.Run(text='TARGET', font=ldm.Font(size=10)), ldm.BookmarkEnd(name='Target')],
        paragraph_format=ldm.ParagraphFormat(is_heading=kind == 'heading', outline_level=0 if kind == 'heading' else 9))
    cell = ldm.Cell(paragraphs=[target])
    if nested:
        cell = ldm.Cell(tables=[ldm.Table(rows=[ldm.Row(cells=[cell])])])
    table = ldm.Table(rows=[
        ldm.Row(cells=[cell], row_format=ldm.RowFormat(heading_format=True)),
        ldm.Row(cells=[ldm.Cell(paragraphs=[ldm.Paragraph(children=[ldm.Run(
            text='\n'.join(f'BODY{i:02}' for i in range(40)), font=ldm.Font(size=10))])])]),
    ])
    prefix = ldm.Paragraph(children=[ldm.Run(text='[JUMP](#Target)', is_hyperlink=True,
        font=ldm.Font(size=10))], paragraph_format=ldm.ParagraphFormat(space_after=100))
    return ldm.Document(sections=[ldm.Section(page_setup=ldm.PageSetup(
        page_width=300, page_height=200, left_margin=20, right_margin=20,
        top_margin=20, bottom_margin=20), body=ldm.Body(children=[prefix, table]))])


@pytest.mark.parametrize('kind', ['heading', 'bookmark'])
@pytest.mark.parametrize('nested', [False, True])
def test_repeated_header_keeps_first_destination_and_one_outline(kind, nested):
    model = repeated_header_model(kind, nested)
    snapshot = model.model_dump()
    options = PdfSaveOptions()
    options.outline_options.headings_outline_levels = 6
    options.outline_options.create_outlines_for_headings_in_tables = True
    options.outline_options.bookmarks_outline_levels = {'Target': 1}
    writer = LdmPdfWriter(options)
    for _ in range(2):
        raw = writer.write_to_bytes(model)
        parsed = PdfReader(BytesIO(raw))
        with pymupdf.open(stream=raw, filetype='pdf') as pdf:
            targets = [(i, word) for i, page in enumerate(pdf) for word in page.get_text('words') if word[4] == 'TARGET']
            assert len(targets) > 2  # The visual header still repeats.
            first_page, first_word = targets[0]
            link = pdf[0].get_links()[0]
            assert link['page'] == first_page
            assert abs(link['to'].x - first_word[0]) < 5
            assert abs(link['to'].y - first_word[1]) < 5
            assert len(parsed.outline) == 1
            item = parsed.outline[0]
            assert item.title == ('TARGET' if kind == 'heading' else 'Target')
            assert parsed.get_destination_page_number(item) == first_page
            text = ''.join(page.get_text() for page in pdf)
            assert all(text.count(f'BODY{i:02}') == 1 for i in range(40))
    assert model.model_dump() == snapshot


@pytest.mark.parametrize('visible_text,expected', [
    ('[Title](https://example.com)', ['TITLE']), ('\t  ', []),
])
def test_cell_heading_outline_excludes_neighboring_paragraphs_and_hidden_runs(visible_text, expected):
    heading = ldm.Paragraph(children=[ldm.Run(text='Secret', font=ldm.Font(hidden=True)),
        ldm.Run(text=visible_text, font=ldm.Font(all_caps=True))],
        paragraph_format=ldm.ParagraphFormat(is_heading=True, outline_level=0))
    table = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[heading,
        ldm.Paragraph(children=[ldm.Run(text='Ordinary body')])])])])
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[table]))])
    options = PdfSaveOptions()
    options.outline_options.headings_outline_levels = 6
    options.outline_options.create_outlines_for_headings_in_tables = True
    parsed = PdfReader(BytesIO(LdmPdfWriter(options).write_to_bytes(model)))
    assert [item.title for item in parsed.outline] == expected


def test_public_docx_save_keeps_repeated_header_bookmark_at_first_appearance():
    import aspose.words_foss as aw
    from aspose.words_foss.docx_writer import LdmDocxWriter

    source = LdmDocxWriter().write_to_bytes(repeated_header_model('bookmark', nested=True))
    document = aw.Document(BytesIO(source))
    options = PdfSaveOptions()
    options.outline_options.bookmarks_outline_levels = {'Target': 1}
    raw = document.to_bytes(options)
    parsed = PdfReader(BytesIO(raw))
    assert len(parsed.outline) == 1 and parsed.outline[0].title == 'Target'
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        first_page = next(i for i, page in enumerate(pdf) if 'TARGET' in page.get_text())
        assert pdf[0].get_links()[0]['page'] == first_page
        assert parsed.get_destination_page_number(parsed.outline[0]) == first_page


def without_artifact_text(raw):
    writer = PdfWriter(clone_from=PdfReader(BytesIO(raw)))
    parents = writer.root_object['/StructTreeRoot']['/ParentTree']['/Nums']
    arrays = {int(parents[i]): parents[i + 1].get_object() for i in range(0, len(parents), 2)}
    artifacts = 0
    for page in writer.pages:
        content = page.get_contents()
        stack, kept = [], []
        for operands, operator in content.operations:
            if operator in (b'BMC', b'BDC'):
                artifact = operands[0] == '/Artifact' or bool(stack and stack[-1])
                stack.append(artifact)
                if operands[0] == '/Artifact':
                    artifacts += 1
                if artifact and operator == b'BDC':
                    assert '/MCID' not in operands[1]
                if operator == b'BDC' and '/MCID' in operands[1]:
                    mcid = int(operands[1]['/MCID'])
                    element = arrays[int(page['/StructParents'])][mcid].get_object()
                    assert any((kid == mcid and element['/Pg'].indirect_reference == page.indirect_reference)
                               if isinstance(kid, int) else
                               (kid.get('/MCID') == mcid and kid['/Pg'].indirect_reference == page.indirect_reference)
                               for kid in element['/K'])
            elif operator == b'EMC':
                assert stack
                stack.pop()
            elif not (stack and stack[-1]):
                kept.append((operands, operator))
        assert not stack
        content.operations = kept
        page.replace_contents(content)
    stream = BytesIO()
    writer.write(stream)
    return ''.join(page.extract_text() for page in PdfReader(stream).pages), artifacts


@pytest.mark.parametrize('nested', [False, True])
@pytest.mark.parametrize('image', [False, True])
@pytest.mark.parametrize('columns', [1, 2])
def test_repeated_headers_are_artifacts_with_one_logical_appearance(nested, image, columns):
    from PIL import Image

    model = repeated_header_model('heading', nested)
    model.sections[0].page_setup.text_columns = ldm.TextColumns(count=columns, spacing=15)
    target_cell = model.sections[0].body.children[1].rows[0].cells[0]
    if nested:
        target_cell = target_cell.tables[0].rows[0].cells[0]
    target_cell.paragraphs[0].runs[0].text = 'TARGET [LINK](https://example.com)'
    if image:
        stream = BytesIO()
        Image.new('RGB', (12, 12), 'blue').save(stream, format='PNG')
        target_cell.paragraphs[0]._children.append(ldm.Shape(has_image=True, width=16, height=16,
            alternative_text='BLUE SQUARE', image_data=ldm.ImageData(image_bytes=stream.getvalue())))
    snapshot = model.model_dump()
    options = PdfSaveOptions()
    options.export_document_structure = True
    options.outline_options.headings_outline_levels = 6
    options.outline_options.create_outlines_for_headings_in_tables = True
    raw = LdmPdfWriter(options).write_to_bytes(model)
    semantic, artifacts = without_artifact_text(raw)
    assert artifacts > 0
    assert semantic.count('TARGET') == semantic.count('LINK') == 1
    assert all(semantic.count(f'BODY{i:02}') == 1 for i in range(40))
    reader = PdfReader(BytesIO(raw))
    assert len(reader.outline) == 1
    document = reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()
    figures = []

    def walk(node):
        if node['/S'] == '/Figure':
            figures.append(node)
        for child in node['/K']:
            if not isinstance(child, int):
                walk(child.get_object())

    walk(document)
    assert len(figures) == int(image)
    with pymupdf.open(stream=raw, filetype='pdf') as actual, \
            pymupdf.open(stream=LdmPdfWriter().write_to_bytes(model), filetype='pdf') as reference:
        assert len(actual) == len(reference)
        targets = [(i, word) for i, page in enumerate(actual) for word in page.get_text('words')
                   if word[4] == 'TARGET']
        assert len(targets) > 1
        assert reader.get_destination_page_number(reader.outline[0]) == targets[0][0]
        for page, control in zip(actual, reference, strict=True):
            assert page.get_text('words') == control.get_text('words')
            assert page.get_pixmap().samples == control.get_pixmap().samples
    assert model.model_dump() == snapshot


@pytest.mark.parametrize('failure', ['drawing', 'closing'])
def test_repeated_header_artifact_failure_restores_state_and_writer_recovers(tmp_path, monkeypatch, failure):
    from fpdf import FPDF
    from aspose.words_foss.pdf_writer.run_renderer import RunRenderer

    model = repeated_header_model('heading', nested=True)
    options = PdfSaveOptions()
    options.export_document_structure = True
    writer = LdmPdfWriter(options)
    output = tmp_path / 'keep.pdf'
    output.write_bytes(b'KEEP')
    captured, files = [], []
    original_render = RunRenderer._render_segment_row
    original_out = FPDF._out

    def render(renderer, pdf, *args, **kwargs):
        if getattr(pdf, '_in_structure_artifact', False):
            captured.append(pdf)
            files.extend(font.ttfont.reader.file for font in pdf.fonts.values()
                         if font.ttfont.reader is not None)
            if failure == 'drawing':
                raise RuntimeError('artifact drawing failed')
        return original_render(renderer, pdf, *args, **kwargs)

    def close(pdf, text):
        if failure == 'closing' and text == 'EMC' and getattr(pdf, '_in_structure_artifact', False):
            raise RuntimeError('artifact closing failed')
        return original_out(pdf, text)

    monkeypatch.setattr(RunRenderer, '_render_segment_row', render)
    monkeypatch.setattr(FPDF, '_out', close)
    with pytest.raises(RuntimeError, match=f'artifact {failure} failed'):
        writer.write(model, output)
    assert output.read_bytes() == b'KEEP'
    assert captured and files and all(file.closed for file in files)
    pdf = captured[0]
    assert '_in_structure_artifact' not in pdf.__dict__
    assert 'add_page' not in pdf.__dict__
    assert all(parent is pdf.struct_builder.doc_struct_elem for parent in pdf.struct_builder.parents.values())
    monkeypatch.setattr(RunRenderer, '_render_segment_row', original_render)
    monkeypatch.setattr(FPDF, '_out', original_out)
    semantic, artifacts = without_artifact_text(writer.write_to_bytes(model))
    assert artifacts > 0 and semantic.count('TARGET') == 1


def test_nested_artifact_restores_existing_flag_when_closing_fails(monkeypatch):
    from fpdf import FPDF

    pdf = FPDF()
    pdf.add_page()
    pdf._in_structure_artifact = False
    options = PdfSaveOptions()
    options.export_document_structure = True
    writer = LdmPdfWriter(options)
    commands = []

    def fail_close(text):
        commands.append(text)
        if text == 'EMC':
            raise RuntimeError('artifact closing failed')

    monkeypatch.setattr(pdf, '_out', fail_close)
    with pytest.raises(RuntimeError, match='artifact closing failed'):
        with writer._artifact(pdf):
            with writer._artifact(pdf):
                with writer._structure(pdf, '/Table', object()):
                    with writer._tag(pdf, '/P'):
                        assert pdf._in_structure_artifact
    assert pdf._in_structure_artifact is False
    assert commands == ['/Artifact BMC', 'EMC']
    assert pdf.struct_builder.empty()


def page_band_model(kind='plain', columns=1, skip_first=False):
    from aspose.words_foss.docx_reader import PAGE_FIELD_SENTINEL
    from PIL import Image

    doc = repeated_header_model('bookmark')
    section = doc.sections[0]
    section.page_setup.top_margin = 50
    section.page_setup.bottom_margin = 35
    section.page_setup.text_columns = ldm.TextColumns(count=columns, spacing=15)
    section.page_setup.different_first_page_header_footer = skip_first
    section.body.children = [ldm.Paragraph(children=[ldm.Run(text='\n'.join(f'BODY{i:02}' for i in range(50)),
                                                            font=ldm.Font(size=10))])]
    stream = BytesIO()
    Image.new('RGB', (12, 12), 'blue').save(stream, format='PNG')
    for band, name in ((0, 'HEADER'), (1, 'FOOTER')):
        text = name + (' ' + PAGE_FIELD_SENTINEL if band else ' [LINK](https://example.com)')
        para = ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(size=7)),
            ldm.Shape(has_image=True, width=14, height=14, alternative_text=name + ' IMAGE',
                      image_data=ldm.ImageData(image_bytes=stream.getvalue()))])
        children = [ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[para])])])] if kind == 'table' else [para]
        section.headers_footers.append(ldm.HeaderFooter(header_footer_type=band, children=children))
    return doc


@pytest.mark.parametrize('kind', ['plain', 'table'])
@pytest.mark.parametrize('columns', [1, 2])
@pytest.mark.parametrize('skip_first', [False, True])
def test_page_bands_are_pagination_artifacts_and_keep_body_layout(kind, columns, skip_first):
    doc = page_band_model(kind, columns, skip_first)
    snapshot = doc.model_dump()
    options = PdfSaveOptions()
    options.export_document_structure = True
    raw = LdmPdfWriter(options).write_to_bytes(doc)
    semantic, artifacts = without_artifact_text(raw)
    assert 'HEADER' not in semantic and 'FOOTER' not in semantic and 'LINK' not in semantic
    assert all(semantic.count(f'BODY{i:02}') == 1 for i in range(50))
    reader = PdfReader(BytesIO(raw))
    assert artifacts == 2 * (len(reader.pages) - int(skip_first))
    for index, page in enumerate(reader.pages):
        bands = [args[1] for args, op in page.get_contents().operations
                 if op == b'BDC' and args[0] == '/Artifact']
        assert [item['/Subtype'] for item in bands] == ([] if skip_first and index == 0 else ['/Header', '/Footer'])
        assert all(item['/Type'] == '/Pagination' for item in bands)
    document = reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()
    assert all(child.get_object()['/S'] == '/P' for child in document['/K'])
    with pymupdf.open(stream=raw, filetype='pdf') as actual, \
            pymupdf.open(stream=LdmPdfWriter().write_to_bytes(doc), filetype='pdf') as control:
        assert len(actual) == len(control)
        for index, (page, reference) in enumerate(zip(actual, control, strict=True)):
            assert page.get_text('words') == reference.get_text('words')
            assert page.get_pixmap().samples == reference.get_pixmap().samples
            if not (skip_first and index == 0):
                assert f'FOOTER {index + 1}' in page.get_text()
    assert doc.model_dump() == snapshot


@pytest.mark.parametrize('band', [0, 1])
@pytest.mark.parametrize('mode', ['before', 'pure', 'mixed', 'keep'])
@pytest.mark.parametrize('tagged', [False, True])
def test_page_band_paragraphs_do_not_advance_the_body_page(band, mode, tagged):
    doc = page_band_model()
    selected = doc.sections[0].headers_footers[band]
    para = selected.children[0]
    if mode == 'before':
        doc.sections[0].page_setup.header_distance = 60
        para.paragraph_format.page_break_before = True
    elif mode == 'pure':
        selected.children.insert(0, ldm.Paragraph(children=[ldm.Run(text='\f')]))
    elif mode == 'mixed':
        para.runs[0].text += '\f'
    else:
        para.paragraph_format.keep_together = True
        para.runs[0].text = '\n'.join(f'BAND{i:02}' for i in range(12))
    control = doc.model_copy(deep=True)
    target = control.sections[0].headers_footers[band]
    if mode == 'pure':
        target.children.pop(0)
    else:
        target.children[0].paragraph_format.page_break_before = False
        target.children[0].paragraph_format.keep_together = False
        target.children[0].runs[0].text = target.children[0].runs[0].text.replace('\f', '')
    options = PdfSaveOptions()
    options.export_document_structure = tagged
    raw = LdmPdfWriter(options).write_to_bytes(doc)
    if tagged:
        semantic, artifacts = without_artifact_text(raw)
        assert artifacts > 0 and all(semantic.count(f'BODY{i:02}') == 1 for i in range(50))
    with pymupdf.open(stream=raw, filetype='pdf') as actual, \
            pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(control), filetype='pdf') as reference:
        assert len(actual) == len(reference)
        for page, expected in zip(actual, reference, strict=True):
            assert page.get_text('words') == expected.get_text('words')
            assert page.get_pixmap().samples == expected.get_pixmap().samples


@pytest.mark.parametrize('version', ['1.4', '1.7', '2.0'])
def test_page_band_artifact_properties_respect_pdf_version(version):
    from fpdf import FPDF

    pdf = FPDF()
    pdf.pdf_version = version
    pdf.add_page()
    options = PdfSaveOptions()
    options.export_document_structure = True
    with LdmPdfWriter(options)._artifact(pdf, 'Header'):
        pass
    raw = bytes(pdf.output())
    assert raw.startswith(f'%PDF-{version}'.encode())
    properties = next(args[1] for args, op in PdfReader(BytesIO(raw)).pages[0].get_contents().operations
                      if op == b'BDC' and args[0] == '/Artifact')
    assert properties['/Type'] == '/Pagination'
    assert properties.get('/Subtype') == ('/Header' if version >= '1.7' else None)


@pytest.mark.parametrize('band', ['header', 'footer'])
@pytest.mark.parametrize('failure', ['drawing', 'closing'])
@pytest.mark.parametrize('repeated', [False, True])
def test_page_band_artifact_failure_preserves_output_and_writer_recovers(tmp_path, monkeypatch, band, failure, repeated):
    from fpdf import FPDF
    from aspose.words_foss.pdf_writer.run_renderer import RunRenderer

    model = page_band_model()
    options = PdfSaveOptions()
    options.export_document_structure = True
    writer = LdmPdfWriter(options)
    output = tmp_path / 'keep.pdf'
    output.write_bytes(b'KEEP')
    captured, files = [], []
    original_render = RunRenderer._render_segment_row
    original_out = FPDF._out

    def render(renderer, pdf, *args, **kwargs):
        if (getattr(pdf, f'_in_{band}_render', False) and
                getattr(pdf, '_in_repeated_page_band', False) == repeated):
            assert pdf._in_structure_artifact
            captured.append(pdf)
            files.extend(font.ttfont.reader.file for font in pdf.fonts.values()
                         if font.ttfont.reader is not None)
            if failure == 'drawing':
                raise RuntimeError('band drawing failed')
        return original_render(renderer, pdf, *args, **kwargs)

    def close(pdf, text):
        if (failure == 'closing' and text == 'EMC' and getattr(pdf, f'_in_{band}_render', False) and
                getattr(pdf, '_in_repeated_page_band', False) == repeated):
            raise RuntimeError('band closing failed')
        return original_out(pdf, text)

    monkeypatch.setattr(RunRenderer, '_render_segment_row', render)
    monkeypatch.setattr(FPDF, '_out', close)
    with pytest.raises(RuntimeError, match=f'band {failure} failed'):
        writer.write(model, output)
    assert output.read_bytes() == b'KEEP'
    assert captured and files and all(file.closed for file in files)
    pdf = captured[0]
    assert not pdf._in_header_render
    assert not getattr(pdf, '_in_footer_render', False)
    assert not pdf._in_repeated_page_band
    assert '_in_structure_artifact' not in pdf.__dict__
    assert pdf.auto_page_break and pdf.b_margin == pytest.approx(35 * 25.4 / 72)
    assert all(parent is pdf.struct_builder.doc_struct_elem for parent in pdf.struct_builder.parents.values())
    monkeypatch.setattr(RunRenderer, '_render_segment_row', original_render)
    monkeypatch.setattr(FPDF, '_out', original_out)
    semantic, artifacts = without_artifact_text(writer.write_to_bytes(model))
    assert artifacts > 0 and 'HEADER' not in semantic and 'FOOTER' not in semantic
    assert all(semantic.count(f'BODY{i:02}') == 1 for i in range(50))


@pytest.mark.parametrize('kind', ['plain', 'table'])
def test_public_docx_conversion_keeps_page_bands_out_of_body_structure(kind):
    import aspose.words_foss as aw
    from aspose.words_foss.docx_writer import LdmDocxWriter

    source = LdmDocxWriter().write_to_bytes(page_band_model(kind))
    document = aw.Document(BytesIO(source))
    snapshot = document.light_document_model.model_dump()
    options = PdfSaveOptions()
    options.export_document_structure = True
    raw = document.to_bytes(options)
    semantic, artifacts = without_artifact_text(raw)
    assert artifacts > 0 and 'HEADER' not in semantic and 'FOOTER' not in semantic
    assert all(semantic.count(f'BODY{i:02}') == 1 for i in range(50))
    with pymupdf.open(stream=raw, filetype='pdf') as actual, \
            pymupdf.open(stream=document.to_bytes(PdfSaveOptions()), filetype='pdf') as control:
        assert len(actual) == len(control)
        for page, reference in zip(actual, control, strict=True):
            assert page.get_text('words') == reference.get_text('words')
            assert page.get_pixmap().samples == reference.get_pixmap().samples
    assert document.light_document_model.model_dump() == snapshot


def page_band_navigation_model(kind='bookmark', layout='plain', columns=1, skip_first=False):
    doc = page_band_model('plain', columns, skip_first)
    for index, name in enumerate(('HeaderTarget', 'FooterTarget')):
        para = doc.sections[0].headers_footers[index].children[0]
        para._children = [ldm.BookmarkStart(name=name), ldm.Run(text=name, font=ldm.Font(size=7)),
                         ldm.BookmarkEnd(name=name)]
        para.paragraph_format.is_heading = kind == 'heading'
        para.paragraph_format.outline_level = 0 if kind == 'heading' else 9
        if layout != 'plain':
            table = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[para])])])
            if layout == 'nested':
                table = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(tables=[table])])])
            doc.sections[0].headers_footers[index].children = [table]
    jump = ldm.Paragraph(children=[ldm.Run(text='[HEADER JUMP](#HeaderTarget) [FOOTER JUMP](#FooterTarget)',
        is_hyperlink=True, font=ldm.Font(size=7))])
    body_heading = ldm.Paragraph(children=[ldm.Run(text='BODY HEADING', font=ldm.Font(size=7))],
        paragraph_format=ldm.ParagraphFormat(is_heading=True, outline_level=0))
    doc.sections[0].body.children[:0] = [jump, body_heading]
    return doc


@pytest.mark.parametrize('kind', ['heading', 'bookmark'])
@pytest.mark.parametrize('layout', ['plain', 'table', 'nested'])
@pytest.mark.parametrize('columns', [1, 2])
@pytest.mark.parametrize('skip_first', [False, True])
@pytest.mark.parametrize('tagged', [False, True])
def test_page_band_navigation_keeps_first_visible_destination(kind, layout, columns, skip_first, tagged):
    model = page_band_navigation_model(kind, layout, columns, skip_first)
    snapshot = model.model_dump()
    options = PdfSaveOptions()
    options.export_document_structure = tagged
    options.outline_options.headings_outline_levels = 6
    options.outline_options.create_outlines_for_headings_in_tables = True
    options.outline_options.bookmarks_outline_levels = {'HeaderTarget': 1, 'FooterTarget': 1}
    raw = LdmPdfWriter(options).write_to_bytes(model)
    reader = PdfReader(BytesIO(raw))
    assert sorted(item.title for item in reader.outline) == ['BODY HEADING', 'FooterTarget', 'HeaderTarget']
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        jumps = [link for page in pdf for link in page.get_links() if link['kind'] == pymupdf.LINK_GOTO]
        assert len(jumps) == 2
        for name in ('HeaderTarget', 'FooterTarget'):
            targets = [(index, word) for index, page in enumerate(pdf)
                       for word in page.get_text('words') if word[4] == name]
            assert len(targets) > 2
            first_page, first_word = targets[0]
            assert first_page == int(skip_first)
            item = next(item for item in reader.outline if item.title == name)
            assert reader.get_destination_page_number(item) == first_page
            assert abs(float(item['/Left']) - first_word[0]) < 5
            assert abs(pdf[first_page].rect.height - float(item['/Top']) - first_word[1]) < 5
            assert any(link['page'] == first_page and abs(link['to'].x - first_word[0]) < 5
                       and abs(link['to'].y - first_word[1]) < 5 for link in jumps)
        assert all(''.join(page.get_text() for page in pdf).count(f'BODY{i:02}') == 1 for i in range(50))
    assert model.model_dump() == snapshot


def test_page_band_navigation_resets_when_writer_reused():
    options = PdfSaveOptions()
    options.export_document_structure = True
    options.outline_options.bookmarks_outline_levels = {'HeaderTarget': 1, 'FooterTarget': 1}
    writer = LdmPdfWriter(options)
    for skip in (False, True, False):
        raw = writer.write_to_bytes(page_band_navigation_model(skip_first=skip))
        reader = PdfReader(BytesIO(raw))
        assert len(reader.outline) == 3
        targets = [item for item in reader.outline if item.title in ('HeaderTarget', 'FooterTarget')]
        assert len(targets) == 2
        assert all(reader.get_destination_page_number(item) == int(skip) for item in targets)


@pytest.mark.parametrize('layout', ['plain', 'nested'])
def test_public_docx_page_band_bookmarks_keep_first_page(layout):
    import aspose.words_foss as aw
    from aspose.words_foss.docx_writer import LdmDocxWriter

    source = LdmDocxWriter().write_to_bytes(page_band_navigation_model(layout=layout, skip_first=True))
    options = PdfSaveOptions()
    options.export_document_structure = True
    options.outline_options.bookmarks_outline_levels = {'HeaderTarget': 1, 'FooterTarget': 1}
    raw = aw.Document(BytesIO(source)).to_bytes(options)
    reader = PdfReader(BytesIO(raw))
    targets = [item for item in reader.outline if item.title in ('HeaderTarget', 'FooterTarget')]
    assert len(targets) == 2
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        for item in targets:
            first_page, word = next((index, word) for index, page in enumerate(pdf)
                for word in page.get_text('words') if word[4] == item.title)
            assert reader.get_destination_page_number(item) == first_page == 1
            assert abs(float(item['/Left']) - word[0]) < 5
        jumps = [link for page in pdf for link in page.get_links() if link['kind'] == pymupdf.LINK_GOTO]
        assert len(jumps) == 2 and all(link['page'] == 1 for link in jumps)
