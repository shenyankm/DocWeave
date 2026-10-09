"""Marked-content sequences and page parent mappings must survive pagination."""
from io import BytesIO

import pymupdf
from pypdf import PdfReader
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.saving import PdfSaveOptions


def model(style='Normal', bands=False, columns=1):
    paragraph = ldm.Paragraph(children=[ldm.Run(text='\n'.join(f'BODY{i:03}' for i in range(50)),
                                               font=ldm.Font(size=12))],
        paragraph_format=ldm.ParagraphFormat(style_name=style, is_heading=style=='Heading',
                                            outline_level=0 if style == 'Heading' else 9))
    section = ldm.Section(page_setup=ldm.PageSetup(page_width=280, page_height=180,
        left_margin=20, right_margin=20, top_margin=25, bottom_margin=25,
        text_columns=ldm.TextColumns(count=columns, spacing=15)), body=ldm.Body(children=[paragraph]))
    if bands:
        section.headers_footers = [ldm.HeaderFooter(header_footer_type=kind, children=[
            ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(size=8))])])
            for kind, text in ((0, 'HEADER'), (1, 'FOOTER'))]
    return ldm.Document(sections=[section])


def assert_page_tags(raw):
    reader = PdfReader(BytesIO(raw))
    assert len(reader.pages) > 1
    parents = reader.trailer['/Root']['/StructTreeRoot']['/ParentTree']['/Nums']
    parent_arrays = {int(parents[i]):parents[i+1].get_object() for i in range(0,len(parents),2)}
    for page in reader.pages:
        stack = []
        mcids = []
        for operands, operator in page.get_contents().operations:
            if operator in (b'BMC', b'BDC'):
                stack.append(operator)
                if operator == b'BDC' and '/MCID' in operands[1]:
                    mcids.append(int(operands[1]['/MCID']))
            elif operator == b'EMC':
                assert stack, 'Unmatched marked-content end on this page'
                stack.pop()
        assert not stack, 'Marked-content sequence left open at the end of this page'
        assert mcids and len(mcids) == len(set(mcids))
        references = parent_arrays[int(page['/StructParents'])]
        for mcid in mcids:
            element = references[mcid].get_object()
            assert any((kid == mcid and element['/Pg'].indirect_reference == page.indirect_reference)
                       if isinstance(kid, int) else
                       (kid.get('/MCID') == mcid and kid['/Pg'].indirect_reference == page.indirect_reference)
                       for kid in element['/K'])
    return reader


@pytest.mark.parametrize('style', ['Normal', 'Heading', 'Quote', 'Code'])
@pytest.mark.parametrize('bands', [False, True])
@pytest.mark.parametrize('columns', [1, 2])
def test_multi_page_tags_are_balanced_and_keep_rendered_content(style, bands, columns):
    doc = model(style, bands, columns)
    options = PdfSaveOptions()
    options.export_document_structure = True
    tagged = LdmPdfWriter(options).write_to_bytes(doc)
    reader = assert_page_tags(tagged)
    untagged = LdmPdfWriter().write_to_bytes(doc)
    with pymupdf.open(stream=tagged,filetype='pdf') as actual, \
            pymupdf.open(stream=untagged,filetype='pdf') as reference:
        assert len(actual) == len(reference) == len(reader.pages)
        text = ''.join(page.get_text() for page in actual)
        assert all(text.count(f'BODY{i:03}') == 1 for i in range(50))
        for page, control in zip(actual, reference, strict=True):
            assert page.get_text('words') == control.get_text('words')
            assert page.get_pixmap().samples == control.get_pixmap().samples


@pytest.mark.parametrize('failure', ['drawing', 'closing'])
@pytest.mark.parametrize('content', ['paragraph', 'table'])
def test_tag_page_hook_restores_after_failure_and_writer_recovers(tmp_path, monkeypatch, failure, content):
    from fpdf import FPDF
    from aspose.words_foss.pdf_writer.run_renderer import RunRenderer

    options = PdfSaveOptions()
    options.export_document_structure = True
    writer = LdmPdfWriter(options)
    output = tmp_path / 'keep.pdf'
    output.write_bytes(b'KEEP')
    captured = []
    files = []
    method = '_render_segment_row' if content == 'table' else 'render_formatted_runs_aligned'
    original = getattr(RunRenderer, method)
    original_out = FPDF._out
    doc = model()
    if content == 'table':
        paragraph = doc.sections[0].body.children[0]
        doc.sections[0].body.children = [ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph])])])]

    def fail(renderer, pdf, *args, **kwargs):
        captured.append(pdf)
        files.extend(font.ttfont.reader.file for font in pdf.fonts.values()
                     if font.ttfont.reader is not None)
        assert 'add_page' in pdf.__dict__
        if failure == 'drawing':
            raise RuntimeError('tagged drawing failed')
        return original(renderer, pdf, *args, **kwargs)

    def fail_close(pdf, text):
        if text == 'EMC':
            raise RuntimeError('tagged closing failed')
        return original_out(pdf, text)

    monkeypatch.setattr(RunRenderer, method, fail)
    if failure == 'closing':
        monkeypatch.setattr(FPDF, '_out', fail_close)
    with pytest.raises(RuntimeError, match=f'tagged {failure} failed'):
        writer.write(doc, output)
    assert output.read_bytes() == b'KEEP'
    assert 'add_page' not in captured[0].__dict__
    builder = captured[0].struct_builder
    assert all(parent is builder.doc_struct_elem for parent in builder.parents.values())
    assert files and all(file.closed for file in files)
    monkeypatch.setattr(RunRenderer, method, original)
    monkeypatch.setattr(FPDF, '_out', original_out)
    assert_page_tags(writer.write_to_bytes(doc))


@pytest.mark.parametrize('overridden', [False, True])
def test_nested_tag_hooks_restore_when_closing_fails(monkeypatch, overridden):
    from fpdf import FPDF

    pdf = FPDF()
    pdf.add_page()
    if overridden:
        native_add_page = pdf.add_page
        pdf.add_page = lambda *args, **kwargs: native_add_page(*args, **kwargs)
    original = pdf.add_page
    original_out = pdf._out
    options = PdfSaveOptions()
    options.export_document_structure = True

    def fail_close(text):
        if text == 'EMC':
            raise RuntimeError('closing failed')
        return original_out(text)

    monkeypatch.setattr(pdf, '_out', fail_close)
    writer = LdmPdfWriter(options)
    with pytest.raises(RuntimeError, match='closing failed'):
        with writer._tag(pdf, '/TD'):
            with writer._tag(pdf, '/P'):
                pass
    assert pdf.add_page == original
    assert ('add_page' in pdf.__dict__) == overridden


@pytest.mark.parametrize('bands', [False, True])
@pytest.mark.parametrize('columns', [1, 2])
def test_table_structure_keeps_nested_and_split_row_relationships(bands, columns):
    def paragraph(text):
        return ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(size=9))])

    def cell(text):
        return ldm.Cell(paragraphs=[paragraph(text)])

    nested = ldm.Table(rows=[ldm.Row(cells=[cell('NESTA'), cell('NESTB')]),
                             ldm.Row(cells=[cell('NESTC'), cell('NESTD')])])
    outer = ldm.Table(rows=[
        ldm.Row(cells=[cell('HEADA'), cell('HEADB')], row_format=ldm.RowFormat(heading_format=True)),
        ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph('PARENT')], tables=[nested]), cell('SIDE')]),
        ldm.Row(cells=[cell('\n'.join(f'LONG{i:03}' for i in range(40))), cell('ENDCELL')]),
    ])
    doc = model(bands=bands, columns=columns)
    doc.sections[0].body.children = [paragraph('BEFORE'), outer, paragraph('AFTER')]
    options = PdfSaveOptions()
    options.export_document_structure = True
    raw = LdmPdfWriter(options).write_to_bytes(doc)
    reader = assert_page_tags(raw)
    document = reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()

    def children(element):
        return [item.get_object() for item in element['/K'] if not isinstance(item, int)]

    roots = children(document)
    tables = [item for item in roots if item['/S'] == '/Table']
    assert len(tables) == 1
    rows = children(tables[0])
    assert len(rows) == 3 and all(item['/S'] == '/TR' for item in rows)
    for index, row in enumerate(rows):
        cells = children(row)
        assert len(cells) == 2
        assert all(item['/S'] == ('/TH' if index == 0 else '/TD') for item in cells)
    parent_cell = children(rows[1])[0]
    nested_tables = [item for item in children(parent_cell) if item['/S'] == '/Table']
    assert len(nested_tables) == 1
    assert len(children(nested_tables[0])) == 2
    paragraphs = children(children(rows[2])[0])
    assert len(paragraphs) == 1 and paragraphs[0]['/S'] == '/P'
    assert len(children(paragraphs[0])) >= 40
    seen = set()

    def walk(element):
        ref = element.indirect_reference
        assert ref.idnum not in seen
        seen.add(ref.idnum)
        for child in children(element):
            assert child['/P'].indirect_reference == ref
            walk(child)

    walk(document)
    with pymupdf.open(stream=raw, filetype='pdf') as actual, \
            pymupdf.open(stream=LdmPdfWriter().write_to_bytes(doc), filetype='pdf') as control:
        assert len(actual) == len(control)
        text = ''.join(page.get_text() for page in actual)
        assert all(text.count(f'LONG{i:03}') == 1 for i in range(40))
        for page, reference in zip(actual, control, strict=True):
            assert page.get_text('words') == reference.get_text('words')
            assert page.get_pixmap().samples == reference.get_pixmap().samples


@pytest.mark.parametrize('kind', ['empty', 'image', 'aliases', 'bands', 'floating'])
def test_table_structure_handles_empty_images_aliases_and_page_bands(kind):
    from PIL import Image

    def paragraph(text):
        return ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(size=9))])

    def children(element):
        return [item.get_object() for item in element['/K'] if not isinstance(item, int)]

    item = ldm.Cell(paragraphs=[paragraph('VALUE')])
    row = ldm.Row(cells=[item, item])
    grid = ldm.Table(rows=[row, row])
    doc = model()
    if kind == 'empty':
        grid = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell()])])
        doc.sections[0].body.children = [grid]
    elif kind == 'image':
        buffer = BytesIO()
        Image.new('RGB', (12, 12), 'blue').save(buffer, format='PNG')
        image = ldm.Shape(has_image=True, width=20, height=20, alternative_text='BLUE SQUARE',
                          image_data=ldm.ImageData(image_bytes=buffer.getvalue()))
        grid = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[ldm.Paragraph(children=[image])])])])
        doc.sections[0].body.children = [grid]
    elif kind == 'aliases':
        same_paragraph = paragraph('DUPLICATE')
        item.paragraphs = [same_paragraph, same_paragraph]
        doc.sections[0].body.children = [grid, paragraph('BETWEEN'), grid]
    elif kind == 'floating':
        grid.text_wrapping = 1
        grid._tblp_pr_attrs = {'tblpXSpec': 'right', 'tblpY': '0'}
        doc.sections[0].body.children = [grid]
    else:
        header_grid = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph('TABLEHEADER')])])])
        doc.sections[0].headers_footers = [ldm.HeaderFooter(header_footer_type=0, children=[header_grid])]
        long_cell = ldm.Cell(paragraphs=[paragraph('\n'.join(f'LINE{i:03}' for i in range(40)))])
        grid = ldm.Table(rows=[ldm.Row(cells=[long_cell])])
        doc.sections[0].body.children = [grid]
    options = PdfSaveOptions()
    options.export_document_structure = True
    raw = LdmPdfWriter(options).write_to_bytes(doc)
    reader = PdfReader(BytesIO(raw))
    document = reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()
    tables = [child for child in children(document) if child['/S'] == '/Table']
    if kind == 'empty':
        assert len(tables) == 1
        assert children(children(tables[0])[0])[0]['/S'] == '/TD'
    elif kind == 'image':
        figure = children(children(children(children(tables[0])[0])[0])[0])[0]
        assert figure['/S'] == '/Figure' and figure['/Alt'] == 'BLUE SQUARE'
    elif kind == 'aliases':
        assert len(tables) == 2
        for table in tables:
            rows = children(table)
            assert len(rows) == 2
            for row in rows:
                cells = children(row)
                assert len(cells) == 2
                assert all(len(children(cell)) == 2 for cell in cells)
    elif kind == 'floating':
        assert len(tables) == 1 and len(children(tables[0])) == 2
        assert all(row['/S'] == '/TR' for row in children(tables[0]))
    else:
        assert len(tables) == 1
        assert all(table['/P'].indirect_reference == document.indirect_reference for table in tables)
        assert_page_tags(raw)
    with pymupdf.open(stream=raw, filetype='pdf') as actual, \
            pymupdf.open(stream=LdmPdfWriter().write_to_bytes(doc), filetype='pdf') as control:
        assert len(actual) == len(control)
        for page, reference in zip(actual, control, strict=True):
            assert page.get_text('words') == reference.get_text('words')
            assert page.get_pixmap().samples == reference.get_pixmap().samples


def test_native_link_structure_does_not_shift_marked_content_parent_indices():
    from fpdf import FPDF
    from aspose.words_foss.pdf_writer.structure import TableStructureBuilder

    pdf = FPDF()
    pdf.struct_builder = TableStructureBuilder(pdf)
    options = PdfSaveOptions()
    options.export_document_structure = True
    writer = LdmPdfWriter(options)
    for _ in range(2):
        pdf.add_page()
        pdf.set_font('Helvetica', size=12)
        pdf.link(10, 10, 20, 5, 'https://example.com', alt_text='EXAMPLE')
        with writer._tag(pdf, '/P'):
            pdf.text(10, 30, 'BODY')
        with writer._tag(pdf, '/P'):
            pdf.text(10, 40, 'MORE')
    reader = assert_page_tags(bytes(pdf.output()))
    assert all(page['/Annots'][0].get_object()['/A']['/URI'] == 'https://example.com'
               for page in reader.pages)


def merged_table_model(kind, paginate=False, nested=False):
    def cell(text, **formatting):
        return ldm.Cell(paragraphs=[ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(size=9))])],
                        cell_format=ldm.CellFormat(**formatting))

    long_text = '\n'.join(f'LONG{i:03}' for i in range(40)) if paginate else 'RIGHT1'
    if kind in ('horizontal', 'legacy'):
        first = ([cell('MERGED', grid_span=2)] if kind == 'horizontal' else
                 [cell('MERGED', horizontal_merge=1), cell('SECOND', horizontal_merge=2)])
        rows = [ldm.Row(cells=[*first, cell('RIGHT0')], row_format=ldm.RowFormat(heading_format=True)),
                ldm.Row(cells=[cell('A1'), cell('B1'), cell(long_text)]),
                ldm.Row(cells=[cell('A2'), cell('B2'), cell('RIGHT2')])]
        columns, row_span, col_span = 3, 1, 2
    else:
        col_span = 2 if kind == 'combined' else 1
        rows = [ldm.Row(cells=[cell('MERGED', grid_span=col_span, vertical_merge=1), cell('RIGHT0')],
                       row_format=ldm.RowFormat(heading_format=True)),
                ldm.Row(cells=[cell('CONT1', grid_span=col_span, vertical_merge=2), cell(long_text)]),
                ldm.Row(cells=[cell('CONT2', grid_span=col_span, vertical_merge=2), cell('RIGHT2')])]
        columns, row_span = col_span + 1, 3
    grid = ldm.Table(rows=rows)
    body_grid = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(tables=[grid])])]) if nested else grid
    doc = model()
    doc.sections[0].body.children = [body_grid]
    return doc, columns, row_span, col_span


@pytest.mark.parametrize('kind', ['horizontal', 'legacy', 'vertical', 'combined'])
@pytest.mark.parametrize('paginate', [False, True])
def test_merged_cell_attributes_reconstruct_the_logical_grid(kind, paginate):
    doc, columns, row_span, col_span = merged_table_model(kind, paginate)
    snapshot = doc.model_dump()
    options = PdfSaveOptions()
    options.export_document_structure = True
    raw = LdmPdfWriter(options).write_to_bytes(doc)
    reader = assert_page_tags(raw) if paginate else PdfReader(BytesIO(raw))
    table = reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()['/K'][0].get_object()
    rows = [item.get_object() for item in table['/K']]
    assert len(rows) == 3
    anchor = rows[0]['/K'][0].get_object()
    attrs = anchor['/A']
    assert attrs['/O'] == '/Table'
    assert attrs.get('/RowSpan', 1) == row_span
    assert attrs.get('/ColSpan', 1) == col_span
    occupied = set()
    for index, row in enumerate(rows):
        assert row['/S'] == '/TR'
        column = 0
        for reference in row['/K']:
            while (index, column) in occupied:
                column += 1
            cell = reference.get_object()
            assert cell['/P'].indirect_reference == row.indirect_reference
            attributes = cell.get('/A', {})
            width, height = attributes.get('/ColSpan', 1), attributes.get('/RowSpan', 1)
            for y in range(index, index + height):
                for x in range(column, column + width):
                    assert (y, x) not in occupied
                    occupied.add((y, x))
            column += width
    assert occupied == {(y, x) for y in range(3) for x in range(columns)}
    with pymupdf.open(stream=raw, filetype='pdf') as actual, \
            pymupdf.open(stream=LdmPdfWriter().write_to_bytes(doc), filetype='pdf') as control:
        assert len(actual) == len(control)
        for page, reference in zip(actual, control, strict=True):
            assert page.get_text('words') == reference.get_text('words')
            assert page.get_pixmap().samples == reference.get_pixmap().samples
    assert doc.model_dump() == snapshot


def test_nested_combined_merge_keeps_one_anchor_and_continuation_content():
    doc, _, _, _ = merged_table_model('combined', nested=True)
    options = PdfSaveOptions()
    options.export_document_structure = True
    raw = LdmPdfWriter(options).write_to_bytes(doc)
    reader = PdfReader(BytesIO(raw))
    outer = reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()['/K'][0].get_object()
    outer_cell = outer['/K'][0].get_object()['/K'][0].get_object()
    nested = outer_cell['/K'][0].get_object()
    assert nested['/S'] == '/Table'
    rows = [item.get_object() for item in nested['/K']]
    assert [len(row['/K']) for row in rows] == [2, 1, 1]
    anchor = rows[0]['/K'][0].get_object()
    assert anchor['/A']['/RowSpan'] == 3 and anchor['/A']['/ColSpan'] == 2
    assert len(anchor['/K']) == 3
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        text = ''.join(page.get_text() for page in pdf)
        assert all(text.count(token) == 1 for token in ('MERGED', 'CONT1', 'CONT2'))


@pytest.mark.parametrize('invalid', ['orphan', 'width'])
def test_invalid_vertical_continuations_are_not_attached_to_unrelated_cells(invalid):
    doc, _, _, _ = merged_table_model('combined')
    grid = doc.sections[0].body.children[0]
    if invalid == 'orphan':
        grid.rows[0].cells[0].cell_format.vertical_merge = 2
    else:
        grid.rows[1].cells[0].cell_format.grid_span = 1
        grid.rows[1].cells.insert(1, ldm.Cell(paragraphs=[ldm.Paragraph(children=[ldm.Run(text='MIDDLE')])]))
    options = PdfSaveOptions()
    options.export_document_structure = True
    raw = LdmPdfWriter(options).write_to_bytes(doc)
    reader = PdfReader(BytesIO(raw))
    table = reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()['/K'][0].get_object()
    rows = [item.get_object() for item in table['/K']]
    assert [len(row['/K']) for row in rows] == ([2, 2, 2] if invalid == 'orphan' else [2, 3, 2])
    assert all('/RowSpan' not in ref.get_object().get('/A', {}) for row in rows for ref in row['/K'])


def test_merge_attributes_follow_model_changes_between_writer_calls(monkeypatch):
    from aspose.words_foss.pdf_writer import structure

    doc, _, _, _ = merged_table_model('combined', paginate=True)
    original = structure.iter_grid_cells
    scanned = []

    def scan(row):
        scanned.append(row)
        return original(row)

    monkeypatch.setattr(structure, 'iter_grid_cells', scan)
    options = PdfSaveOptions()
    options.export_document_structure = True
    writer = LdmPdfWriter(options)

    def anchor(raw):
        reader = PdfReader(BytesIO(raw))
        table = reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()['/K'][0].get_object()
        return table['/K'][0].get_object()['/K'][0].get_object()

    assert anchor(writer.write_to_bytes(doc))['/A']['/RowSpan'] == 3
    assert len(scanned) == 3
    doc.sections[0].body.children[0].rows[2].cells[0].cell_format.vertical_merge = 0
    assert anchor(writer.write_to_bytes(doc))['/A']['/RowSpan'] == 2
    assert len(scanned) == 6


def test_public_docx_conversion_keeps_combined_cell_spans():
    import aspose.words_foss as aw
    from aspose.words_foss.docx_writer import LdmDocxWriter

    doc, _, _, _ = merged_table_model('combined')
    source = LdmDocxWriter().write_to_bytes(doc)
    document = aw.Document(BytesIO(source))
    options = PdfSaveOptions()
    options.export_document_structure = True
    raw = document.to_bytes(options)
    reader = PdfReader(BytesIO(raw))
    table = reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()['/K'][0].get_object()
    rows = [ref.get_object() for ref in table['/K']]
    assert [len(row['/K']) for row in rows] == [2, 1, 1]
    attrs = rows[0]['/K'][0].get_object()['/A']
    assert attrs['/RowSpan'] == 3 and attrs['/ColSpan'] == 2
    with pymupdf.open(stream=raw, filetype='pdf') as actual, \
            pymupdf.open(stream=document.to_bytes(PdfSaveOptions()), filetype='pdf') as control:
        assert len(actual) == len(control)
        for page, reference in zip(actual, control, strict=True):
            assert page.get_text('words') == reference.get_text('words')
            assert page.get_pixmap().samples == reference.get_pixmap().samples


@pytest.mark.parametrize('style', ['Normal', 'Heading', 'Quote', 'Code'])
@pytest.mark.parametrize('bands', [False, True])
@pytest.mark.parametrize('columns', [1, 2])
def test_cross_page_body_has_one_logical_element(style, bands, columns):
    options = PdfSaveOptions()
    options.export_document_structure = True
    reader = PdfReader(BytesIO(LdmPdfWriter(options).write_to_bytes(model(style, bands, columns))))
    document = reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()
    assert len(document['/K']) == 1
    element = document['/K'][0].get_object()
    assert element['/S'] == {'Normal':'/P', 'Heading':'/H1', 'Quote':'/BlockQuote', 'Code':'/Code'}[style]
    assert len(element['/K']) == len(reader.pages)
    parents = reader.trailer['/Root']['/StructTreeRoot']['/ParentTree']['/Nums']
    arrays = {int(parents[i]):parents[i+1].get_object() for i in range(0,len(parents),2)}
    for page, content in zip(reader.pages, element['/K'], strict=True):
        assert content['/Type'] == '/MCR'
        assert content['/Pg'].indirect_reference == page.indirect_reference
        target = arrays[int(page['/StructParents'])][int(content['/MCID'])]
        assert target.indirect_reference == element.indirect_reference


def test_reused_body_paragraph_is_two_logical_occurrences():
    doc = model()
    para = doc.sections[0].body.children[0]
    doc.sections[0].body.children.append(para)
    options = PdfSaveOptions(); options.export_document_structure = True
    reader = assert_page_tags(LdmPdfWriter(options).write_to_bytes(doc))
    roots = reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()['/K']
    assert len(roots) == 2 and roots[0].indirect_reference != roots[1].indirect_reference
    assert all(len(root.get_object()['/K']) > 1 for root in roots)


@pytest.mark.parametrize('failure', ['drawing', 'closing'])
def test_cross_page_continuation_failure_preserves_output_and_recovers(tmp_path, monkeypatch, failure):
    from fpdf import FPDF
    from aspose.words_foss.pdf_writer.run_renderer import RunRenderer

    options = PdfSaveOptions(); options.export_document_structure = True
    writer = LdmPdfWriter(options)
    output = tmp_path / 'keep.pdf'; output.write_bytes(b'KEEP')
    captured, files = [], []
    original_render = RunRenderer._render_segment_row
    original_out = FPDF._out

    def render(renderer, pdf, *args, **kwargs):
        if pdf.page_no() >= 3:
            captured.append(pdf)
            files.extend(font.ttfont.reader.file for font in pdf.fonts.values() if font.ttfont.reader is not None)
            if failure == 'drawing':
                raise RuntimeError('continuation drawing failed')
        return original_render(renderer, pdf, *args, **kwargs)

    def close(pdf, text):
        if failure == 'closing' and text == 'EMC' and pdf.page_no() >= 3:
            raise RuntimeError('continuation closing failed')
        return original_out(pdf, text)

    monkeypatch.setattr(RunRenderer, '_render_segment_row', render)
    monkeypatch.setattr(FPDF, '_out', close)
    with pytest.raises(RuntimeError, match=f'continuation {failure} failed'):
        writer.write(model(), output)
    assert output.read_bytes() == b'KEEP'
    assert captured and files and all(file.closed for file in files)
    assert 'add_page' not in captured[0].__dict__
    monkeypatch.setattr(RunRenderer, '_render_segment_row', original_render)
    monkeypatch.setattr(FPDF, '_out', original_out)
    reader = assert_page_tags(writer.write_to_bytes(model()))
    assert len(reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()['/K']) == 1


def test_public_docx_cross_page_body_is_one_logical_paragraph():
    import aspose.words_foss as aw
    from aspose.words_foss.docx_writer import LdmDocxWriter

    options = PdfSaveOptions(); options.export_document_structure = True
    document = aw.Document(BytesIO(LdmDocxWriter().write_to_bytes(model(bands=True, columns=2))))
    reader = assert_page_tags(document.to_bytes(options))
    root = reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()
    assert len(root['/K']) == 1 and root['/K'][0].get_object()['/S'] == '/P'
    assert len(root['/K'][0].get_object()['/K']) == len(reader.pages)
