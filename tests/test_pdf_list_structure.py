"""List tags must preserve item relationships without changing layout."""
from io import BytesIO

import pymupdf
from pypdf import PdfReader
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.saving import PdfSaveOptions
from tests.test_pdf_structure_pages import assert_page_tags


def list_model(kind='nested', columns=1, long=False):
    levels = {'flat':[0,0,0], 'nested':[0,1,2,1,0], 'gaps':[2,4,3,2],
              'interrupt':[0,0,0], 'different':[0,0,0]}[kind]
    children=[]
    for index, level in enumerate(levels):
        if kind == 'interrupt' and index == 2:
            children.append(ldm.Paragraph(children=[ldm.Run(text='BREAK',font=ldm.Font(size=9))]))
        text=f'ITEM{index}'
        if long and index == 0:
            text+='\n'+'\n'.join(f'LONG{i:02}' for i in range(30))
        children.append(ldm.Paragraph(children=[ldm.Run(text=text,font=ldm.Font(size=9))],
            list_format=ldm.ListFormat(is_list_item=True,list_level_number=level,
                                      list_id=2 if kind == 'different' and index == 2 else 1),
            list_label=ldm.ListLabel(label_string=f'{index+1}.')))
    return ldm.Document(sections=[ldm.Section(page_setup=ldm.PageSetup(page_width=280,page_height=180,
        left_margin=20,right_margin=20,top_margin=25,bottom_margin=25,
        text_columns=ldm.TextColumns(count=columns,spacing=15)),body=ldm.Body(children=children))])


def structure_children(node):
    return [kid.get_object() for kid in node['/K'] if not isinstance(kid,int) and kid.get_object().get('/Type') == '/StructElem']


@pytest.mark.parametrize('kind',['flat','nested','gaps','interrupt','different'])
@pytest.mark.parametrize('columns',[1,2])
@pytest.mark.parametrize('long',[False,True])
def test_list_hierarchy_labels_bodies_and_layout(kind,columns,long):
    doc=list_model(kind,columns,long);snapshot=doc.model_dump()
    options=PdfSaveOptions();options.export_document_structure=True
    raw=LdmPdfWriter(options).write_to_bytes(doc)
    reader=PdfReader(BytesIO(raw));document=reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object()
    roots=structure_children(document)
    assert [node['/S'] for node in roots] == (['/L','/P','/L'] if kind == 'interrupt' else ['/L','/L'] if kind == 'different' else ['/L'])
    items=[]; depths=[]
    def walk_list(node,depth):
        assert node['/S']=='/L'
        for item in structure_children(node):
            assert item['/S']=='/LI';items.append(item);depths.append(depth)
            kids=structure_children(item)
            assert [kid['/S'] for kid in kids]==['/Lbl','/LBody']
            assert all(kid['/P'].indirect_reference==item.indirect_reference for kid in kids)
            for nested in structure_children(kids[1]):walk_list(nested,depth+1)
    for root in roots:
        if root['/S']=='/L':walk_list(root,0)
    assert depths=={'flat':[0,0,0],'nested':[0,1,2,1,0],'gaps':[0,1,1,0],
                    'interrupt':[0,0,0],'different':[0,0,0]}[kind]
    with pymupdf.open(stream=raw,filetype='pdf') as actual, \
            pymupdf.open(stream=LdmPdfWriter().write_to_bytes(doc),filetype='pdf') as control:
        assert len(actual)==len(control)
        for page,reference in zip(actual,control,strict=True):
            assert page.get_text('words')==reference.get_text('words')
            assert page.get_pixmap().samples==reference.get_pixmap().samples
        text=''.join(page.get_text() for page in actual)
        assert all(text.count(f'ITEM{i}')==1 for i in range(len(items)))
    if long:assert_page_tags(raw)
    assert doc.model_dump()==snapshot


@pytest.mark.parametrize('boundary',['table','heading','code'])
def test_non_list_rendering_ends_list_even_when_source_has_list_format(boundary):
    doc=list_model('flat')
    if boundary=='table':
        middle=ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[ldm.Paragraph(children=[ldm.Run(text='MIDDLE')])])])])
    else:
        middle=ldm.Paragraph(children=[ldm.Run(text='MIDDLE')],
            paragraph_format=ldm.ParagraphFormat(is_heading=boundary=='heading',
                outline_level=0 if boundary=='heading' else 9,style_name='Code' if boundary=='code' else ''),
            list_format=ldm.ListFormat(is_list_item=True,list_id=1))
    doc.sections[0].body.children.insert(1,middle)
    options=PdfSaveOptions();options.export_document_structure=True
    reader=PdfReader(BytesIO(LdmPdfWriter(options).write_to_bytes(doc)))
    roots=structure_children(reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object())
    assert [node['/S'] for node in roots]==['/L',{'table':'/Table','heading':'/H1','code':'/Code'}[boundary],'/L']
    assert [len(structure_children(node)) for node in (roots[0],roots[-1])]==[1,2]


@pytest.mark.parametrize('long',[False,True])
def test_public_docx_list_hierarchy_survives_conversion(long):
    import aspose.words_foss as aw
    from aspose.words_foss.docx_writer import LdmDocxWriter

    doc=list_model('nested',2,long)
    doc.lists=[ldm.DocList(list_id=1,is_multi_level=True,
        list_levels=[ldm.ListLevel(number_format=f'%{index+1}.') for index in range(3)])]
    document=aw.Document(BytesIO(LdmDocxWriter().write_to_bytes(doc)))
    options=PdfSaveOptions();options.export_document_structure=True
    raw=document.to_bytes(options)
    reader=PdfReader(BytesIO(raw))
    roots=structure_children(reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object())
    assert len(roots)==1 and roots[0]['/S']=='/L'
    roles=[]
    def walk(node):
        roles.append(node['/S'])
        for child in structure_children(node):walk(child)
    walk(roots[0])
    assert roles.count('/L')==3 and roles.count('/LI')==roles.count('/Lbl')==roles.count('/LBody')==5
    with pymupdf.open(stream=raw,filetype='pdf') as pdf:
        text=''.join(page.get_text() for page in pdf)
        assert all(text.count(f'ITEM{i}')==1 for i in range(5))
    if long:assert_page_tags(raw)


def test_list_labels_and_body_text_are_separate_content_items():
    from pypdf import PdfWriter

    options=PdfSaveOptions();options.export_document_structure=True
    raw=LdmPdfWriter(options).write_to_bytes(list_model('flat'))
    def extract_role(role):
        writer=PdfWriter(clone_from=PdfReader(BytesIO(raw)))
        nums=writer.root_object['/StructTreeRoot']['/ParentTree']['/Nums']
        arrays={int(nums[i]):nums[i+1].get_object() for i in range(0,len(nums),2)}
        for page in writer.pages:
            content=page.get_contents();stack=[];kept=[]
            for args,op in content.operations:
                if op==b'BDC':
                    stack.append(arrays[int(page['/StructParents'])][int(args[1]['/MCID'])].get_object()['/S'])
                elif op==b'EMC':stack.pop()
                if op not in (b'Tj',b'TJ') or stack and stack[-1]==role:kept.append((args,op))
            content.operations=kept;page.replace_contents(content)
        stream=BytesIO();writer.write(stream)
        return ''.join(page.extract_text() for page in PdfReader(stream).pages)
    labels=extract_role('/Lbl');body=extract_role('/LBody')
    assert all(f'{i+1}.' in labels and f'ITEM{i}' not in labels for i in range(3))
    assert all(f'ITEM{i}' in body and f'{i+1}.' not in body for i in range(3))


@pytest.mark.parametrize('failure',['drawing','closing'])
def test_list_structure_failure_preserves_output_and_writer_recovers(tmp_path,monkeypatch,failure):
    from fpdf import FPDF
    from aspose.words_foss.pdf_writer.run_renderer import RunRenderer

    options=PdfSaveOptions();options.export_document_structure=True
    writer=LdmPdfWriter(options);output=tmp_path/'keep.pdf';output.write_bytes(b'KEEP')
    captured=[];files=[]
    original_render=RunRenderer._render_segment_row;original_out=FPDF._out
    def render(renderer,pdf,*args,**kwargs):
        captured.append(pdf)
        files.extend(font.ttfont.reader.file for font in pdf.fonts.values() if font.ttfont.reader is not None)
        if failure=='drawing':raise RuntimeError('list drawing failed')
        return original_render(renderer,pdf,*args,**kwargs)
    def close(pdf,text):
        if failure=='closing' and text=='EMC' and captured:raise RuntimeError('list closing failed')
        return original_out(pdf,text)
    monkeypatch.setattr(RunRenderer,'_render_segment_row',render);monkeypatch.setattr(FPDF,'_out',close)
    with pytest.raises(RuntimeError,match=f'list {failure} failed'):writer.write(list_model(),output)
    assert output.read_bytes()==b'KEEP' and files and all(file.closed for file in files)
    pdf=captured[0]
    assert 'add_page' not in pdf.__dict__
    assert all(parent is pdf.struct_builder.doc_struct_elem for parent in pdf.struct_builder.parents.values())
    monkeypatch.setattr(RunRenderer,'_render_segment_row',original_render);monkeypatch.setattr(FPDF,'_out',original_out)
    for kind,count in [('flat',1),('interrupt',2),('nested',1)]:
        reader=PdfReader(BytesIO(writer.write_to_bytes(list_model(kind))))
        roots=structure_children(reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object())
        assert sum(node['/S']=='/L' for node in roots)==count


def test_unlabeled_nested_items_keep_body_and_nesting():
    doc=list_model('nested')
    for para in doc.sections[0].body.children:para.list_label.label_string=''
    options=PdfSaveOptions();options.export_document_structure=True
    raw=LdmPdfWriter(options).write_to_bytes(doc)
    reader=PdfReader(BytesIO(raw))
    root=structure_children(reader.trailer['/Root']['/StructTreeRoot']['/K'][0].get_object())[0]
    items=[]
    def walk(node):
        for item in structure_children(node):
            items.append(item);kids=structure_children(item)
            assert [kid['/S'] for kid in kids]==['/LBody']
            for nested in structure_children(kids[0]):walk(nested)
    walk(root);assert len(items)==5
    with pymupdf.open(stream=raw,filetype='pdf') as pdf:
        text=''.join(page.get_text() for page in pdf)
        assert all(text.count(f'ITEM{i}')==1 for i in range(5))
