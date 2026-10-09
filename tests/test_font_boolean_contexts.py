"""Conversion Boolean values follow owned native category observations."""

import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from aspose.words_foss import Document
from aspose.words_foss import light_document_model as ldm

ROOT = Path(__file__).parents[1] / 'docs/benchmarks'
REPORT = json.loads((ROOT / 'font-boolean-contexts-26.9.json').read_text())
PDF_REPORT = json.loads((ROOT / 'font-hidden-rendering-26.9.json').read_text())
PDF_VISIBLE = {r['input']: r['import_visible'] for r in PDF_REPORT['records']}
STYLE_REPORT = json.loads((ROOT / 'hidden-style-contexts-26.9.json').read_text())


@pytest.mark.parametrize('row', REPORT['records'], ids=[r['input'] for r in REPORT['records']])
def test_boolean_category_values_match_native_before_and_after_json(row):
    with ZipFile(ROOT / REPORT['corpus']) as archive:
        source = archive.read(row['input'])
    model = Document(BytesIO(source)).light_document_model
    for current in (model, ldm.Document.model_validate_json(model.model_dump_json())):
        font = current.sections[0].body.paragraphs[0].runs[0].font
        assert getattr(font, row['field']) is row['value']


@pytest.mark.parametrize('row', [r for r in REPORT['records'] if r['field'] == 'hidden'],
                         ids=[r['input'] for r in REPORT['records'] if r['field'] == 'hidden'])
@pytest.mark.parametrize('json_roundtrip', [False, True])
def test_hidden_visibility_matches_native_pdf(row, json_roundtrip):
    import pymupdf

    from aspose.words_foss.pdf_writer import LdmPdfWriter

    with ZipFile(ROOT / REPORT['corpus']) as archive:
        source = archive.read(row['input'])
    doc = Document(BytesIO(source))
    if json_roundtrip:
        model = ldm.Document.model_validate_json(doc.light_document_model.model_dump_json())
        pdf = LdmPdfWriter().write_to_bytes(model)
    else:
        pdf = doc.to_bytes('pdf')
    with pymupdf.open(stream=pdf, filetype='pdf') as document:
        text = ''.join(page.get_text() for page in document)
        dark = sum(any(pixel < 240 for pixel in page.get_pixmap().samples) for page in document)
    assert ('IMPORT' in text) is PDF_VISIBLE[row['input']]
    assert bool(dark) is PDF_VISIBLE[row['input']]


@pytest.mark.parametrize('hidden_text', ['SECRET', '\f', 'SECRET\f', '\tSECRET\n' * 50])
@pytest.mark.parametrize('keep', [False, True])
def test_hidden_content_does_not_change_visible_layout(hidden_text, keep):
    import pymupdf

    from aspose.words_foss.pdf_writer import LdmPdfWriter

    def render(secret):
        paragraph = ldm.Paragraph(
            paragraph_format=ldm.ParagraphFormat(keep_together=keep),
            children=[ldm.Run(text=secret, font=ldm.Font(hidden=True, size=100)),
                      ldm.Run(text='VISIBLE', font=ldm.Font(size=12))])
        model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[paragraph]))])
        with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(model), filetype='pdf') as pdf:
            return [(page.get_text('words'), page.get_pixmap().samples) for page in pdf]

    assert render(hidden_text) == render('')


@pytest.mark.parametrize('value', [False, True])
def test_direct_hidden_edit_replaces_inherited_rendering_state(value):
    with ZipFile(ROOT / REPORT['corpus']) as archive:
        model = Document(BytesIO(archive.read('hidden-1-0-1-n.docx'))).light_document_model
    font = model.sections[0].body.paragraphs[0].runs[0].font
    assert font.hidden is True and font.render_hidden is False
    font = ldm.Font.model_validate_json(font.model_dump_json())
    assert font.hidden is True and font.render_hidden is False
    copy = font.model_copy(update={'hidden': value})
    assert copy.render_hidden is value
    font.hidden = value
    assert font.hidden_rendering is None and font.render_hidden is value


@pytest.mark.parametrize('target', ['font-boolean-contexts-26.9.json', 'font-hidden-rendering-26.9.json'])
def test_independent_verifier_rejects_changed_observations(tmp_path, target):
    import runpy
    import shutil

    check = runpy.run_path(str(ROOT.parents[1] / 'scripts/verify_commercial_baseline.py'))['verify_font_boolean_contexts']
    (tmp_path / 'corpus').mkdir()
    for name in ('font-boolean-contexts-26.9.json', 'font-hidden-rendering-26.9.json',
                 REPORT['corpus'], PDF_REPORT['outputs']):
        shutil.copyfile(ROOT / name, tmp_path / name)
    report = json.loads((tmp_path / target).read_text())
    key = 'value' if target.startswith('font-boolean') else 'import_visible'
    report['records'][0][key] = not report['records'][0][key]
    (tmp_path / target).write_text(json.dumps(report))
    with pytest.raises(AssertionError):
        check(tmp_path)


@pytest.mark.parametrize('orientation', [0, 1, 2])
@pytest.mark.parametrize('render_hidden', [False, True])
def test_table_visibility_and_supported_inline_image(orientation, render_hidden):
    import pymupdf
    from PIL import Image

    from aspose.words_foss.pdf_writer import LdmPdfWriter

    image = BytesIO()
    Image.new('RGB', (4, 4), 'red').save(image, format='PNG')
    paragraph = ldm.Paragraph(children=[
        ldm.Run(text='TARGET', font=ldm.Font(hidden=True, hidden_rendering=render_hidden, size=40)),
        ldm.Shape(has_image=True, is_inline=True, width=12, height=12,
                  image_data=ldm.ImageData(image_type=ldm.ImageData.from_mime('image/png'), image_bytes=image.getvalue())),
        ldm.Run(text='CONTROL', font=ldm.Font(size=12))])
    if orientation:
        paragraph._children.pop(1)
    cell = ldm.Cell(paragraphs=[paragraph], cell_format=ldm.CellFormat(orientation=orientation))
    table = ldm.Table(rows=[ldm.Row(cells=[cell])])
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[table]))])
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(model), filetype='pdf') as pdf:
        text = ''.join(page.get_text() for page in pdf)
        assert ('TARGET' in text) is not render_hidden
        assert 'CONTROL' in text
        if render_hidden:
            control = next(span for page in pdf for block in page.get_text('dict')['blocks']
                           for line in block.get('lines', []) for span in line['spans']
                           if 'CONTROL' in span['text'])
            assert control['size'] == pytest.approx(12, abs=0.01)


@pytest.mark.parametrize('row', STYLE_REPORT['records'], ids=[r['input'] for r in STYLE_REPORT['records']])
@pytest.mark.parametrize('json_roundtrip', [False, True])
def test_default_paragraph_reference_presence_matches_native(row, json_roundtrip):
    import pymupdf

    from aspose.words_foss.pdf_writer import LdmPdfWriter

    with ZipFile(ROOT / STYLE_REPORT['corpus']) as archive:
        doc = Document(BytesIO(archive.read(row['input'])))
    model = doc.light_document_model
    if json_roundtrip:
        model = ldm.Document.model_validate_json(model.model_dump_json())
    font = model.sections[0].body.paragraphs[0].runs[0].font
    assert font.hidden is row['hidden']
    assert font.render_hidden is not row['visible']
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(model), filetype='pdf') as pdf:
        assert any('IMPORT' in page.get_text() for page in pdf) is row['visible']
        assert any(any(pixel < 240 for pixel in page.get_pixmap().samples) for page in pdf) is row['visible']


@pytest.mark.parametrize('reverse', [False, True])
def test_inherited_font_cache_distinguishes_implicit_paragraph_style(reverse):
    from copy import deepcopy
    from xml.etree import ElementTree as ET

    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    with ZipFile(ROOT / STYLE_REPORT['corpus']) as archive:
        raw = archive.read('implicit-default-0-1-0-1.docx')
    source = BytesIO()
    expected = [False, True, False, True]
    if reverse:
        expected.reverse()
    with ZipFile(BytesIO(raw)) as original, ZipFile(source, 'w') as modified:
        for name in original.namelist():
            data = original.read(name)
            if name == 'word/document.xml':
                xml = ET.fromstring(data)
                body = xml.find(w + 'body')
                para = body.find(w + 'p')
                body.remove(para)
                for index, hidden in enumerate(expected):
                    current = deepcopy(para)
                    if hidden:
                        ET.SubElement(current.find(w + 'pPr'), w + 'pStyle', {w + 'val': 'P'})
                    body.insert(index, current)
                data = ET.tostring(xml)
            modified.writestr(name, data)
    model = Document(BytesIO(source.getvalue())).light_document_model
    fonts = [p.runs[0].font for p in model.sections[0].body.paragraphs]
    assert all(font.hidden for font in fonts)
    assert [font.render_hidden for font in fonts] == expected


def test_implicit_style_verifier_rejects_visible_forgery(tmp_path):
    import runpy
    import shutil

    check = runpy.run_path(str(ROOT.parents[1] / 'scripts/verify_commercial_baseline.py'))['verify_hidden_style_contexts']
    (tmp_path / 'corpus').mkdir()
    for name in ('hidden-style-contexts-26.9.json', STYLE_REPORT['corpus'], STYLE_REPORT['outputs']):
        shutil.copyfile(ROOT / name, tmp_path / name)
    report = json.loads((tmp_path / 'hidden-style-contexts-26.9.json').read_text())
    report['records'][0]['visible'] = not report['records'][0]['visible']
    (tmp_path / 'hidden-style-contexts-26.9.json').write_text(json.dumps(report))
    with pytest.raises(AssertionError):
        check(tmp_path)


SAVE_ROWS = ([r for r in REPORT['records'] if r['field'] == 'hidden'] + STYLE_REPORT['records'])


@pytest.mark.parametrize('row', SAVE_ROWS, ids=[r['input'] for r in SAVE_ROWS])
@pytest.mark.parametrize('format', ['docx', 'flat_opc'])
@pytest.mark.parametrize('json_roundtrip', [False, True])
def test_hidden_save_retains_direct_setting_and_paragraph_reference(row, format, json_roundtrip):
    import runpy

    from aspose.words_foss import SaveFormat
    from aspose.words_foss.docx_writer import LdmDocxWriter
    from aspose.words_foss.saving import OoxmlSaveOptions

    current_report = STYLE_REPORT if row['input'].startswith(('explicit-', 'implicit-')) else REPORT
    with ZipFile(ROOT / current_report['corpus']) as corpus:
        raw = corpus.read(row['input'])
    model = Document(BytesIO(raw)).light_document_model
    if json_roundtrip:
        model = ldm.Document.model_validate_json(model.model_dump_json())
    data = LdmDocxWriter(OoxmlSaveOptions(SaveFormat.DOCX if format == 'docx' else SaveFormat.FLAT_OPC)).write_to_bytes(model)
    parts = runpy.run_path(str(ROOT.parents[1] / 'scripts/verify_commercial_baseline.py'))['font_size_parts']
    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    source, output = (parts(value)[0] for value in (raw, data))
    for path in ('.//' + w + 'p/' + w + 'pPr/' + w + 'pStyle', './/' + w + 'r/' + w + 'rPr/' + w + 'vanish'):
        before, after = source.find(path), output.find(path)
        assert (before is None) == (after is None)
        if before is not None:
            assert before.get(w + 'val', '1') == after.get(w + 'val', '1')
    reopened = Document(BytesIO(data)).light_document_model.sections[0].body.paragraphs[0].runs[0].font
    assert reopened.hidden is row.get('value', row.get('hidden'))
    assert reopened.render_hidden is not (row.get('visible') if 'visible' in row else PDF_VISIBLE[row['input']])


@pytest.mark.parametrize('copy', [False, True])
def test_direct_paragraph_style_edit_clears_implicit_origin(copy):
    pf = ldm.ParagraphFormat(style_name='P', style_explicit=False)
    pf = ldm.ParagraphFormat.model_validate_json(pf.model_dump_json())
    if copy:
        pf = pf.model_copy(update={'style_name': 'P'})
    else:
        pf.style_name = 'P'
    assert pf.style_explicit is None


def test_saved_hidden_verifier_rejects_native_visibility_forgery(tmp_path):
    import runpy
    import shutil

    check = runpy.run_path(str(ROOT.parents[1] / 'scripts/verify_commercial_baseline.py'))['verify_hidden_font_roundtrip']
    report = json.loads((ROOT / 'hidden-font-roundtrip-26.9.json').read_text())
    (tmp_path / 'corpus').mkdir()
    for name in ('font-boolean-contexts-26.9.json', REPORT['corpus'],
                 'font-hidden-rendering-26.9.json', 'hidden-style-contexts-26.9.json', STYLE_REPORT['corpus'],
                 report['outputs'], report['before_origin_fix']['outputs']):
        shutil.copyfile(ROOT / name, tmp_path / name)
    row = report['native_reread']['records'][0]
    row['visible'] = not row['visible']
    (tmp_path / 'hidden-font-roundtrip-26.9.json').write_text(json.dumps(report))
    with pytest.raises(AssertionError):
        check(tmp_path)
