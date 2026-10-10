"""Self-generated textbox layout and source/JSON/package regressions."""
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from aspose.words_foss import Document, SaveFormat, light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from tests.test_drawing_line_joins import join_docx

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WPS = '{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def textbox_docx(size=12, line=24, anchor='t', text='ROTATED'):
    with ZipFile(BytesIO(join_docx('round'))) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    root = ET.fromstring(parts['word/document.xml'])
    wsp = root.find('.//' + WPS + 'wsp')
    wsp.find(WPS + 'spPr/' + A + 'ln').set('w', str(8 * 12700))
    paragraph = wsp.find('.//' + W + 'txbxContent/' + W + 'p')
    paragraph.clear()
    properties = ET.SubElement(paragraph, W + 'pPr')
    ET.SubElement(properties, W + 'spacing', {
        W + 'before': '0', W + 'after': '0', W + 'line': str(line * 20), W + 'lineRule': 'exact'})
    run = ET.SubElement(paragraph, W + 'r')
    run_properties = ET.SubElement(run, W + 'rPr')
    ET.SubElement(run_properties, W + 'rFonts', {
        W + 'ascii': 'Document Sans SC', W + 'hAnsi': 'Document Sans SC'})
    ET.SubElement(run_properties, W + 'sz', {W + 'val': str(size * 2)})
    ET.SubElement(run, W + 't').text = text
    body = wsp.find(WPS + 'bodyPr')
    body.set('anchor', anchor)
    for key in ('lIns', 'tIns', 'rIns', 'bIns'):
        body.set(key, '72000')
    parts['word/document.xml'] = ET.tostring(root)
    output = BytesIO()
    with ZipFile(output, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return output.getvalue()


@pytest.mark.parametrize('size,line,anchor,text,rows', [
    (12, 8, 't', 'ROTATED', 1), (20, 12, 't', 'ROTATED', 1),
    (8, 24, 't', 'ROTATED', 1),
    (12, 24, 'ctr', 'ROTATED', 1), (12, 24, 'b', 'ROTATED', 1),
    (12, 12, 'ctr', 'ROTATED ROTATED ROTATED', 3),
    (12, 12, 'b', 'ROTATED ROTATED ROTATED', 3),
])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_exact_textbox_json_package_baseline_and_after(tmp_path, size, line, anchor, text, rows, format):
    fitz = pytest.importorskip('pymupdf')
    owner = Document(BytesIO(textbox_docx(size, line, anchor, text)))
    model = ldm.Document.model_validate_json(owner.light_document_model.model_dump_json())
    shape = model.get_child_nodes(ldm.NodeType.SHAPE, True)[0]
    shape.stroke.color = '#0080C0'
    data = LdmDocxWriter().write_to_bytes(model)
    owner = Document(BytesIO(data))
    output = tmp_path / 'textbox'
    owner.save(output, format)
    cold = Document(output)
    before = cold.light_document_model.model_dump_json()
    page = fitz.open(stream=cold.to_bytes(SaveFormat.PDF), filetype='pdf')[0]
    spans = [(span, row['dir']) for block in page.get_text('dict')['blocks'] if 'lines' in block
             for row in block['lines'] for span in row['spans']]
    rotated = [span for span, _ in spans if 'ROTATED' in span['text']]
    assert len(rotated) == rows
    pt_per_mm = 72 / 25.4
    pad = 2 * pt_per_mm + 4
    inner_height = 20 * pt_per_mm - 2 * pad
    offset = max(0, inner_height - rows * line) * {'t': 0, 'ctr': .5, 'b': 1}[anchor]
    for index, span in enumerate(rotated):
        assert span['origin'][0] == pytest.approx(60 * pt_per_mm + pad, abs=.01)
        assert span['origin'][1] == pytest.approx(60 * pt_per_mm + pad + offset + .8 * line + index * line, abs=.01)
    assert any(span['text'] == 'AFTER' and direction == (1, 0) for span, direction in spans)
    assert cold.light_document_model.model_dump_json() == before
    assert cold.light_document_model.get_child_nodes(ldm.NodeType.SHAPE, True)[0].text_box['insets_mm'] == (2, 2, 2, 2)


def test_textbox_without_outline_keeps_source_and_following_context():
    fitz = pytest.importorskip('pymupdf')
    owner = Document(BytesIO(textbox_docx()))
    shape = owner.light_document_model.get_child_nodes(ldm.NodeType.SHAPE, True)[0]
    shape.stroke = None
    before = owner.light_document_model.model_dump_json()
    page = fitz.open(stream=owner.to_bytes(SaveFormat.PDF), filetype='pdf')[0]
    spans = [span for block in page.get_text('dict')['blocks'] if 'lines' in block
             for row in block['lines'] for span in row['spans']]
    text = next(span for span in spans if span['text'] == 'ROTATED')
    assert text['origin'][0] == pytest.approx(62 * 72 / 25.4, abs=.01)
    assert text['origin'][1] == pytest.approx(62 * 72 / 25.4 + .8 * 24, abs=.01)
    assert any(span['text'] == 'AFTER' for span in spans)
    assert owner.light_document_model.model_dump_json() == before
