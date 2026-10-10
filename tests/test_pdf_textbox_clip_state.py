"""Captured clipping must leave the next textbox's font usable."""
from io import BytesIO

import pytest

from aspose.words_foss import Document, SaveFormat, light_document_model as ldm
from aspose.words_foss.model.enums import LineSpacingRule
from tests.test_drawing_textbox_layout import textbox_docx


@pytest.mark.parametrize('format', [None, SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_same_size_text_after_natural_textbox_is_not_lost(tmp_path, format):
    fitz = pytest.importorskip('pymupdf')
    owner = Document(BytesIO(textbox_docx(size=8)))
    shape = owner.light_document_model.get_child_nodes(ldm.NodeType.SHAPE, True)[0]
    pf = shape.text_box['paragraphs'][0].paragraph_format
    pf.line_spacing_rule = LineSpacingRule.MULTIPLE
    pf.line_spacing = 12
    for run in owner.light_document_model.get_child_nodes(ldm.NodeType.RUN, True):
        if run.text == 'AFTER':
            run.font.size = 8
    model = ldm.Document.model_validate_json(owner.light_document_model.model_dump_json())
    from aspose.words_foss.docx_writer import LdmDocxWriter
    owner = Document(BytesIO(LdmDocxWriter().write_to_bytes(model)))
    if format is not None:
        output = tmp_path / 'textbox'
        owner.save(output, format)
        owner = Document(output)
    before = owner.light_document_model.model_dump_json()
    page = fitz.open(stream=owner.to_bytes(SaveFormat.PDF), filetype='pdf')[0]
    spans = [span for block in page.get_text('dict')['blocks'] if 'lines' in block
             for row in block['lines'] for span in row['spans']]
    assert any(span['text'] == 'ROTATED' and span['size'] == 8 for span in spans)
    assert any(span['text'] == 'AFTER' and span['size'] == 8 for span in spans)
    assert owner.light_document_model.model_dump_json() == before
