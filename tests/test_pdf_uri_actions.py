"""PDF URI projection preserves the document's external relationship targets."""
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from pypdf import PdfReader
import pytest

from aspose.words_foss import Document, light_document_model as ldm
from aspose.words_foss._flat_opc import encode
from aspose.words_foss._links import format_link
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.pdf_writer.writer import _PdfCanvas
from aspose.words_foss.saving import PdfSaveOptions


@pytest.mark.parametrize('target,action', [
    ('http://one.test#', 'http://one.test/'),
    ('http://one.test/path#', 'http://one.test/path'),
    ('https://one.test#', 'https://one.test/'),
    ('https://one.test/path#', 'https://one.test/path'),
    ('https://one.test/path?x=1#', 'https://one.test/path?x=1'),
    ('https://one.test/path#anchor', 'https://one.test/path#anchor'),
    ('https://one.test', 'https://one.test/'),
    ('https://one.test?', 'https://one.test/?'),
    ('https://one.test?#', 'https://one.test/?'),
    ('https://%E4%BE%8B%E5%AD%90.test/a#', 'https://%E4%BE%8B%E5%AD%90.test/a'),
    ('https://%E4%BE%8B%E5%AD%90.test/a', 'https://%E4%BE%8B%E5%AD%90.test/a'),
    ('mailto:user@example.test#', 'mailto:user@example.test'),
    ('relative/path#', 'relative/path'),
    ('https://one.test/path##', 'https://one.test/path##'),
    ('https://one.test/path%23', 'https://one.test/path%23'),
    ('https://one.test/path#%23', 'https://one.test/path#%23'),
])
def test_pdf_action_projects_empty_fragment_without_decoding_source(target, action):
    pdf = _PdfCanvas(unit='pt', format=(300, 400)); pdf.add_page()
    pdf.link(10, 20, 30, 40, target, alt_text='LINK')
    item = PdfReader(BytesIO(bytes(pdf.output()))).pages[0]['/Annots'][0].get_object()
    assert item['/A']['/URI'] == action
    assert list(map(float, item['/Rect'])) == [10, 340, 40, 380]


@pytest.mark.parametrize('target,action', [
    ('https://one.test#', 'https://one.test/'),
    ('https://one.test/path#anchor', 'https://one.test/path#anchor'),
    ('https://one.test?#', 'https://one.test/?'),
    ('https://%E4%BE%8B%E5%AD%90.test/a', 'https://%E4%BE%8B%E5%AD%90.test/a'),
    ('https://one.test/path##', 'https://one.test/path##'),
])
@pytest.mark.parametrize('shaping', [False, True])
@pytest.mark.parametrize('in_table', [False, True])
def test_uri_projection_is_pdf_only_across_json_docx_and_flat(target, action, shaping, in_table):
    paragraph = ldm.Paragraph(children=[ldm.Run(text=format_link('LINK', target), is_hyperlink=True)])
    child = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(children=[paragraph])])]) if in_table else paragraph
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[child]))])
    original = model.model_dump_json()
    package = LdmDocxWriter().write_to_bytes(model)
    with ZipFile(BytesIO(package)) as z:
        targets = [el.get('Target') for el in ET.fromstring(z.read('word/_rels/document.xml.rels'))
                   if el.get('Type', '').endswith('/hyperlink')]
    assert targets == [target]
    options = PdfSaveOptions(); options.text_shaping = shaping
    for restored in [ldm.Document.model_validate_json(original),
                     Document(BytesIO(package)).light_document_model,
                     Document(BytesIO(encode(package))).light_document_model]:
        before = restored.model_dump_json()
        pdf = PdfReader(BytesIO(LdmPdfWriter(options).write_to_bytes(restored)))
        assert [item.get_object()['/A']['/URI'] for item in pdf.pages[0]['/Annots']] == [action]
        assert restored.model_dump_json() == before
    assert model.model_dump_json() == original
