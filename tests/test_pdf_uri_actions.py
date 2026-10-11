"""PDF URI projection preserves the document's external relationship targets."""
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from pypdf import PdfReader
import pytest

from dockiln import Document, light_document_model as ldm
from dockiln._flat_opc import encode
from dockiln._links import format_link
from dockiln.writers.docx import LdmDocxWriter
from dockiln.writers.pdf import LdmPdfWriter
from dockiln.writers.pdf.writer import _PdfCanvas
from dockiln.saving import PdfSaveOptions


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
    ('file:///tmp/a%23b.txt', 'tmp/a#b.txt'),
    ('file:///tmp/报告.txt', 'tmp/%E6%8A%A5%E5%91%8A.txt'),
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


@pytest.mark.parametrize("target,action", [
    ('file:///tmp/input.txt#', 'tmp/input.txt'),
    ('file:////server/share/Input.txt', '/server/share/Input.txt'),
    ('file:///tmp/report%20name.txt', 'tmp/report%20name.txt'),
    ('file:///tmp/报告.txt', 'tmp/%E6%8A%A5%E5%91%8A.txt'),
    ('file:///tmp/a%23.txt', 'tmp/a#.txt'),
    ('file:///tmp/a.txt#bookmark', 'tmp/a.txt#bookmark'),
    ('file:///tmp/a.txt?x=1#bookmark', 'tmp/a.txt?x=1#bookmark'),
    ('file:///tmp/a.txt##', 'tmp/a.txt##'),
    ('file:///tmp/../a.txt', 'tmp/../a.txt'),
    ('file:///tmp//a.txt', 'tmp//a.txt'),
    ('file:///tmp/a%2Fb.txt', 'tmp/a/b.txt'),
    ('file:///tmp/a%3Fb.txt', 'tmp/a?b.txt'),
    ('file:///tmp/a%25b.txt', 'tmp/a%b.txt'),
    ('file:///tmp/a%5Cb.txt', 'tmp/a/b.txt'),
    ('file:///tmp/a%22b.txt', 'tmp/a%22b.txt'),
    ('file:///tmp/a%2Bb.txt', 'tmp/a+b.txt'),
    ('file:///tmp/a%26b.txt', 'tmp/a&b.txt'),
    ('file:///tmp/a%3Db.txt', 'tmp/a=b.txt'),
    ('file:///tmp/a%3Ab.txt', 'tmp/a:b.txt'),
    ('file:///tmp/a%40b.txt', 'tmp/a@b.txt'),
    ('file:///tmp/a%7Eb.txt', 'tmp/a~b.txt'),
    ('file:///tmp/a%21b.txt', 'tmp/a!b.txt'),
    ('file:///tmp/a%24b.txt', 'tmp/a$b.txt'),
    ('file:///tmp/a%27b.txt', "tmp/a'b.txt"),
    ('file:///tmp/a%28b.txt', 'tmp/a(b.txt'),
    ('file:///tmp/a%3Bb.txt', 'tmp/a;b.txt'),
    ('file:///tmp/a%00b.txt', 'tmp/a%00b.txt'),
    ('file:///tmp/a%7fb.txt', 'tmp/a%7Fb.txt'),
    ('file:///tmp/a%0ab.txt', 'tmp/a%0Ab.txt'),
    ('file:///tmp/a%ZZb.txt', 'tmp/a%ZZb.txt'),
    ('file:///tmp/a%b.txt', 'tmp/a%b.txt'),
    ('file:///tmp/a b.txt', 'tmp/a%20b.txt'),
    ('file:///tmp/a[1].txt', 'tmp/a[1].txt'),
    ('file:///tmp/a{1}.txt', 'tmp/a{1}.txt'),
    ('file:///tmp/报告 %E6%8A%A5.txt', 'tmp/%E6%8A%A5%E5%91%8A%20%C3%A6%C2%8A%C2%A5.txt'),
    ('file:///tmp/a%23b.txt#锚点', 'tmp/a#b.txt#%E9%94%9A%E7%82%B9'),
    ('file:///tmp/a.txt#章节 空格', 'tmp/a.txt#%E7%AB%A0%E8%8A%82%20%E7%A9%BA%E6%A0%BC'),
    ('file:///tmp/含🙂.txt', 'tmp/%E5%90%AB%F0%9F%99%82.txt'),
    ('file:///tmp/a.txt?x= a#book mark', 'tmp/a.txt?x=%20a#book%20mark'),
    ('file:///tmp:report', 'tmp:report'),
    ('file:///?x=1', '?x=1'),
    ('file:///tmp/a<b>.txt', 'tmp/a<b>.txt'),
    ('file:///tmp/a^b.txt', 'tmp/a^b.txt'),
    ('file:///tmp/a`b.txt', 'tmp/a`b.txt'),
    ('file:///tmp/a|b.txt', 'tmp/a|b.txt'),
    ('file:///tmp/a\\b.txt', 'tmp/a/b.txt'),
    ('file:///tmp/%C3%A9.txt', 'tmp/%C3%83%C2%A9.txt'),
    ('file:///tmp/%FF.txt', 'tmp/%C3%BF.txt'),
    ('file:///tmp/a\nb.txt', 'tmp/a%0Ab.txt'),
])
def test_local_file_pdf_action(target, action):
    pdf = _PdfCanvas(unit="pt", format=(300, 400)); pdf.add_page()
    pdf.link(10, 20, 30, 40, target, alt_text="LINK")
    item = PdfReader(BytesIO(bytes(pdf.output()))).pages[0]["/Annots"][0].get_object()
    assert item["/A"]["/URI"] == action
    assert list(map(float, item["/Rect"])) == [10, 340, 40, 380]


def test_empty_local_file_path_has_no_pdf_action():
    pdf = _PdfCanvas(unit="pt", format=(300, 400)); pdf.add_page()
    pdf.link(10, 20, 30, 40, "file:///", alt_text="LINK")
    assert not PdfReader(BytesIO(bytes(pdf.output()))).pages[0].get("/Annots")
