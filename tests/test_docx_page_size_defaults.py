"""Owned page-size regressions calibrated against fixed 26.9 loading behavior."""
from io import BytesIO

import pytest
from pypdf import PdfReader

import aspose.words_foss as aw
from aspose.words_foss.pdf_writer import LdmPdfWriter
from .test_docx_dom import package

CASES = [
    ('', 612, 792, 1),
    ('<w:pgSz/>', 612, 792, 1),
    ('<w:pgSz w:orient="landscape"/>', 612, 792, 2),
    ('<w:pgSz w:w="15840"/>', 792, 792, 1),
    ('<w:pgSz w:h="12240"/>', 612, 612, 1),
    ('<w:pgSz w:w="11906" w:h="16838"/>', 595.3, 841.9, 1),
    ('<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>', 792, 612, 2),
]


def dimensions(document):
    ps = document.light_document_model.sections[0].page_setup
    return ps.page_width, ps.page_height, int(ps.orientation)


@pytest.mark.parametrize('size,width,height,orientation', CASES)
@pytest.mark.parametrize('format', [aw.SaveFormat.DOCX, aw.SaveFormat.FLAT_OPC])
def test_missing_dimensions_have_letter_defaults_through_load_save_pdf(
        size, width, height, orientation, format, tmp_path):
    source = tmp_path / 'source.docx'
    package(source, body='<w:p><w:r><w:t>Owned</w:t></w:r></w:p>'
            f'<w:sectPr>{size}</w:sectPr>')
    document = aw.Document(source)
    assert dimensions(document) == (width, height, orientation)
    snapshot = aw.DocxDocument(source).to_light_document()
    assert (snapshot.sections[0].page_setup.page_width,
            snapshot.sections[0].page_setup.page_height) == (width, height)
    output = tmp_path / 'saved'
    document.save(output, format)
    cold = aw.Document(output)
    assert dimensions(cold) == (width, height, orientation)
    page = PdfReader(BytesIO(LdmPdfWriter().write_to_bytes(cold.light_document_model))).pages[0]
    assert float(page.mediabox.width) == pytest.approx(width, abs=.01)
    assert float(page.mediabox.height) == pytest.approx(height, abs=.01)


def test_absent_section_properties_also_use_letter(tmp_path):
    source = tmp_path / 'source.docx'
    package(source, body='<w:p><w:r><w:t>Owned</w:t></w:r></w:p>')
    assert dimensions(aw.Document(source)) == (612, 792, 1)


@pytest.mark.parametrize('second', ['', '<w:pgSz w:w="12000"/>'])
def test_later_section_omissions_use_defaults_not_prior_section(tmp_path, second):
    source = tmp_path / 'source.docx'
    package(source, body='<w:p><w:pPr><w:sectPr><w:pgSz w:w="10000" w:h="14000"/>'
            '</w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Second</w:t></w:r></w:p>'
            f'<w:sectPr>{second}</w:sectPr>')
    sections = aw.Document(source).light_document_model.sections
    assert [(s.page_setup.page_width, s.page_setup.page_height) for s in sections] == [
        (500, 700), (600 if second else 612, 792)]
