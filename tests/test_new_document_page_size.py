"""New and round-tripped Word documents use the fixed baseline's Letter dimensions."""

from io import BytesIO

import pytest
from pypdf import PdfReader

from ._word_fixtures import word_bytes

import dockiln as dk
from dockiln import light_document_model as ldm
from dockiln.writers.pdf import LdmPdfWriter


@pytest.mark.parametrize('kind', ['new', 'docx'])
def test_default_page_size_survives_json_outputs_and_pdf(kind, tmp_path):
    if kind == 'new':
        doc = dk.Document()
    else:
        source = tmp_path / ('input.' + kind)
        source.write_bytes(word_bytes("OWNED_WORD"))
        doc = dk.Document(source)
    model = doc.light_document_model
    page = model.sections[0].page_setup
    assert (page.page_width, page.page_height, page.paper_size) == (612, 792, 1)
    restored = ldm.Document.model_validate_json(model.model_dump_json())
    assert restored.sections[0].page_setup == page
    pdf = PdfReader(BytesIO(doc.to_bytes(dk.SaveFormat.PDF))).pages[0]
    assert tuple(map(float, (pdf.mediabox.width, pdf.mediabox.height))) == (612, 792)
    for format in (dk.SaveFormat.DOCX, dk.SaveFormat.FLAT_OPC):
        cold = dk.Document(BytesIO(doc.to_bytes(format)))
        actual = cold.light_document_model.sections[0].page_setup
        assert (actual.page_width, actual.page_height, actual.paper_size) == (612, 792, 1)


def test_new_section_and_empty_json_page_setup_render_letter():
    for setup in (ldm.PageSetup(), ldm.PageSetup.model_validate({})):
        model = ldm.Document(sections=[ldm.Section(page_setup=setup)])
        pdf = PdfReader(BytesIO(LdmPdfWriter().write_to_bytes(model))).pages[0]
        assert float(pdf.mediabox.width) == 612
        assert float(pdf.mediabox.height) == 792


def test_explicit_a4_page_dimensions_are_preserved():
    setup = ldm.PageSetup(page_width=595.3, page_height=841.9, paper_size=9)
    assert ldm.PageSetup.model_validate_json(setup.model_dump_json()) == setup
    model = ldm.Document(sections=[ldm.Section(page_setup=setup)])
    pdf = PdfReader(BytesIO(LdmPdfWriter().write_to_bytes(model))).pages[0]
    assert float(pdf.mediabox.width) == 595.3
    assert float(pdf.mediabox.height) == 841.9
