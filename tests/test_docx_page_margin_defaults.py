"""Owned default/explicit-zero margin regressions calibrated against 26.9."""

import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.pdf_writer.constants import PT_TO_MM
from .test_docx_dom import package

FIELDS = ('top_margin', 'bottom_margin', 'left_margin', 'right_margin',
          'header_distance', 'footer_distance', 'gutter')
DEFAULTS = (70.85, 70.85, 70.85, 70.85, 35.4, 35.4, 0)
CASES = [
    ('', DEFAULTS),
    ('<w:pgMar/>', DEFAULTS),
    ('<w:pgMar w:top="1440"/>', (72, *DEFAULTS[1:])),
    ('<w:pgMar w:top=""/>', (0, *DEFAULTS[1:])),
    ('<w:pgMar w:top="oops"/>', (0, *DEFAULTS[1:])),
    ('<w:pgMar w:top="-1440"/>', (-72, *DEFAULTS[1:])),
    ('<w:pgMar w:top="0" w:bottom="0" w:left="0" w:right="0" '
     'w:header="0" w:footer="0" w:gutter="0"/>', (0,) * 7),
]


def values(model):
    return tuple(getattr(model.sections[0].page_setup, key) for key in FIELDS)


@pytest.mark.parametrize('margin,expected', CASES)
@pytest.mark.parametrize('format', [aw.SaveFormat.DOCX, aw.SaveFormat.FLAT_OPC])
def test_default_partial_zero_and_normalized_margins_roundtrip(margin, expected, format, tmp_path):
    source = tmp_path / 'source.docx'
    package(source, body='<w:p><w:r><w:t>Owned margin</w:t></w:r></w:p>'
            f'<w:sectPr>{margin}</w:sectPr>')
    document = aw.Document(source)
    model = document.light_document_model
    assert values(model) == expected
    assert values(aw.DocxDocument(source).to_light_document()) == expected
    assert values(ldm.Document.model_validate_json(model.model_dump_json())) == expected
    target = tmp_path / 'saved'
    document.save(target, format)
    assert values(aw.Document(target).light_document_model) == expected
    if expected[0] >= 0:
        writer = LdmPdfWriter()
        writer.write_to_bytes(model)
        assert writer._page_margin_left == pytest.approx(expected[2] * PT_TO_MM)
        assert writer._page_margin_bottom == pytest.approx(expected[1] * PT_TO_MM)


def test_default_new_model_and_later_section_do_not_inherit_zero(tmp_path):
    assert tuple(getattr(ldm.PageSetup(), key) for key in FIELDS) == DEFAULTS
    source = tmp_path / 'source.docx'
    package(source, body='<w:p><w:pPr><w:sectPr><w:pgMar w:top="0"/>'
            '</w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Second</w:t></w:r></w:p><w:sectPr/>')
    sections = aw.Document(source).light_document_model.sections
    assert sections[0].page_setup.top_margin == 0
    assert tuple(getattr(sections[1].page_setup, key) for key in FIELDS) == DEFAULTS
