"""Header/footer distance changes move owned text by the declared point delta."""

import pymupdf
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter


def band_origin(distance, band):
    paragraph = ldm.Paragraph(
        children=[ldm.Run(text='OWNED_BAND', font=ldm.Font(name='Document Sans SC', size=11))],
        paragraph_format=ldm.ParagraphFormat(line_spacing=14, line_spacing_rule=1),
    )
    setup = ldm.PageSetup(page_width=612, page_height=792, top_margin=200,
                          bottom_margin=70.85, header_distance=distance,
                          footer_distance=distance)
    doc = ldm.Document(sections=[ldm.Section(page_setup=setup,
        body=ldm.Body(children=[ldm.Paragraph(children=[ldm.Run(text='BODY')])]),
        headers_footers=[ldm.HeaderFooter(header_footer_type=band, children=[paragraph])])])
    snapshot = doc.model_dump()
    raw = LdmPdfWriter().write_to_bytes(doc)
    assert doc.model_dump() == snapshot
    with pymupdf.open(stream=raw, filetype='pdf') as pdf:
        assert len(pdf) == 1
        spans = [span for block in pdf[0].get_text('dict')['blocks'] if 'lines' in block
                 for line in block['lines'] for span in line['spans']
                 if span['text'] == 'OWNED_BAND']
        assert len(spans) == 1
        return spans[0]['origin']


@pytest.mark.parametrize('band', [0, 1])
@pytest.mark.parametrize('distance', [0, 2, 72, 144])
def test_band_distance_respects_zero_and_does_not_clamp_to_body_margin(band, distance):
    reference = band_origin(35.4, band)
    actual = band_origin(distance, band)
    assert actual[0] == pytest.approx(reference[0], abs=.01)
    expected_delta = (distance - 35.4) * (1 if band == 0 else -1)
    assert actual[1] - reference[1] == pytest.approx(expected_delta, abs=.01)
