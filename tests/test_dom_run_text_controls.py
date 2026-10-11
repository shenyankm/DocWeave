"""Raw text views retain Word controls without changing plain range offsets."""
from io import BytesIO

import pytest

from aspose.words_foss._flat_opc import encode
from aspose.words_foss.dom import DocxDocument
from tests.test_dom_get_text import controlled_document


@pytest.mark.parametrize('flat', [False, True])
@pytest.mark.parametrize('control,expected', [
    ('<w:tab/>', '\t'), ('<w:ptab w:alignment="left" w:relativeTo="margin" w:leader="none"/>', ''),
    ('<w:br/>', '\x0b'), ('<w:br w:type="page"/>', '\x0c'),
    ('<w:br w:type="column"/>', '\x0e'), ('<w:cr/>', '\r'),
    ('<w:noBreakHyphen/>', '\x1e'), ('<w:softHyphen/>', '\x1f'),
])
def test_raw_text_control_views_are_read_only_and_survive_saves(flat, control, expected):
    raw = controlled_document('<w:r><w:t>OWNED A</w:t>' + control + '<w:t>OWNED B</w:t></w:r>')
    document = DocxDocument(BytesIO(encode(raw) if flat else raw))
    for _ in range(2):
        before = document.to_bytes()
        paragraph = document.body.paragraphs[0]
        assert paragraph.runs[0].text == paragraph.text == 'OWNED A' + expected + 'OWNED B'
        assert document.to_bytes() == before
        with pytest.raises(NotImplementedError):
            paragraph.range()
        document = DocxDocument(BytesIO(document.to_flat_opc() if flat else document.to_bytes()))


def test_raw_text_view_keeps_unmarked_whitespace_and_plain_range_offsets():
    document = DocxDocument(BytesIO(controlled_document('<w:r><w:t>  OWNED A  </w:t></w:r>')))
    paragraph = document.body.paragraphs[0]
    assert paragraph.runs[0].text == paragraph.text == '  OWNED A  '
    selected = paragraph.range(2, 7)
    assert selected.text == 'OWNED'
    selected.replace('EDITED')
    assert paragraph.text == '  EDITED A  '
