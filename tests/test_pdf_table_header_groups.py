"""Leading table header rows stay together in every page or column region."""

import pymupdf
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter


def paragraph(text, **formatting):
    return ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(size=10))],
                         paragraph_format=ldm.ParagraphFormat(**formatting))


def header_table(header_count, lines=1):
    return ldm.Table(rows=[
        *(ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph(f'HEADER{i}')])],
                  row_format=ldm.RowFormat(heading_format=True)) for i in range(header_count)),
        ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph('\n'.join(f'BODY{i:02}' for i in range(lines)))])]),
    ])


def render(children, columns=1):
    model = ldm.Document(sections=[ldm.Section(page_setup=ldm.PageSetup(
        page_width=300, page_height=200, left_margin=20, right_margin=20,
        top_margin=20, bottom_margin=20,
        text_columns=ldm.TextColumns(count=columns, spacing=15)),
        body=ldm.Body(children=children))])
    return pymupdf.open(stream=LdmPdfWriter().write_to_bytes(model), filetype='pdf')


@pytest.mark.parametrize('columns', [1, 2])
@pytest.mark.parametrize('header_count', [2, 3])
def test_initial_header_group_is_never_split(columns, header_count):
    with render([paragraph('PREFIX', space_after=120), header_table(header_count)], columns) as pdf:
        positions = [(page_no, word) for page_no, page in enumerate(pdf) for word in page.get_text('words')]
        first = next((page_no, word) for page_no, word in positions if word[4] == 'HEADER0')
        # Every appearance of the first header must have the complete ordered group beside it.
        for page_no, word in positions:
            if word[4] == 'HEADER0':
                for i in range(header_count):
                    matching = [other for p, other in positions if p == page_no and other[4] == f'HEADER{i}' and abs(other[0] - word[0]) < 1]
                    assert matching and matching[0][1] >= word[1]
        assert all(p == first[0] and abs(word[0] - first[1][0]) < 1
                   for p, word in positions if word[4].startswith('HEADER'))
        assert ''.join(page.get_text() for page in pdf).count('BODY00') == 1


def test_multiple_headers_repeat_above_each_long_row_fragment():
    with render([header_table(2, lines=40)]) as pdf:
        assert len(pdf) > 2
        text = ''.join(page.get_text() for page in pdf)
        assert all(text.count(f'BODY{i:02}') == 1 for i in range(40))
        for page in pdf:
            words = page.get_text('words')
            headers = [word for word in words if word[4].startswith('HEADER')]
            assert [word[4] for word in headers] == ['HEADER0', 'HEADER1']
            assert all(word[1] > headers[-1][1] for word in words if word[4].startswith('BODY'))
            assert all(19 <= word[1] < word[3] <= 181 for word in words)


def test_header_only_table_fits_without_an_empty_first_page():
    table = header_table(3)
    table.rows.pop()
    with render([table]) as pdf:
        assert len(pdf) == 1
        assert [word[4] for word in pdf[0].get_text('words')] == ['HEADER0', 'HEADER1', 'HEADER2']
