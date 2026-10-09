"""Inspect real page content and coordinates for Word paragraph pagination flags."""

import pymupdf
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter


def paragraph(text, **formatting):
    return ldm.Paragraph(paragraph_format=ldm.ParagraphFormat(**formatting),
                         children=[ldm.Run(text=text, font=ldm.Font(size=10))])


def render(children, page_height=200, columns=1):
    model = ldm.Document(sections=[ldm.Section(page_setup=ldm.PageSetup(
        page_width=300, page_height=page_height, top_margin=20, bottom_margin=20,
        left_margin=20, right_margin=20, text_columns=ldm.TextColumns(count=columns, spacing=15)),
        body=ldm.Body(children=children))])
    raw = LdmPdfWriter().write_to_bytes(model)
    return pymupdf.open(stream=raw, filetype='pdf')


@pytest.mark.parametrize('columns', [1, 2])
def test_page_break_before_starts_new_page_without_empty_first_page(columns):
    with render([paragraph('FIRST', page_break_before=True),
                 paragraph('SECOND', page_break_before=True)], columns=columns) as pdf:
        assert len(pdf) == 2
        assert 'FIRST' in pdf[0].get_text() and 'SECOND' not in pdf[0].get_text()
        assert 'SECOND' in pdf[1].get_text()
        assert abs(pdf[0].get_text('words')[0][0] - pdf[1].get_text('words')[0][0]) < 1


@pytest.mark.parametrize('keep', [False, True])
def test_keep_with_next_keeps_heading_with_start_of_long_body(keep):
    children = [paragraph('PREFIX', space_after=120), paragraph('TITLE', keep_with_next=keep),
                paragraph('\n'.join(f'BODY{i:02}' for i in range(20)))]
    with render(children) as pdf:
        title_page = next(i for i, page in enumerate(pdf) if 'TITLE' in page.get_text())
        body_page = next(i for i, page in enumerate(pdf) if 'BODY00' in page.get_text())
        assert title_page == body_page if keep else title_page < body_page
        text = ''.join(page.get_text() for page in pdf)
        assert text.count('TITLE') == 1
        assert all(text.count(f'BODY{i:02}') == 1 for i in range(20))


def test_keep_with_next_chain_moves_together_and_honors_explicit_break():
    with render([paragraph('PREFIX', space_after=100), paragraph('TITLE', keep_with_next=True),
                 paragraph('SUBTITLE', keep_with_next=True), paragraph('BODY')]) as pdf:
        pages = [next(i for i, p in enumerate(pdf) if token in p.get_text())
                 for token in ('TITLE', 'SUBTITLE', 'BODY')]
        assert len(set(pages)) == 1
    with render([paragraph('TITLE', keep_with_next=True), paragraph('BODY', page_break_before=True)]) as pdf:
        assert 'TITLE' in pdf[0].get_text() and 'BODY' in pdf[1].get_text()


@pytest.mark.parametrize('keep', [False, True])
def test_column_balancing_respects_keep_chain(keep):
    with render([paragraph('TITLE', keep_with_next=keep),
                 paragraph('SUBTITLE', keep_with_next=keep), paragraph('BODY')], columns=2) as pdf:
        assert len(pdf) == 1
        words = {word[4]: word for word in pdf[0].get_text('words')}
        xs = [words[token][0] for token in ('TITLE', 'SUBTITLE', 'BODY')]
        assert max(xs) - min(xs) < 1 if keep else max(xs) - min(xs) > 100


def test_column_balancing_keeps_title_with_first_table_row():
    table = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph('CELL')])])])
    with render([paragraph('TITLE', keep_with_next=True), table], columns=2) as pdf:
        assert len(pdf) == 1
        words = {word[4]: word for word in pdf[0].get_text('words')}
        assert abs(words['TITLE'][0] - words['CELL'][0]) < 20


def test_column_balancer_moves_a_keep_group_as_one_unit():
    with render([paragraph('PREFIX'), paragraph('TITLE', keep_with_next=True),
                 paragraph('BODY')], columns=2) as pdf:
        words = {word[4]: word for word in pdf[0].get_text('words')}
        assert words['TITLE'][0] - words['PREFIX'][0] > 100
        assert abs(words['TITLE'][0] - words['BODY'][0]) < 1


@pytest.mark.parametrize('following', ['list', 'table'])
def test_keep_with_next_measures_list_and_first_table_row(following):
    if following == 'list':
        node = paragraph('ITEM0\nITEM1')
        node.list_format = ldm.ListFormat(is_list_item=True)
    else:
        node = ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(paragraphs=[paragraph('ITEM0\nITEM1')])])])
    with render([paragraph('PREFIX', space_after=120), paragraph('TITLE', keep_with_next=True), node]) as pdf:
        title_page = next(i for i, p in enumerate(pdf) if 'TITLE' in p.get_text())
        assert 'ITEM0' in pdf[title_page].get_text()


@pytest.mark.parametrize('widow', [False, True])
def test_widow_control_does_not_leave_a_single_last_line(widow):
    # 170 pt of usable height fits twelve 14-pt lines, leaving one line on page two.
    with render([paragraph('\n'.join(f'LINE{i:02}' for i in range(13)), widow_control=widow)], page_height=210) as pdf:
        assert len(pdf) == 2
        lines = [page.get_text().splitlines() for page in pdf]
        assert sum(len(items) for items in lines) == 13
        assert len(lines[-1]) >= 2 if widow else len(lines[-1]) == 1
        assert [item for items in lines for item in items] == [f'LINE{i:02}' for i in range(13)]


def test_widow_control_moves_first_two_lines_and_does_not_loop_on_tiny_pages():
    with render([paragraph('PREFIX', space_after=130), paragraph('LINE0\nLINE1\nLINE2')]) as pdf:
        assert 'LINE0' not in pdf[0].get_text()
        assert all(word in pdf[1].get_text() for word in ('LINE0', 'LINE1'))
    with render([paragraph('ONE\nTWO\nTHREE')], page_height=58) as pdf:
        assert len(pdf) == 3
        assert ''.join(page.get_text() for page in pdf).split() == ['ONE', 'TWO', 'THREE']
