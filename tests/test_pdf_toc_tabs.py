"""TOC right tabs use the source position rather than the column edge."""

import pymupdf
import pytest

from aspose.words_foss import Document
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.saving import PdfSaveOptions


def toc_tab_model(position=180, columns=1, location="body"):
    pf = ldm.ParagraphFormat(style_name="TOC 1", left_indent=12, right_indent=6)
    pf.tab_stops.add(position, alignment=2)
    paragraph = ldm.Paragraph(
        paragraph_format=pf,
        children=[
            ldm.Run(text="TITLE", font=ldm.Font(size=12)),
            ldm.Run(text="\t123", font=ldm.Font(size=16, bold=True)),
        ],
    )
    return ldm.Document(
        sections=[
            ldm.Section(
                page_setup=ldm.PageSetup(
                    page_width=300,
                    page_height=180,
                    left_margin=20,
                    right_margin=20,
                    top_margin=20,
                    bottom_margin=20,
                    text_columns=ldm.TextColumns(count=columns, spacing=15),
                ),
                body=ldm.Body(
                    children=[paragraph]
                    if location == "body"
                    else [ldm.Paragraph(children=[ldm.Run(text="BODY")])]
                ),
                headers_footers=[]
                if location == "body"
                else [
                    ldm.HeaderFooter(
                        header_footer_type=0 if location == "header" else 1,
                        children=[paragraph],
                    )
                ],
            )
        ]
    )


@pytest.mark.parametrize("columns", [1, 2])
@pytest.mark.parametrize("location", ["body", "header", "footer"])
@pytest.mark.parametrize("shaping", [False, True])
def test_toc_page_number_ends_at_explicit_right_tab(columns, location, shaping):
    position = 180 if columns == 1 else 100
    model = toc_tab_model(position, columns, location)
    original = model.model_dump()
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(
        stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf"
    ) as pdf:
        assert len(pdf) == 1
        words = {w[4]: w for w in pdf[0].get_text("words")}
        assert words["123"][2] == pytest.approx(20 + position, abs=0.05)
        assert words["123"][0] > words["TITLE"][2]
    assert model.model_dump() == original


@pytest.mark.parametrize("shaping", [False, True])
def test_public_docx_toc_right_tab_roundtrip(tmp_path, shaping):
    model = toc_tab_model()
    source = tmp_path / "tab.docx"
    LdmDocxWriter().write(model, source)
    document = Document(source)
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=document.to_bytes(options), filetype="pdf") as pdf:
        words = {w[4]: w for w in pdf[0].get_text("words")}
        assert words["123"][2] == pytest.approx(200, abs=0.05)
        assert words["123"][0] > words["TITLE"][2]


def test_out_of_area_toc_tab_does_not_replace_existing_target(tmp_path):
    model = toc_tab_model(position=280)
    target = tmp_path / "keep.pdf"
    target.write_bytes(b"KEEP")
    with pytest.raises(ValueError, match="TOC tab|usable text"):
        LdmPdfWriter().write(model, target)
    assert target.read_bytes() == b"KEEP"


@pytest.mark.parametrize("shaping", [False, True])
def test_custom_toc_tab_survives_title_page_and_column_breaks(shaping):
    model = toc_tab_model(position=100, columns=2)
    tokens = [f"T{index:03}" for index in range(48)]
    model.sections[0].body.children[0]._children[0].text = " ".join(tokens)
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(
        stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf"
    ) as pdf:
        assert len(pdf) > 1
        records = [
            (page_index, word)
            for page_index, page in enumerate(pdf)
            for word in page.get_text("words")
        ]
        assert [word[4] for _, word in records] == tokens + ["123"]
        number_page, number = records[-1]
        last_page, last = records[-2]
        column = 1 if number[0] > 157.5 else 0
        left = 20 + 137.5 * column
        assert number_page == last_page
        assert (last[0] > 157.5) == (column == 1)
        assert number[0] > last[2]
        assert number[2] == pytest.approx(left + 100, abs=0.05)
        for _, word in records[:-1]:
            title_left = 157.5 if word[0] > 157.5 else 20
            assert word[0] >= title_left + 12 - 0.05
            assert word[2] < title_left + 100 - 0.05


@pytest.mark.parametrize("shaping", [False, True])
@pytest.mark.parametrize("mode", ["cleared", "unsorted"])
def test_toc_field_selection_ignores_clear_and_uses_position_order(shaping, mode):
    model = toc_tab_model()
    paragraph = model.sections[0].body.children[0]
    if mode == "cleared":
        paragraph.paragraph_format.tab_stops.tab_stops[0].is_clear = True
        expected = 300 - 20 - 6
    else:
        paragraph.paragraph_format.tab_stops.tab_stops = [
            ldm.TabStop(position=220, alignment=2),
            ldm.TabStop(position=140, alignment=2),
            ldm.TabStop(position=120, alignment=2, is_clear=True),
            ldm.TabStop(position=80, alignment=0),
        ]
        expected = 100
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(
        stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf"
    ) as pdf:
        words = {w[4]: w for w in pdf[0].get_text("words")}
        actual = words["123"][2 if mode == "cleared" else 0]
        assert actual == pytest.approx(expected, abs=0.05)


@pytest.mark.parametrize("columns", [1, 2])
@pytest.mark.parametrize("shaping", [False, True])
def test_explicit_toc_tab_can_end_at_column_boundary(columns, shaping):
    position = (260 - 15 * (columns - 1)) / columns
    model = toc_tab_model(position=position, columns=columns)
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(
        stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf"
    ) as pdf:
        words = {w[4]: w for w in pdf[0].get_text("words")}
        assert set(words) == {"TITLE", "123"}
        assert words["123"][2] == pytest.approx(20 + position, abs=0.05)
