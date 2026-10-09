"""Long directory titles must wrap without painting over their page number."""

import re
from io import BytesIO

import pymupdf
import pytest
from pypdf import PdfReader

from aspose.words_foss import Document
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.pdf_writer.constants import PT_TO_MM
from aspose.words_foss.saving import PdfSaveOptions


def long_toc_model(columns=1, chinese=False, count=48, page_height=180):
    tokens = [f"{'条目' if chinese else 'ITEM'}{index:03}" for index in range(count)]
    mid = count // 2
    paragraph = ldm.Paragraph(
        paragraph_format=ldm.ParagraphFormat(
            style_name="TOC 1", left_indent=12, right_indent=6
        ),
        children=[
            ldm.Run(text=" ".join(tokens[:mid]) + " ", font=ldm.Font(size=12)),
            ldm.Run(text=" ".join(tokens[mid:]), font=ldm.Font(size=14, bold=True)),
            ldm.Run(text="\t999", font=ldm.Font(size=16)),
        ],
    )
    model = ldm.Document(
        sections=[
            ldm.Section(
                page_setup=ldm.PageSetup(
                    page_width=300,
                    page_height=page_height,
                    left_margin=20,
                    right_margin=20,
                    top_margin=20,
                    bottom_margin=20,
                    text_columns=ldm.TextColumns(count=columns, spacing=15),
                ),
                body=ldm.Body(
                    children=[
                        paragraph,
                        ldm.Paragraph(children=[ldm.Run(text="AFTER")]),
                    ]
                ),
            )
        ]
    )
    return model, tokens


def assert_toc_content_and_bounds(raw, tokens, columns):
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        text = re.sub(r"\s+", "", "".join(page.get_text() for page in pdf))
        assert text == "".join(tokens) + "999AFTER"
        width = (260 - 15 * (columns - 1)) / columns
        number = None
        last = None
        for page_index, page in enumerate(pdf):
            for word in page.get_text("words"):
                if word[4] == "AFTER":
                    continue
                column = 1 if columns == 2 and word[0] > 157.5 else 0
                left = 20 + column * (width + 15)
                assert word[0] >= left + 12 - 0.05
                assert word[2] <= left + width - 6 + 0.05
                assert 20 - 0.05 <= word[1] < word[3] <= 160 + 0.05
                if word[4] == "999":
                    number = (page_index, word, column, left)
                if tokens[-1] in word[4]:
                    last = (page_index, word, column)
        assert number is not None and last is not None
        assert number[0] == last[0] and number[2] == last[2]
        assert number[1][0] - last[1][2] >= 72 / 25.4 - 0.05
        assert number[1][2] == pytest.approx(
            number[3] + width - 6, abs=0.05
        )
        # A 14/16 pt pair has different boxes; separate rows have a full line-height gap.
        assert abs(number[1][1] - last[1][1]) < 3
        return len(pdf)


@pytest.mark.parametrize("columns", [1, 2])
@pytest.mark.parametrize("chinese", [False, True])
@pytest.mark.parametrize("shaping", [False, True])
def test_long_toc_preserves_all_title_text_and_last_line_page_number(
    columns, chinese, shaping
):
    model, tokens = long_toc_model(columns, chinese)
    snapshot = model.model_dump()
    options = PdfSaveOptions()
    options.text_shaping = shaping
    options.export_document_structure = True
    raw = LdmPdfWriter(options).write_to_bytes(model)
    assert assert_toc_content_and_bounds(raw, tokens, columns) > 1
    reader = PdfReader(BytesIO(raw))
    root = reader.trailer["/Root"]["/StructTreeRoot"]
    paragraph = root["/K"][0].get_object()["/K"][0]
    assert paragraph.get_object()["/S"] == "/P"
    fragments = paragraph.get_object()["/K"]
    assert len(fragments) > 1
    nums = root["/ParentTree"]["/Nums"]
    parents = {int(nums[i]): nums[i + 1].get_object() for i in range(0, len(nums), 2)}
    for fragment in fragments:
        page = fragment["/Pg"]
        assert parents[int(page["/StructParents"])][int(fragment["/MCID"])] == paragraph
    assert model.model_dump() == snapshot


@pytest.mark.parametrize("shaping", [False, True])
def test_public_long_docx_toc_conversion(tmp_path, shaping):
    model, tokens = long_toc_model(chinese=True)
    source = tmp_path / "long-toc.docx"
    LdmDocxWriter().write(model, source)
    options = PdfSaveOptions()
    options.text_shaping = shaping
    raw = Document(source).to_bytes(options)
    assert_toc_content_and_bounds(raw, tokens, 1)


@pytest.mark.parametrize("shaping", [False, True])
def test_toc_height_measurement_matches_actual_wrapped_render(
    tmp_path, monkeypatch, shaping
):
    model, _ = long_toc_model(count=24, page_height=1200)
    model.sections[0].body.children = model.sections[0].body.children[:1]
    options = PdfSaveOptions()
    options.text_shaping = shaping
    writer = LdmPdfWriter(options)
    original = writer._paragraph_renderer.render_paragraph
    measured = []

    def record(pdf, paragraph):
        expected = writer._estimate_paragraph_height(paragraph, pdf.epw)
        start = pdf.y
        original(pdf, paragraph)
        measured.append((expected, pdf.y - start))

    monkeypatch.setattr(writer._paragraph_renderer, "render_paragraph", record)
    writer.write(model, tmp_path / "measured.pdf")
    assert len(measured) == 1
    assert measured[0][0] > 19.2 * PT_TO_MM
    assert measured[0][0] == pytest.approx(measured[0][1], abs=0.05)


@pytest.mark.parametrize("shaping", [False, True])
@pytest.mark.parametrize("widow", [False, True])
def test_toc_widow_rule_keeps_first_two_title_rows_together(shaping, widow):
    model, tokens = long_toc_model(count=6)
    paragraph = model.sections[0].body.children[0]
    paragraph.paragraph_format.widow_control = widow
    model.sections[0].body.children.insert(
        0,
        ldm.Paragraph(
            paragraph_format=ldm.ParagraphFormat(space_after=100),
            children=[ldm.Run(text="PREFIX", font=ldm.Font(size=12))],
        ),
    )
    options = PdfSaveOptions()
    options.text_shaping = shaping
    raw = LdmPdfWriter(options).write_to_bytes(model)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        assert len(pdf) == 2
        assert "PREFIX" in pdf[0].get_text()
        assert (tokens[0] in pdf[0].get_text()) == (not widow)
        assert tokens[-1] in pdf[1].get_text() and "999" in pdf[1].get_text()


def test_toc_without_title_room_preserves_existing_output(tmp_path):
    model, _ = long_toc_model(count=6)
    paragraph = model.sections[0].body.children[0]
    paragraph._children[-1].text = "\t" + "9" * 100
    target = tmp_path / "keep.pdf"
    target.write_bytes(b"KEEP")
    with pytest.raises(ValueError, match="TOC page number|usable text"):
        LdmPdfWriter().write(model, target)
    assert target.read_bytes() == b"KEEP"


@pytest.mark.parametrize("shaping", [False, True])
@pytest.mark.parametrize("side", ["title", "number"])
def test_toc_missing_one_side_keeps_remaining_content(side, shaping):
    model, tokens = long_toc_model(count=12)
    paragraph = model.sections[0].body.children[0]
    if side == "title":
        paragraph._children[0].text = paragraph._children[1].text = ""
        expected = "999AFTER"
    else:
        paragraph._children[-1].text = "\t"
        expected = "".join(tokens) + "AFTER"
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(
        stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf"
    ) as pdf:
        assert re.sub(r"\s+", "", "".join(page.get_text() for page in pdf)) == expected


@pytest.mark.parametrize("shaping", [False, True])
def test_wrapped_toc_retains_title_and_page_number_links(shaping):
    model, tokens = long_toc_model(columns=2)
    paragraph = model.sections[0].body.children[0]
    for run in paragraph._children[:2]:
        run.text = f"[{run.text}](https://example.org/title)"
        run.is_hyperlink = True
    paragraph._children[-1].text = "\t[999](https://example.org/page)"
    paragraph._children[-1].is_hyperlink = True
    options = PdfSaveOptions()
    options.text_shaping = shaping
    raw = LdmPdfWriter(options).write_to_bytes(model)
    assert_toc_content_and_bounds(raw, tokens, 2)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        uris = [link["uri"] for page in pdf for link in page.get_links()]
        assert uris.count("https://example.org/page") == 1
        assert uris.count("https://example.org/title") > 1
        assert set(uris) == {"https://example.org/page", "https://example.org/title"}
