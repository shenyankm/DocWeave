"""Default paragraph tabs align to the document grid rather than a space."""

import pymupdf
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter


def default_tab_model(step=36, count=1, indent=0, location=2, columns=1):
    paragraph = ldm.Paragraph(
        paragraph_format=ldm.ParagraphFormat(left_indent=indent),
        children=[ldm.Run(text="\t" * count + "RIGHT", font=ldm.Font(size=11))],
    )
    section = ldm.Section(
        page_setup=ldm.PageSetup(
            page_width=300, page_height=180,
            left_margin=20, right_margin=20, top_margin=20, bottom_margin=20,
            text_columns=ldm.TextColumns(count=columns, spacing=15),
        ),
        body=ldm.Body(children=[paragraph]),
    )
    if location != 2:
        section.body.children = [ldm.Paragraph(children=[ldm.Run(text="BODY")])]
        section.headers_footers = [ldm.HeaderFooter(header_footer_type=location, children=[paragraph])]
    return ldm.Document(default_tab_stop=step, sections=[section])


@pytest.mark.parametrize("step", [24, 36, 54])
@pytest.mark.parametrize("count", [1, 2])
@pytest.mark.parametrize("indent", [0, 12])
def test_default_tab_glyph_starts_at_document_grid(step, count, indent):
    model = default_tab_model(step, count, indent)
    snapshot = model.model_dump()
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(model), filetype="pdf") as pdf:
        words = pdf[0].get_text("words")
        assert len(words) == 1 and words[0][4] == "RIGHT"
        assert words[0][0] == pytest.approx(20 + step * count, abs=0.05)
    assert model.model_dump() == snapshot


@pytest.mark.parametrize("location", [0, 1, 2])
@pytest.mark.parametrize("step", [36, 54])
def test_default_tabs_use_page_band_or_body_origin(location, step):
    with pymupdf.open(
        stream=LdmPdfWriter().write_to_bytes(default_tab_model(step, location=location)), filetype="pdf"
    ) as pdf:
        word = next(w for w in pdf[0].get_text("words") if w[4] == "RIGHT")
        assert word[0] == pytest.approx(20 + step, abs=0.05)


def test_default_tabs_follow_second_column_origin():
    model = default_tab_model(columns=2)
    model.sections[0].body.children[:0] = [
        ldm.Paragraph(children=[ldm.Run(text=f"FILL{i}", font=ldm.Font(size=11))])
        for i in range(11)
    ]
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(model), filetype="pdf") as pdf:
        assert len(pdf) == 1
        word = next(w for w in pdf[0].get_text("words") if w[4] == "RIGHT")
        assert word[0] == pytest.approx(157.5 + 36, abs=0.05)


@pytest.mark.parametrize("step", [0, -5])
def test_invalid_default_step_uses_existing_36_point_fallback(step):
    with pymupdf.open(
        stream=LdmPdfWriter().write_to_bytes(default_tab_model(step)), filetype="pdf"
    ) as pdf:
        assert pdf[0].get_text("words")[0][0] == pytest.approx(56, abs=0.05)


def test_default_grid_applies_after_last_explicit_stop():
    model = default_tab_model(count=2)
    model.sections[0].body.children[0].paragraph_format.tab_stops.add(24)
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(model), filetype="pdf") as pdf:
        assert pdf[0].get_text("words")[0][0] == pytest.approx(56, abs=0.05)


@pytest.mark.parametrize("step", [36, 54])
def test_public_docx_default_tab_setting_reaches_pdf(tmp_path, step):
    from aspose.words_foss import Document, SaveFormat
    from aspose.words_foss.docx_writer import LdmDocxWriter

    path = tmp_path / "default-tabs.docx"
    LdmDocxWriter().write(default_tab_model(step), path)
    with pymupdf.open(stream=Document(path).to_bytes(SaveFormat.PDF), filetype="pdf") as pdf:
        assert pdf[0].get_text("words")[0][0] == pytest.approx(20 + step, abs=0.05)


@pytest.mark.parametrize("decoration", ["highlight_color", "strike_through", "link"])
def test_tab_pieces_preserve_inline_rendering(decoration):
    model = default_tab_model(72)
    run = model.sections[0].body.children[0].runs[0]
    run.text = "LEFT\tRIGHT"
    if decoration == "link":
        run.text = "[LEFT](https://example.test)\tRIGHT"
    elif decoration == "highlight_color":
        run.font.highlight_color = "FFFF00"
    else:
        run.font.strike_through = True
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(model), filetype="pdf") as pdf:
        words = {w[4]: w for w in pdf[0].get_text("words")}
        assert set(words) == {"LEFT", "RIGHT"}
        assert words["RIGHT"][0] == pytest.approx(92, abs=0.05)
        if decoration == "link":
            assert [link["uri"] for link in pdf[0].get_links()] == ["https://example.test"]
        else:
            assert len(pdf[0].get_drawings()) == 2


@pytest.mark.parametrize("step", [260, 300])
def test_default_stop_beyond_line_preserves_visible_text(step):
    with pymupdf.open(
        stream=LdmPdfWriter().write_to_bytes(default_tab_model(step)), filetype="pdf"
    ) as pdf:
        assert len(pdf) == 1
        words = pdf[0].get_text("words")
        assert len(words) == 1 and words[0][4] == "RIGHT"
        assert 20 <= words[0][0] < words[0][2] <= 280


@pytest.mark.parametrize("alignment", [0, 1, 2, 3])
def test_explicit_tab_coordinates_share_native_inner_margin_fix(alignment):
    model = default_tab_model()
    paragraph = model.sections[0].body.children[0]
    paragraph.runs[0].text = "\t12.3"
    paragraph.paragraph_format.tab_stops.add(144, alignment=alignment)
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(model), filetype="pdf") as pdf:
        chars = [c for block in pdf[0].get_text("rawdict")["blocks"]
                 for line in block.get("lines", []) for span in line["spans"] for c in span["chars"]]
        if alignment == 0:
            actual = chars[0]["bbox"][0]
        elif alignment == 1:
            actual = (chars[0]["bbox"][0] + chars[-1]["bbox"][2]) / 2
        elif alignment == 2:
            actual = chars[-1]["bbox"][2]
        else:
            actual = next(c["bbox"][0] for c in chars if c["c"] == ".")
        assert actual == pytest.approx(164, abs=0.05)
