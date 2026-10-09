"""TOC center and decimal fields use their configured anchors across font runs."""

import pymupdf
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.saving import PdfSaveOptions


def aligned_toc_model(alignment, columns=1, number="12", separator=".", suffix="30", long=False):
    pf = ldm.ParagraphFormat(style_name="TOC 1", left_indent=12, right_indent=6)
    pf.tab_stops.add(80 if columns == 2 else 180, alignment=alignment, leader=1)
    return ldm.Document(sections=[ldm.Section(
        page_setup=ldm.PageSetup(page_width=300, page_height=180,
                                left_margin=20, right_margin=20, top_margin=20, bottom_margin=20,
                                text_columns=ldm.TextColumns(count=columns, spacing=15)),
        body=ldm.Body(children=[ldm.Paragraph(paragraph_format=pf, children=[
            ldm.Run(text=" ".join(f"T{i:03}" for i in range(48)) if long else "TITLE",
                    font=ldm.Font(size=12)),
            ldm.Run(text="\t" + number, font=ldm.Font(size=16, bold=True)),
            ldm.Run(text=separator, font=ldm.Font(size=12)),
            ldm.Run(text=suffix, font=ldm.Font(size=18, italic=True)),
        ])]),
    )])


def page_chars(page):
    return [c for block in page.get_text("rawdict")["blocks"]
            for line in block.get("lines", []) for span in line["spans"] for c in span["chars"]]


def field_chars(page):
    return [c for block in page.get_text("rawdict")["blocks"]
            for line in block.get("lines", []) for span in line["spans"] for c in span["chars"]
            if c["c"] in ".," or (span["size"] > 12.1 and c["c"] in "1230")]


def field_anchor(chars, alignment):
    if alignment == 0:
        return min(c["bbox"][0] for c in chars)
    if alignment == 1:
        return (min(c["bbox"][0] for c in chars) + max(c["bbox"][2] for c in chars)) / 2
    return next(c["bbox"][0] for c in chars if c["c"] in ".,")


@pytest.mark.parametrize("alignment", [0, 1, 3])
@pytest.mark.parametrize("columns", [1, 2])
@pytest.mark.parametrize("shaping", [False, True])
def test_toc_field_anchor_matches_source(alignment, columns, shaping):
    model = aligned_toc_model(alignment, columns)
    snapshot = model.model_dump()
    options = PdfSaveOptions()
    options.text_shaping = shaping
    options.export_document_structure = True
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf") as pdf:
        assert len(pdf) == 1
        chars = field_chars(pdf[0])
        assert "".join(c["c"] for c in chars) == "12.30"
        actual = field_anchor(chars, alignment)
        assert actual == pytest.approx(100 if columns == 2 else 200, abs=0.05)
        title = next(w for w in pdf[0].get_text("words") if w[4] == "TITLE")
        drawing, = pdf[0].get_drawings()
        assert title[2] < drawing["rect"].x0 < drawing["rect"].x1 < min(c["bbox"][0] for c in chars)
    assert model.model_dump() == snapshot


@pytest.mark.parametrize("shaping", [False, True])
def test_decimal_comma_anchor_uses_mixed_run_widths(shaping):
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(
        aligned_toc_model(3, separator=",")
    ), filetype="pdf") as pdf:
        comma = next(c for c in page_chars(pdf[0]) if c["c"] == ",")
        assert comma["bbox"][0] == pytest.approx(200, abs=0.05)


@pytest.mark.parametrize("shaping", [False, True])
def test_decimal_without_separator_aligns_at_field_end(shaping):
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(
        aligned_toc_model(3, separator="", suffix="")
    ), filetype="pdf") as pdf:
        number = next(w for w in pdf[0].get_text("words") if w[4] == "12")
        assert number[2] == pytest.approx(200, abs=0.05)


@pytest.mark.parametrize("shaping", [False, True])
def test_center_field_can_be_wider_than_its_anchor_distance(shaping):
    model = aligned_toc_model(1, number="12345678901234")
    model.sections[0].body.children[0].paragraph_format.tab_stops.tab_stops[0].position = 120
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf") as pdf:
        chars = page_chars(pdf[0])
        number = [c for c in chars if c["c"] in "0123456789."]
        assert "".join(c["c"] for c in number) == "12345678901234.30"
        assert (min(c["bbox"][0] for c in number) + max(c["bbox"][2] for c in number)) / 2 == pytest.approx(140, abs=0.05)


@pytest.mark.parametrize("prefix", ["ffi12", "e\u030112", "中文12"])
@pytest.mark.parametrize("shaping", [False, True])
def test_decimal_prefix_inside_run_uses_rendered_width(prefix, shaping):
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(
        aligned_toc_model(3, number=prefix + ".3", separator="", suffix="0")
    ), filetype="pdf") as pdf:
        point = next(c for c in page_chars(pdf[0]) if c["c"] == ".")
        assert point["bbox"][0] == pytest.approx(200, abs=0.05)


@pytest.mark.parametrize("alignment", [0, 1, 3])
@pytest.mark.parametrize("columns", [1, 2])
@pytest.mark.parametrize("shaping", [False, True])
def test_long_title_keeps_field_on_last_page_and_column(alignment, columns, shaping):
    model = aligned_toc_model(alignment, columns, long=True)
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf") as pdf:
        words = [w[4] for page in pdf for w in page.get_text("words")]
        assert [w for w in words if w.startswith("T")] == [f"T{i:03}" for i in range(48)]
        assert len(pdf) > 1
        last = next(w for w in pdf[-1].get_text("words") if w[4] == "T047")
        chars = field_chars(pdf[-1])
        actual = field_anchor(chars, alignment)
        origin = 157.5 if columns == 2 and last[0] > 150 else 20
        assert actual == pytest.approx(origin + (80 if columns == 2 else 180), abs=0.05)
        assert sum(len(page.get_drawings()) for page in pdf) == 1


@pytest.mark.parametrize("alignment", [1, 3])
@pytest.mark.parametrize("position", [20, 260])
@pytest.mark.parametrize("shaping", [False, True])
def test_outside_field_preserves_existing_target(tmp_path, alignment, position, shaping):
    from aspose.words_foss import Document
    from aspose.words_foss.docx_writer import LdmDocxWriter

    model = aligned_toc_model(alignment)
    model.sections[0].body.children[0].paragraph_format.tab_stops.tab_stops[0].position = position
    source = tmp_path / "outside.docx"
    LdmDocxWriter().write(model, source)
    target = tmp_path / "existing.pdf"
    target.write_bytes(b"keep-existing")
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pytest.raises(ValueError, match="TOC field is outside"):
        Document(source).save(target, options)
    assert target.read_bytes() == b"keep-existing"


@pytest.mark.parametrize("alignment", [0, 1, 3])
@pytest.mark.parametrize("shaping", [False, True])
def test_public_docx_field_anchor(tmp_path, alignment, shaping):
    from aspose.words_foss import Document
    from aspose.words_foss.docx_writer import LdmDocxWriter

    source = tmp_path / "field.docx"
    LdmDocxWriter().write(aligned_toc_model(alignment), source)
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=Document(source).to_bytes(options), filetype="pdf") as pdf:
        chars = field_chars(pdf[0])
        actual = field_anchor(chars, alignment)
        assert actual == pytest.approx(200, abs=0.05)


@pytest.mark.parametrize("shaping", [False, True])
@pytest.mark.parametrize("cleared", [False, True])
def test_center_decimal_selection_respects_order_and_clear(shaping, cleared):
    model = aligned_toc_model(3)
    model.sections[0].body.children[0].paragraph_format.tab_stops.tab_stops = [
        ldm.TabStop(position=220, alignment=3, is_clear=cleared),
        ldm.TabStop(position=160, alignment=1, is_clear=True),
        ldm.TabStop(position=140, alignment=1, is_clear=cleared),
    ]
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf") as pdf:
        chars = field_chars(pdf[0])
        if cleared:
            actual = max(c["bbox"][2] for c in chars)
            expected = 300 - 20 - 6
        else:
            actual = (min(c["bbox"][0] for c in chars) + max(c["bbox"][2] for c in chars)) / 2
            expected = 160
        assert actual == pytest.approx(expected, abs=0.05)


@pytest.mark.parametrize("position,alignment", [(13, 1), (260, 0)])
@pytest.mark.parametrize("shaping", [False, True])
def test_near_left_or_right_field_preserves_existing_target(tmp_path, position, alignment, shaping):
    from aspose.words_foss import Document
    from aspose.words_foss.docx_writer import LdmDocxWriter

    model = aligned_toc_model(alignment)
    model.sections[0].body.children[0].paragraph_format.tab_stops.tab_stops[0].position = position
    source = tmp_path / "outside-left.docx"
    LdmDocxWriter().write(model, source)
    target = tmp_path / "existing.pdf"
    target.write_bytes(b"keep-existing")
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pytest.raises(ValueError, match="TOC .*outside"):
        Document(source).save(target, options)
    assert target.read_bytes() == b"keep-existing"


@pytest.mark.parametrize("shaping", [False, True])
def test_earlier_left_stop_takes_priority_over_later_right(shaping):
    model = aligned_toc_model(0)
    model.sections[0].body.children[0].paragraph_format.tab_stops.tab_stops = [
        ldm.TabStop(position=180, alignment=2),
        ldm.TabStop(position=80, alignment=0, leader=1),
    ]
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf") as pdf:
        assert field_anchor(field_chars(pdf[0]), 0) == pytest.approx(100, abs=0.05)


@pytest.mark.parametrize("location", [0, 1])
@pytest.mark.parametrize("shaping", [False, True])
def test_left_field_in_page_band_uses_page_text_origin(location, shaping):
    model = aligned_toc_model(0)
    section = model.sections[0]
    paragraph = section.body.children[0]
    section.body.children = [ldm.Paragraph(children=[ldm.Run(text="BODY")])]
    section.headers_footers = [ldm.HeaderFooter(header_footer_type=location, children=[paragraph])]
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf") as pdf:
        assert len(pdf) == 1
        assert field_anchor(field_chars(pdf[0]), 0) == pytest.approx(200, abs=0.05)
        assert "BODY" in pdf[0].get_text()
