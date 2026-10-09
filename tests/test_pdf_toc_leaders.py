"""Directory leaders are bounded vector decoration, not extracted body text."""

from io import BytesIO

import pymupdf
import pytest
from pypdf import PdfReader

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.saving import PdfSaveOptions


def leader_model(leader, columns=1, long=False):
    title = " ".join(f"T{i:03}" for i in range(48)) if long else "TITLE"
    pf = ldm.ParagraphFormat(style_name="TOC 1", left_indent=12, right_indent=6)
    pf.tab_stops.add(100 if columns == 2 else 180, alignment=2, leader=leader)
    paragraph = ldm.Paragraph(
        paragraph_format=pf,
        children=[
            ldm.Run(text=title, font=ldm.Font(size=12)),
            ldm.Run(text="\t123", font=ldm.Font(size=16)),
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
                body=ldm.Body(children=[paragraph]),
            )
        ]
    )


@pytest.mark.parametrize("leader", [1, 2, 3, 4, 5])
@pytest.mark.parametrize("shaping", [False, True])
@pytest.mark.parametrize("columns", [1, 2])
def test_toc_leader_is_vector_artifact_and_keeps_text_and_positions(
    leader, shaping, columns
):
    model = leader_model(leader, columns)
    original = model.model_dump()
    options = PdfSaveOptions()
    options.text_shaping = shaping
    options.export_document_structure = True
    raw = LdmPdfWriter(options).write_to_bytes(model)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        assert len(pdf) == 1
        words = {w[4]: w for w in pdf[0].get_text("words")}
        assert set(words) == {"TITLE", "123"}
        drawings = pdf[0].get_drawings()
        assert len(drawings) == 1
        line = drawings[0]
        assert line["rect"].x0 > words["TITLE"][2]
        assert line["rect"].x1 < words["123"][0]
        assert line["rect"].x0 - line["width"] / 2 > words["TITLE"][2]
        assert line["rect"].x1 + line["width"] / 2 < words["123"][0]
        assert line["width"] > 0
        assert words["123"][2] == pytest.approx(120 if columns == 2 else 200, abs=0.05)
    reader = PdfReader(BytesIO(raw))
    stack = []
    artifact_strokes = 0
    for operands, operator in reader.pages[0].get_contents().operations:
        if operator in (b"BMC", b"BDC"):
            stack.append(operands[0])
        elif operator == b"EMC":
            stack.pop()
        elif operator == b"S":
            assert "/Artifact" in stack
            artifact_strokes += 1
        elif operator in (b"Tj", b"TJ"):
            assert "/Artifact" not in stack
    assert not stack and artifact_strokes == 1
    assert model.model_dump() == original


@pytest.mark.parametrize("shaping", [False, True])
def test_long_toc_leader_is_only_on_last_row_and_same_column(shaping):
    model = leader_model(1, columns=2, long=True)
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(
        stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf"
    ) as pdf:
        assert len(pdf) > 1
        text = (
            "".join(page.get_text() for page in pdf).replace("\n", "").replace(" ", "")
        )
        assert text == "".join(f"T{i:03}" for i in range(48)) + "123"
        drawings = [(i, d) for i, page in enumerate(pdf) for d in page.get_drawings()]
        assert len(drawings) == 1
        i, line = drawings[0]
        words = {w[4]: w for w in pdf[i].get_text("words")}
        assert "T047" in words and "123" in words
        assert line["rect"].x0 > words["T047"][2]
        assert line["rect"].x1 < words["123"][0]


def test_toc_leader_styles_have_distinct_patterns_and_vertical_positions():
    lines = {}
    for leader in (1, 2, 3, 4, 5):
        raw = LdmPdfWriter().write_to_bytes(leader_model(leader))
        with pymupdf.open(stream=raw, filetype="pdf") as pdf:
            lines[leader] = pdf[0].get_drawings()[0]
    assert lines[4]["width"] > lines[3]["width"]
    assert lines[3]["dashes"] == lines[4]["dashes"] == "[] 0"
    assert lines[1]["dashes"] != lines[2]["dashes"]
    assert lines[1]["lineCap"][0] == lines[5]["lineCap"][0] == 1
    assert lines[2]["lineCap"][0] == 0
    assert (
        lines[5]["rect"].y0
        < lines[2]["rect"].y0
        < lines[1]["rect"].y0
        < lines[3]["rect"].y0
    )


@pytest.mark.parametrize("leader", [1, 2, 3, 4, 5])
def test_toc_leader_tagging_does_not_change_pixels(leader):
    model = leader_model(leader)
    options = PdfSaveOptions()
    options.export_document_structure = True
    tagged = LdmPdfWriter(options).write_to_bytes(model)
    untagged = LdmPdfWriter().write_to_bytes(model)
    with (
        pymupdf.open(stream=tagged, filetype="pdf") as actual,
        pymupdf.open(stream=untagged, filetype="pdf") as control,
    ):
        assert actual[0].get_text("words") == control[0].get_text("words")
        assert actual[0].get_pixmap().samples == control[0].get_pixmap().samples


def test_toc_leader_restores_graphics_state_and_keeps_following_text(monkeypatch):
    model = leader_model(4)
    model.sections[0].body.children[0]._children[
        -1
    ].font.color = "Color [A=255, R=255, G=0, B=0]"
    model.sections[0].body.children.append(
        ldm.Paragraph(children=[ldm.Run(text="AFTER")])
    )
    writer = LdmPdfWriter()
    original = writer._run_renderer._draw_toc_leader
    seen = []

    def record(pdf, *args):
        before = (pdf.line_width, pdf.dash_pattern.copy(), pdf.draw_color)
        original(pdf, *args)
        assert (pdf.line_width, pdf.dash_pattern, pdf.draw_color) == before
        seen.append(before)

    monkeypatch.setattr(writer._run_renderer, "_draw_toc_leader", record)
    raw = writer.write_to_bytes(model)
    assert len(seen) == 1
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        assert pdf[0].get_drawings()[0]["color"] == (1, 0, 0)
        spans = [
            s
            for b in pdf[0].get_text("dict")["blocks"]
            for line in b.get("lines", [])
            for s in line["spans"]
        ]
        assert next(s for s in spans if "AFTER" in s["text"])["color"] == 0


@pytest.mark.parametrize("leader", [1, 2, 3, 4, 5])
@pytest.mark.parametrize("shaping", [False, True])
def test_public_docx_toc_leader_is_drawn_without_text_noise(tmp_path, leader, shaping):
    from aspose.words_foss import Document
    from aspose.words_foss.docx_writer import LdmDocxWriter

    source = tmp_path / "leader.docx"
    LdmDocxWriter().write(leader_model(leader), source)
    options = PdfSaveOptions()
    options.text_shaping = shaping
    with pymupdf.open(stream=Document(source).to_bytes(options), filetype="pdf") as pdf:
        assert {w[4] for w in pdf[0].get_text("words")} == {"TITLE", "123"}
        assert len(pdf[0].get_drawings()) == 1


@pytest.mark.parametrize("leader", [0, 6])
def test_absent_or_unknown_toc_leader_keeps_existing_output(leader):
    with pymupdf.open(
        stream=LdmPdfWriter().write_to_bytes(leader_model(leader)), filetype="pdf"
    ) as pdf:
        assert {w[4] for w in pdf[0].get_text("words")} == {"TITLE", "123"}
        assert not pdf[0].get_drawings()


@pytest.mark.parametrize("location", [0, 1])
@pytest.mark.parametrize("leader", [1, 2, 3, 4, 5])
def test_page_band_toc_leader_preserves_body_flow(location, leader):
    model = leader_model(leader)
    section = model.sections[0]
    paragraph = section.body.children[0]
    section.body.children = [ldm.Paragraph(children=[ldm.Run(text="BODY")])]
    section.headers_footers = [
        ldm.HeaderFooter(header_footer_type=location, children=[paragraph])
    ]
    options = PdfSaveOptions()
    options.export_document_structure = True
    with pymupdf.open(
        stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf"
    ) as pdf:
        assert len(pdf) == 1
        words = {w[4]: w for w in pdf[0].get_text("words")}
        assert set(words) == {"TITLE", "123", "BODY"}
        next(iter(paragraph.paragraph_format.tab_stops)).leader = 0
        with pymupdf.open(
            stream=LdmPdfWriter(options).write_to_bytes(model), filetype="pdf"
        ) as baseline:
            assert words["BODY"] == next(
                w for w in baseline[0].get_text("words") if w[4] == "BODY"
            )
        lines = pdf[0].get_drawings()
        assert len(lines) == 1
        assert lines[0]["rect"].x0 > words["TITLE"][2]
        assert lines[0]["rect"].x1 < words["123"][0]
