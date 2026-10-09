"""Body lists retain their text margins across pages and columns."""

import pymupdf
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.saving import PdfSaveOptions
from tests.test_pdf_structure_pages import assert_page_tags


def body_list_model(columns=1, widow=True):
    paras = []
    for i, level in enumerate([0, 1, 0]):
        paras.append(
            ldm.Paragraph(
                children=[
                    ldm.Run(
                        text=f"BODY{i}\n"
                        + "\n".join(f"LINE{i}_{j:02}" for j in range(25)),
                        font=ldm.Font(size=9),
                    )
                ],
                paragraph_format=ldm.ParagraphFormat(
                    left_indent=24 + 12 * level,
                    first_line_indent=-12,
                    right_indent=6,
                    widow_control=widow,
                ),
                list_format=ldm.ListFormat(
                    is_list_item=True, list_id=1, list_level_number=level
                ),
                list_label=ldm.ListLabel(label_string=f"{i + 1}."),
            )
        )
    paras.append(ldm.Paragraph(children=[ldm.Run(text="AFTER", font=ldm.Font(size=9))]))
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
                body=ldm.Body(children=paras),
            )
        ]
    )


@pytest.mark.parametrize("columns", [1, 2])
@pytest.mark.parametrize("widow", [False, True])
@pytest.mark.parametrize("shaping", [False, True])
def test_body_list_hanging_indents_survive_pages_and_columns(columns, widow, shaping):
    from aspose.words_foss.pdf_writer.constants import PT_TO_MM

    doc = body_list_model(columns, widow)
    snapshot = doc.model_dump()
    options = PdfSaveOptions()
    options.text_shaping = shaping
    options.export_document_structure = True
    raw = LdmPdfWriter(options).write_to_bytes(doc)
    col_width = (260 - 15 * (columns - 1)) / columns

    def left(word):
        return 20 + (col_width + 15 if columns == 2 and word[0] > 150 else 0)

    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        words = [word for page in pdf for word in page.get_text("words")]
        for i, level in enumerate([0, 1, 0]):
            body = next(w for w in words if w[4] == f"BODY{i}")
            marker = next(w for w in words if w[4] == f"{i + 1}.")
            assert body[0] - marker[0] == pytest.approx(12, abs=0.05)
            for word in words:
                if word[4] == f"BODY{i}" or word[4].startswith(f"LINE{i}_"):
                    assert word[0] - left(word) == pytest.approx(
                        24 + 12 * level + 1 / PT_TO_MM, abs=0.05
                    )
                    assert word[2] <= left(word) + col_width - 6 + 0.05
            assert all(
                sum(w[4] == f"LINE{i}_{j:02}" for w in words) == 1 for j in range(25)
            )
        after = next(w for w in words if w[4] == "AFTER")
        assert after[0] - left(after) == pytest.approx(1 / PT_TO_MM, abs=0.05)
    assert_page_tags(raw)
    assert doc.model_dump() == snapshot


@pytest.mark.parametrize("columns", [1, 2])
def test_list_marker_stays_with_body_when_widow_control_advances(columns):
    doc = body_list_model(columns)
    section = doc.sections[0]
    prefix = ldm.Paragraph(
        children=[ldm.Run(text="PREFIX", font=ldm.Font(size=9))],
        paragraph_format=ldm.ParagraphFormat(space_after=115),
    )
    section.body.children = [prefix, section.body.children[0]]
    raw = LdmPdfWriter().write_to_bytes(doc)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        matches = {}
        for page in pdf:
            for word in page.get_text("words"):
                if word[4] in ("1.", "BODY0"):
                    matches[word[4]] = (page.number, word)
        assert matches["1."][0] == matches["BODY0"][0]
        assert matches["BODY0"][1][0] - matches["1."][1][0] == pytest.approx(
            12, abs=0.05
        )
        assert matches["BODY0"][1][1] == pytest.approx(matches["1."][1][1], abs=0.05)


@pytest.mark.parametrize("invalid", ["indent", "marker"])
def test_invalid_body_list_geometry_preserves_output(tmp_path, invalid):
    doc = body_list_model()
    para = doc.sections[0].body.children[0]
    if invalid == "indent":
        para.paragraph_format.left_indent = 500
    else:
        para.list_label.label_string = "W" * 300
    output = tmp_path / "keep.pdf"
    output.write_bytes(b"KEEP")
    with pytest.raises(
        ValueError, match="No usable text width after body list indents"
    ):
        LdmPdfWriter().write(doc, output)
    assert output.read_bytes() == b"KEEP"


def test_default_body_list_levels_indent_first_and_continuation_lines():
    from aspose.words_foss.pdf_writer.constants import PT_TO_MM

    doc = body_list_model()
    for para in doc.sections[0].body.children[:3]:
        para.paragraph_format.left_indent = 0
        para.paragraph_format.first_line_indent = 0
    with pymupdf.open(stream=LdmPdfWriter().write_to_bytes(doc), filetype="pdf") as pdf:
        words = [word for page in pdf for word in page.get_text("words")]
        starts = [
            next(word[0] for word in words if word[4] == f"BODY{i}") for i in range(3)
        ]
        assert starts[1] - starts[0] == pytest.approx(5 / PT_TO_MM, abs=0.05)
        for i in range(3):
            assert all(
                word[0] == pytest.approx(starts[i], abs=0.05)
                for word in words
                if word[4].startswith(f"LINE{i}_")
            )


def test_public_docx_body_list_hanging_indent_reaches_pdf():
    from io import BytesIO
    import aspose.words_foss as aw
    from aspose.words_foss.docx_writer import LdmDocxWriter
    from aspose.words_foss.pdf_writer.constants import PT_TO_MM

    doc = body_list_model(2)
    doc.lists = [
        ldm.DocList(
            list_id=1,
            is_multi_level=True,
            list_levels=[
                ldm.ListLevel(number_format="%1."),
                ldm.ListLevel(number_format="%2."),
            ],
        )
    ]
    options = PdfSaveOptions()
    options.export_document_structure = True
    raw = aw.Document(BytesIO(LdmDocxWriter().write_to_bytes(doc))).to_bytes(options)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        words = [word for page in pdf for word in page.get_text("words")]
        for i, level in enumerate([0, 1, 0]):
            for word in words:
                if word[4] == f"BODY{i}" or word[4].startswith(f"LINE{i}_"):
                    base = 157.5 if word[0] > 150 else 20
                    assert word[0] - base == pytest.approx(
                        24 + 12 * level + 1 / PT_TO_MM, abs=0.05
                    )
    assert_page_tags(raw)
