"""Grid fidelity, mixed text, images, merges, nested tables and real pagination."""

from io import BytesIO
from zipfile import ZipFile

from PIL import Image
import pymupdf
import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer.diagnostics import PdfContentLossWarning


def paragraph(text, **font):
    return ldm.Paragraph(children=[ldm.Run(text=text, font=ldm.Font(**font))])


def cell(text, **formatting):
    return ldm.Cell(paragraphs=[paragraph(text)], cell_format=ldm.CellFormat(**formatting))


def row(*cells, heading=False):
    return ldm.Row(cells=list(cells), row_format=ldm.RowFormat(heading_format=heading))


def table(*rows):
    borders = [ldm.Border(line_style=1, line_width=0.5) for _ in range(6)]
    for item in rows:
        item.row_format.borders = borders
    return ldm.Table(rows=list(rows))


def convert(grid, *before, page_height=360):
    model = ldm.Document(sections=[ldm.Section(
        page_setup=ldm.PageSetup(page_width=360, page_height=page_height,
                               left_margin=24, right_margin=24, top_margin=24, bottom_margin=24),
        body=ldm.Body(children=[*before, grid]),
    )])
    doc = aw.Document(BytesIO(LdmDocxWriter().write_to_bytes(model)))
    return doc, pymupdf.open(stream=doc.to_bytes("pdf"), filetype="pdf")


def assert_in_bounds(pdf):
    for page in pdf:
        assert page.get_pixmap().samples
        for block in page.get_text("dict")["blocks"]:
            x0, y0, x1, y1 = block["bbox"]
            assert x0 >= 23 and x1 <= page.rect.width - 23
            assert y0 >= 22 and y1 <= page.rect.height - 22


def test_rich_cell_retains_fonts_and_clickable_link():
    rich = ldm.Cell(paragraphs=[ldm.Paragraph(children=[
        ldm.Run(text="普通", font=ldm.Font(size=11)),
        ldm.Run(text="粗体", font=ldm.Font(size=16, bold=True)),
        ldm.Run(text="[链接](https://example.com)", font=ldm.Font(italic=True)),
    ])])
    _, pdf = convert(table(row(rich, cell("右侧"))))
    with pdf:
        assert "普通粗体链接" in pdf[0].get_text().replace("\n", "")
        spans = [s for b in pdf[0].get_text("dict")["blocks"] for line in b.get("lines", []) for s in line["spans"]]
        bold = next(s for s in spans if "粗体" in s["text"])
        assert bold["size"] == pytest.approx(16) and bold["flags"] & 16
        assert any(link.get("uri") == "https://example.com" for link in pdf[0].get_links())
        assert len(pdf[0].get_drawings()) >= 4
        assert_in_bounds(pdf)


def test_image_stays_in_its_cell_not_flattened():
    data = BytesIO()
    Image.new("RGB", (60, 30), "red").save(data, format="PNG")
    image_cell = ldm.Cell(paragraphs=[ldm.Paragraph(children=[
        ldm.Run(text="左图"), ldm.Shape(has_image=True, is_inline=True, width=72, height=36,
            image_data=ldm.ImageData(image_type=ldm.ImageData.from_mime("image/png"), image_bytes=data.getvalue())),
        ldm.Run(text="图后"),
    ])])
    _, pdf = convert(table(row(image_cell, cell("右侧同一行"))))
    with pdf:
        blocks = pdf[0].get_text("dict")["blocks"]
        image = next(b for b in blocks if b["type"] == 1)
        right = next(b for b in blocks if "右侧" in "".join(s["text"] for line in b.get("lines", []) for s in line["spans"]))
        assert image["bbox"][2] < right["bbox"][0]
        assert right["bbox"][1] < image["bbox"][3]
        assert "图后" in pdf[0].get_text()
        assert len(pdf[0].get_drawings()) >= 4
        assert_in_bounds(pdf)


def test_long_table_repeats_header_without_losing_rows():
    grid = table(row(cell("表头"), cell("项目"), heading=True),
                 *(row(cell(f"行{i:03}"), cell("中文数据")) for i in range(65)))
    _, pdf = convert(grid)
    with pdf:
        assert len(pdf) > 1
        text = "".join(page.get_text() for page in pdf)
        for i in range(65):
            assert text.count(f"行{i:03}") == 1
        assert all("表头" in page.get_text() for page in pdf)
        assert_in_bounds(pdf)


@pytest.mark.parametrize("cant_split", [False, True])
def test_taller_than_page_row_splits_and_keeps_every_character(cant_split):
    text = "中文长行" * 180
    long_row = row(cell(text), cell("邻居"))
    long_row.row_format.allow_break_across_pages = not cant_split
    grid = table(row(cell("表头"), cell("项目"), heading=True), long_row)
    if cant_split:
        with pytest.warns(PdfContentLossWarning, match="cantSplit"):
            doc, pdf = convert(grid, page_height=240)
        assert any(d.code == "pdf.table_row_split" for d in doc.diagnostics)
    else:
        _, pdf = convert(grid, page_height=240)
    with pdf:
        assert len(pdf) > 1
        extracted = "".join(page.get_text() for page in pdf).replace("\n", "")
        for extra in ("表头", "项目", "邻居"):
            extracted = extracted.replace(extra, "")
        assert extracted == text
        assert_in_bounds(pdf)


def test_row_moves_intact_when_it_can_fit_a_fresh_page():
    grid = table(row(cell("中文" * 20), cell("邻居")))
    _, pdf = convert(grid, *(paragraph("前文") for _ in range(11)), page_height=240)
    with pdf:
        assert len(pdf) == 2
        assert "邻居" not in pdf[0].get_text() and "邻居" in pdf[1].get_text()
        assert_in_bounds(pdf)


def test_grid_span_round_trip_and_pdf_geometry():
    grid = table(row(cell("合并", grid_span=2)), row(cell("A"), cell("B")))
    doc, pdf = convert(grid)
    assert doc.light_document_model.tables[0].rows[0].cells[0].cell_format.grid_span == 2
    with ZipFile(BytesIO(doc.to_bytes("docx"))) as archive:
        xml = archive.read("word/document.xml").decode()
        assert '<w:gridSpan w:val="2"' in xml
        assert xml.count("<w:gridCol ") == 2
    with pdf:
        assert "合并" in pdf[0].get_text()
        assert_in_bounds(pdf)


def test_nested_table_is_rendered_inside_parent_cell():
    nested = table(row(cell("内部A"), cell("内部B")))
    parent = ldm.Cell(paragraphs=[paragraph("父单元格")], tables=[nested])
    _, pdf = convert(table(row(parent, cell("外侧"))))
    with pdf:
        text = pdf[0].get_text()
        assert all(value in text for value in ("内部A", "内部B", "父单元格", "外侧"))
        assert_in_bounds(pdf)


def test_vertical_merge_does_not_draw_an_internal_horizontal_border():
    grid = table(row(cell("合并", vertical_merge=1), cell("A")),
                 row(cell("", vertical_merge=2), cell("B")))
    _, pdf = convert(grid)
    with pdf:
        horizontals = [item[1:3] for drawing in pdf[0].get_drawings()
                       for item in drawing["items"] if item[0] == "l" and abs(item[1].y - item[2].y) < 0.01]
        left_cell_edges = [(p1, p2) for p1, p2 in horizontals if p2.x < 181]
        assert len({round(p1.y, 2) for p1, _ in left_cell_edges}) == 2
        assert_in_bounds(pdf)


@pytest.mark.parametrize("span", [0, 1025])
def test_untrusted_grid_span_is_bounded(span):
    raw = LdmDocxWriter().write_to_bytes(ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[
        table(row(cell("text", grid_span=2))),
    ]))]))
    source = BytesIO()
    with ZipFile(BytesIO(raw)) as archive, ZipFile(source, "w") as output:
        for name in archive.namelist():
            payload = archive.read(name)
            if name == "word/document.xml":
                payload = payload.replace(b'<w:gridSpan w:val="2"', f'<w:gridSpan w:val="{span}"'.encode())
            output.writestr(name, payload)
    with pytest.raises(ValueError, match="gridSpan"):
        aw.Document(BytesIO(source.getvalue()))


def test_text_after_grid_uses_a_valid_font_and_is_not_clipped():
    model = ldm.Document(sections=[ldm.Section(body=ldm.Body(children=[
        table(row(cell("表格A"), cell("表格B"))), paragraph("表后正文"),
    ]))])
    doc = aw.Document(BytesIO(LdmDocxWriter().write_to_bytes(model)))
    with pymupdf.open(stream=doc.to_bytes("pdf"), filetype="pdf") as pdf:
        assert "表后正文" in pdf[0].get_text()
        assert_in_bounds(pdf)


def test_rotated_text_cell_keeps_its_orientation():
    _, pdf = convert(table(row(cell("VERTICAL", orientation=1), cell("normal"))))
    with pdf:
        lines = [line for block in pdf[0].get_text("dict")["blocks"] for line in block.get("lines", [])]
        vertical = next(line for line in lines if any("VERTICAL" in span["text"] for span in line["spans"]))
        assert abs(vertical["dir"][0]) < 0.01 and abs(vertical["dir"][1]) > 0.99
        assert_in_bounds(pdf)


def test_impossible_header_fails_without_destroying_output(tmp_path):
    source = tmp_path / "source.md"
    source.write_text("text", encoding="utf-8")
    doc = aw.Document(source)
    doc.sections[0].body.children = [table(row(cell("text", top_padding=900), heading=True))]
    output = tmp_path / "existing.pdf"
    output.write_bytes(b"old")
    with pytest.raises(ValueError, match="headers do not fit"):
        doc.save(output)
    assert output.read_bytes() == b"old"
