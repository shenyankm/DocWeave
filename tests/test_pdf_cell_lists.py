"""Cell lists retain visible markers and logical items across row fragments."""

from io import BytesIO

import pymupdf
from pypdf import PdfReader
import pytest

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer import LdmPdfWriter
from aspose.words_foss.saving import PdfSaveOptions
from tests.test_pdf_structure_pages import assert_page_tags


def cell_list_model(long=False, explicit=True):
    children = []
    for index, level in enumerate([0, 1, 0]):
        text = f"CELLITEM{index}"
        if long and index == 0:
            text += "\n" + "\n".join(f"LINE{i:02}" for i in range(35))
        children.append(
            ldm.Paragraph(
                children=[ldm.Run(text=text, font=ldm.Font(size=9))],
                list_format=ldm.ListFormat(
                    is_list_item=True, list_id=1, list_level_number=level
                ),
                list_label=ldm.ListLabel(
                    label_string=f"{index + 1}." if explicit else ""
                ),
            )
        )
    children.append(
        ldm.Paragraph(children=[ldm.Run(text="AFTER", font=ldm.Font(size=9))])
    )
    table = ldm.Table(
        rows=[
            ldm.Row(
                cells=[
                    ldm.Cell(paragraphs=children),
                    ldm.Cell(
                        paragraphs=[ldm.Paragraph(children=[ldm.Run(text="OTHER")])]
                    ),
                ]
            )
        ]
    )
    return ldm.Document(
        lists=[
            ldm.DocList(
                list_id=1,
                is_multi_level=True,
                list_levels=[
                    ldm.ListLevel(number_format="%1."),
                    ldm.ListLevel(number_format="%2."),
                ],
            )
        ],
        sections=[
            ldm.Section(
                page_setup=ldm.PageSetup(
                    page_width=350,
                    page_height=180,
                    left_margin=20,
                    right_margin=20,
                    top_margin=20,
                    bottom_margin=20,
                ),
                body=ldm.Body(children=[table]),
            )
        ],
    )


def children(node):
    return [
        kid.get_object()
        for kid in node["/K"]
        if not isinstance(kid, int) and kid.get_object().get("/Type") == "/StructElem"
    ]


@pytest.mark.parametrize("long", [False, True])
@pytest.mark.parametrize("explicit", [False, True])
@pytest.mark.parametrize("shaping", [False, True])
def test_cell_list_markers_hierarchy_and_page_mapping(long, explicit, shaping):
    doc = cell_list_model(long, explicit)
    snapshot = doc.model_dump()
    opts = PdfSaveOptions()
    opts.export_document_structure = True
    opts.text_shaping = shaping
    raw = LdmPdfWriter(opts).write_to_bytes(doc)
    control_options = PdfSaveOptions()
    control_options.text_shaping = shaping
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        text = "".join(page.get_text() for page in pdf)
        expected = ["1.", "2.", "3."] if explicit else ["1.", "1.", "2."]
        for label in set(expected):
            assert text.count(label) == expected.count(label)
        assert all(text.count(f"CELLITEM{i}") == 1 for i in range(3))
        if long:
            assert len(pdf) > 1
        with pymupdf.open(
            stream=LdmPdfWriter(control_options).write_to_bytes(doc), filetype="pdf"
        ) as control:
            assert len(pdf) == len(control)
            for actual, reference in zip(pdf, control, strict=True):
                assert actual.get_text("words") == reference.get_text("words")
                assert actual.get_pixmap().samples == reference.get_pixmap().samples
    reader = PdfReader(BytesIO(raw))
    root = reader.trailer["/Root"]["/StructTreeRoot"]["/K"][0].get_object()
    table = children(root)[0]
    cells = children(children(table)[0])
    assert [node["/S"] for node in children(cells[0])] == ["/L", "/P"]
    assert [node["/S"] for node in children(cells[1])] == ["/P"]
    items = []

    def walk(node, depth):
        for item in children(node):
            assert item["/S"] == "/LI"
            items.append(depth)
            parts = children(item)
            assert [p["/S"] for p in parts] == ["/Lbl", "/LBody"]
            for nested in children(parts[1]):
                if nested["/S"] == "/L":
                    walk(nested, depth + 1)

    walk(children(cells[0])[0], 0)
    assert items == [0, 1, 0]
    if long:
        assert_page_tags(raw)
    assert doc.model_dump() == snapshot


def test_public_docx_cell_numbering_reaches_pdf():
    import aspose.words_foss as aw
    from aspose.words_foss.docx_writer import LdmDocxWriter

    original = cell_list_model(explicit=False)
    doc = aw.Document(BytesIO(LdmDocxWriter().write_to_bytes(original)))
    options = PdfSaveOptions()
    options.export_document_structure = True
    raw = doc.to_bytes(options)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        text = "".join(page.get_text() for page in pdf)
        assert text.count("1.") == 2 and text.count("2.") == 1
        assert all(text.count(f"CELLITEM{i}") == 1 for i in range(3))
    reader = PdfReader(BytesIO(raw))
    root = reader.trailer["/Root"]["/StructTreeRoot"]["/K"][0].get_object()
    roles = []

    def walk(node):
        roles.append(node["/S"])
        for kid in children(node):
            walk(kid)

    walk(root)
    assert roles.count("/LI") == 3 and roles.count("/Lbl") == 3


@pytest.mark.parametrize(
    "kind",
    ["nested_table", "image", "interruption", "missing_label", "repeated_header"],
)
def test_cell_list_combinations(kind):
    from PIL import Image

    doc = cell_list_model()
    grid = doc.sections[0].body.children[0]
    cell = grid.rows[0].cells[0]
    if kind == "nested_table":
        doc.sections[0].body.children = [
            ldm.Table(rows=[ldm.Row(cells=[ldm.Cell(tables=[grid])])])
        ]
    elif kind == "image":
        stream = BytesIO()
        Image.new("RGB", (10, 10), "blue").save(stream, format="PNG")
        cell.paragraphs[0]._children.extend(
            [
                ldm.Shape(
                    has_image=True,
                    width=10,
                    height=10,
                    alternative_text="BLUE",
                    image_data=ldm.ImageData(image_bytes=stream.getvalue()),
                ),
                ldm.Run(text="AFTERPIC"),
            ]
        )
    elif kind == "interruption":
        cell.paragraphs.insert(1, ldm.Paragraph(children=[ldm.Run(text="BREAK")]))
    elif kind == "missing_label":
        doc.lists = []
        for p in cell.paragraphs:
            if p.list_label:
                p.list_label.label_string = ""
    else:
        grid.rows[0].row_format.heading_format = True
        grid.rows.append(
            ldm.Row(
                cells=[
                    ldm.Cell(
                        paragraphs=[
                            ldm.Paragraph(
                                children=[
                                    ldm.Run(
                                        text="\n".join(f"ROW{i:02}" for i in range(45)),
                                        font=ldm.Font(size=9),
                                    )
                                ]
                            )
                        ]
                    ),
                    ldm.Cell(),
                ]
            )
        )
    snapshot = doc.model_dump()
    options = PdfSaveOptions()
    options.export_document_structure = True
    raw = LdmPdfWriter(options).write_to_bytes(doc)
    reader = PdfReader(BytesIO(raw))
    root = reader.trailer["/Root"]["/StructTreeRoot"]["/K"][0].get_object()
    nodes = []

    def walk(node):
        nodes.append(node)
        for kid in children(node):
            walk(kid)

    walk(root)
    assert sum(n["/S"] == "/LI" for n in nodes) == 3
    assert sum(n["/S"] == "/Lbl" for n in nodes) == (
        0 if kind == "missing_label" else 3
    )
    if kind == "interruption":
        assert sum(n["/S"] == "/L" for n in nodes) == 3
    if kind == "image":
        figure = next(n for n in nodes if n["/S"] == "/Figure")
        assert figure["/Alt"] == "BLUE"
        assert figure["/P"]["/P"]["/S"] == "/LBody"
    if kind == "repeated_header":
        assert_page_tags(raw)
    with pymupdf.open(stream=raw, filetype="pdf") as pdf:
        text = "".join(page.get_text() for page in pdf)
        assert "AFTER" in text
        assert text.count("CELLITEM0") == (len(pdf) if kind == "repeated_header" else 1)
        with pymupdf.open(
            stream=LdmPdfWriter().write_to_bytes(doc), filetype="pdf"
        ) as control:
            assert len(pdf) == len(control)
            for actual, reference in zip(pdf, control, strict=True):
                assert actual.get_text("words") == reference.get_text("words")
                assert actual.get_pixmap().samples == reference.get_pixmap().samples
    assert doc.model_dump() == snapshot


@pytest.mark.parametrize("shaping", [False, True])
def test_cell_list_labels_and_body_have_separate_content(shaping):
    from pypdf import PdfWriter

    options = PdfSaveOptions()
    options.export_document_structure = True
    options.text_shaping = shaping
    raw = LdmPdfWriter(options).write_to_bytes(cell_list_model())
    extracted = {}
    for role in ("/Lbl", "/LBody"):
        writer = PdfWriter(clone_from=PdfReader(BytesIO(raw)))
        nums = writer.root_object["/StructTreeRoot"]["/ParentTree"]["/Nums"]
        arrays = {
            int(nums[i]): nums[i + 1].get_object() for i in range(0, len(nums), 2)
        }
        for page in writer.pages:
            content = page.get_contents()
            stack = []
            kept = []
            for args, op in content.operations:
                if op == b"BDC":
                    node = arrays[int(page["/StructParents"])][
                        int(args[1]["/MCID"])
                    ].get_object()
                    roles = []
                    while node.get("/S"):
                        roles.append(node["/S"])
                        if node["/S"] in ("/Lbl", "/LBody"):
                            break
                        node = node["/P"]
                    stack.append(role in roles)
                elif op == b"EMC":
                    stack.pop()
                if op not in (b"Tj", b"TJ") or stack and stack[-1]:
                    kept.append((args, op))
            content.operations = kept
            page.replace_contents(content)
        stream = BytesIO()
        writer.write(stream)
        extracted[role] = "".join(
            page.extract_text() for page in PdfReader(stream).pages
        )
    assert all(
        f"{i + 1}." in extracted["/Lbl"] and f"CELLITEM{i}" not in extracted["/Lbl"]
        for i in range(3)
    )
    assert all(
        f"CELLITEM{i}" in extracted["/LBody"] and f"{i + 1}." not in extracted["/LBody"]
        for i in range(3)
    )


@pytest.mark.parametrize("failure", ["drawing", "closing"])
def test_cell_list_failure_preserves_output_and_recovers(
    tmp_path, monkeypatch, failure
):
    from fpdf import FPDF
    from aspose.words_foss.pdf_writer.run_renderer import RunRenderer

    options = PdfSaveOptions()
    options.export_document_structure = True
    writer = LdmPdfWriter(options)
    output = tmp_path / "keep.pdf"
    output.write_bytes(b"KEEP")
    captured = []
    files = []
    original_render = RunRenderer._render_segment_row
    original_out = FPDF._out

    def render(renderer, pdf, *args, **kwargs):
        captured.append(pdf)
        files.extend(
            font.ttfont.reader.file
            for font in pdf.fonts.values()
            if font.ttfont.reader is not None
        )
        if failure == "drawing":
            raise RuntimeError("cell list drawing failed")
        return original_render(renderer, pdf, *args, **kwargs)

    def close(pdf, text):
        if failure == "closing" and text == "EMC" and captured:
            raise RuntimeError("cell list closing failed")
        return original_out(pdf, text)

    monkeypatch.setattr(RunRenderer, "_render_segment_row", render)
    monkeypatch.setattr(FPDF, "_out", close)
    with pytest.raises(RuntimeError, match=f"cell list {failure} failed"):
        writer.write(cell_list_model(), output)
    assert (
        output.read_bytes() == b"KEEP" and files and all(file.closed for file in files)
    )
    assert all(
        parent is captured[0].struct_builder.doc_struct_elem
        for parent in captured[0].struct_builder.parents.values()
    )
    monkeypatch.setattr(RunRenderer, "_render_segment_row", original_render)
    monkeypatch.setattr(FPDF, "_out", original_out)
    assert_page_tags(writer.write_to_bytes(cell_list_model(long=True)))


def test_oversized_cell_marker_fails_without_replacing_output(tmp_path):
    doc = cell_list_model()
    doc.sections[0].body.children[0].rows[0].cells[0].paragraphs[
        0
    ].list_label.label_string = "W" * 300
    output = tmp_path / "keep.pdf"
    output.write_bytes(b"KEEP")
    with pytest.raises(ValueError, match="list marker is wider"):
        LdmPdfWriter().write(doc, output)
    assert output.read_bytes() == b"KEEP"
