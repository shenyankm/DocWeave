"""Owned repeat-header metadata survives native-compatible HTML row groups."""

import warnings
from io import BytesIO
from xml.etree import ElementTree as ET

import pytest

from aspose.words_foss import Document, MarkdownLoadOptions, SaveFormat, saving
from aspose.words_foss import light_document_model as ldm

FLAGS = [
    (False, False, False),
    (True, False, False),
    (True, True, False),
    (False, True, False),
    (True, False, True),
    (True, True, True),
]


def make_table(flags, rtl=False):
    return ldm.Table(
        bidi=rtl,
        rows=[
            ldm.Row(
                row_format=ldm.RowFormat(heading_format=flag),
                cells=[
                    ldm.Cell(
                        paragraphs=[ldm.Paragraph(children=[ldm.Run(text=f"R{i}C{j}")])]
                    )
                    for j in range(2)
                ],
            )
            for i, flag in enumerate(flags)
        ],
    )


def first_table(doc):
    return next(
        node
        for node in doc.light_document_model.sections[0].body.children
        if isinstance(node, ldm.Table)
    )


def export(doc):
    options = saving.MarkdownSaveOptions()
    options.export_as_html = saving.MarkdownExportAsHtml.TABLES
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return doc.to_bytes(options)


@pytest.mark.parametrize("flags", FLAGS)
@pytest.mark.parametrize("rtl", [False, True])
@pytest.mark.parametrize("save_format", [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_real_saved_header_metadata_exports_contiguous_head_only(
    flags, rtl, save_format
):
    doc = Document(BytesIO(b"fixture"), MarkdownLoadOptions())
    doc.light_document_model.sections[0].body.children[:] = [make_table(flags, rtl)]
    doc = Document(BytesIO(doc.to_bytes(SaveFormat.DOCX)))
    assert [row.row_format.heading_format for row in first_table(doc).rows] == list(
        flags
    )
    before = doc.light_document_model.model_dump()
    raw = export(doc)
    assert doc.light_document_model.model_dump() == before
    root = ET.fromstring(raw.decode().strip())
    header_count = next((i for i, flag in enumerate(flags) if not flag), len(flags))
    assert len(root.findall("thead/tr")) == header_count
    assert len(root.findall("tbody/tr")) == (
        len(flags) - header_count if header_count else 0
    )
    assert len(root.findall("tr")) == (0 if header_count else len(flags))
    assert root.get("dir") == ("rtl" if rtl else None)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        loaded = Document(BytesIO(raw), MarkdownLoadOptions())
    cold = Document(BytesIO(loaded.to_bytes(save_format)))
    table = first_table(cold)
    assert [row.row_format.heading_format for row in table.rows] == [
        i < header_count for i in range(3)
    ]
    assert [[cell.paragraphs[0].text for cell in row.cells] for row in table.rows] == [
        [f"R{i}C{j}" for j in range(2)] for i in range(3)
    ]


@pytest.mark.parametrize("save_format", [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_nested_repeat_header_groups_are_independent(save_format):
    outer = make_table((True, False, False))
    inner = make_table((True, True, False))
    outer.rows[1].cells[0] = ldm.Cell(
        children=[
            ldm.Paragraph(children=[ldm.Run(text="BEFORE")]),
            inner,
            ldm.Paragraph(children=[ldm.Run(text="AFTER")]),
        ]
    )
    doc = Document(BytesIO(b"fixture"), MarkdownLoadOptions())
    doc.light_document_model.sections[0].body.children[:] = [outer]
    raw = export(doc)
    root = ET.fromstring(raw.decode().strip())
    assert len(root.findall("thead/tr")) == 1
    nested = root.find(".//tbody/tr/td/table")
    assert len(nested.findall("thead/tr")) == 2
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        loaded = Document(BytesIO(raw), MarkdownLoadOptions())
    cold = Document(BytesIO(loaded.to_bytes(save_format)))
    table = first_table(cold)
    assert [row.row_format.heading_format for row in table.rows] == [True, False, False]
    cell = table.rows[1].cells[0]
    assert [type(node).__name__ for node in cell.children] == [
        "Paragraph",
        "Table",
        "Paragraph",
    ]
    assert [row.row_format.heading_format for row in cell.tables[0].rows] == [
        True,
        True,
        False,
    ]
    assert [p.text for p in cell.paragraphs] == ["BEFORE", "AFTER"]
