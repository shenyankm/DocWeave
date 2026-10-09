"""Point dimensions preserve their own layer and match frozen native edits."""

import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from aspose.words_foss import DocxDocument

ROOT = Path(__file__).parents[1]
BENCHMARKS = ROOT / "docs/benchmarks"
REPORT = json.loads((BENCHMARKS / "paragraph-dimensions-26.9.json").read_text())
LIMITS = json.loads((BENCHMARKS / "paragraph-spacing-limits-26.9.json").read_text())
PROPERTIES = ("left_indent", "right_indent", "first_line_indent", "space_before", "space_after")


def paragraph(document):
    return next(p for p in document.body.paragraphs if "IMPORT" in p.text)


def snapshot(document):
    return {name: {prop: getattr(fmt, prop) for prop in PROPERTIES}
            for name, fmt in (("base", document.styles.get_by_name("Base").paragraph_format),
                              ("derived", document.styles.get_by_name("Derived").paragraph_format),
                              ("paragraph", paragraph(document).effective_paragraph_format))}


def load(name, report=REPORT):
    with ZipFile(BENCHMARKS / report["corpus"]) as archive:
        return DocxDocument(BytesIO(archive.read(name)))


@pytest.mark.parametrize("row", REPORT["records"] + LIMITS["records"], ids=lambda row: f'{row["input"]}-{row["target"]}-{row["value"]}')
def test_dimensions_match_native_edit_and_reopen(row):
    document = load(row["input"], LIMITS if row in LIMITS["records"] else REPORT)
    assert snapshot(document) == row["before_edit"]
    before = document.to_bytes()
    node = document.styles.get_by_name("Derived") if row["target"] == "style" else paragraph(document)
    if row["error"]:
        with pytest.raises(RuntimeError):
            setattr(node.paragraph_format, row["property"], row["value"])
        assert document.to_bytes() == before
        return
    setattr(node.paragraph_format, row["property"], row["value"])
    assert snapshot(document) == row["after_edit"]
    live = document.part_xml("word/styles.xml"), document.part_xml("word/document.xml")
    saved = document.to_bytes()
    assert snapshot(document) == row["after_save_live"]
    assert (document.part_xml("word/styles.xml"), document.part_xml("word/document.xml")) == live
    assert snapshot(DocxDocument(BytesIO(saved))) == row["after_reopen"]
    model = document.to_light_document()
    actual = next(p for p in model.sections[0].body.paragraphs if "IMPORT" in p.text)
    assert {prop: getattr(actual.paragraph_format, prop) for prop in PROPERTIES} == row["after_edit"]["paragraph"]
    direct = node.direct_paragraph_format if row["target"] == "style" else node.paragraph_format
    setattr(direct, row["property"], None)
    assert getattr(direct, row["property"]) is None
    assert snapshot(DocxDocument(BytesIO(document.to_bytes()))) == snapshot(document)


@pytest.mark.parametrize("row", [r for r in REPORT["setter_errors"] if r["value"] is not None])
def test_wrong_dimension_types_fail_without_mutation(row):
    document = load(row["input"])
    before = document.to_bytes()
    node = document.styles.get_by_name("Derived") if row["target"] == "style" else paragraph(document)
    with pytest.raises(TypeError):
        setattr(node.paragraph_format, row["property"], row["value"])
    assert document.to_bytes() == before


@pytest.mark.parametrize("prop", PROPERTIES)
@pytest.mark.parametrize("value", [float("inf"), float("nan"), 1e100])
def test_nonfinite_or_overflow_dimensions_fail_without_mutation(prop, value):
    document = load("left_indent/0.docx")
    before = document.to_bytes()
    error = RuntimeError if prop.startswith("space_") and value > 1584 else ValueError
    with pytest.raises(error):
        setattr(paragraph(document).paragraph_format, prop, value)
    assert document.to_bytes() == before


@pytest.mark.parametrize("prop", PROPERTIES)
def test_resolved_style_rejects_none_and_clearing_restores_base(prop):
    document = load(f"{prop}/5.docx")
    style = document.styles.get_by_name("Derived")
    before = document.to_bytes()
    with pytest.raises(TypeError):
        setattr(style.paragraph_format, prop, None)
    assert document.to_bytes() == before
    setattr(style.direct_paragraph_format, prop, None)
    assert getattr(style.paragraph_format, prop) == getattr(document.styles.get_by_name("Base").paragraph_format, prop)


def with_direct_properties(properties):
    with ZipFile(BENCHMARKS / REPORT["corpus"]) as source:
        raw = source.read("first_line_indent/0.docx")
    output = BytesIO()
    with ZipFile(BytesIO(raw)) as source, ZipFile(output, "w") as archive:
        for name in source.namelist():
            data = source.read(name)
            if name == "word/document.xml":
                data = data.replace(b'<w:pStyle w:val="Derived"/>', b'<w:pStyle w:val="Derived"/>' + properties.encode())
            archive.writestr(name, data)
    return DocxDocument(BytesIO(output.getvalue()))


def test_hanging_precedes_firstline_and_edits_keep_sibling_attributes():
    document = with_direct_properties('<w:ind w:firstLine="240" w:hanging="120" w:left="80"/>')
    p = paragraph(document)
    assert p.paragraph_format.first_line_indent == p.effective_paragraph_format.first_line_indent == -6
    model = document.to_light_document()
    assert next(p for p in model.sections[0].body.paragraphs if "IMPORT" in p.text).paragraph_format.first_line_indent == -6
    p.paragraph_format.first_line_indent = 12.375
    assert p.paragraph_format.first_line_indent == 12.4 and p.paragraph_format.left_indent == 4
    assert 'hanging="' not in document.part_xml("word/document.xml")
    p.paragraph_format.first_line_indent = None
    assert p.paragraph_format.first_line_indent is None and p.paragraph_format.left_indent == 4


@pytest.mark.parametrize("properties", ['<w:ind/><w:ind/>', '<w:spacing/><w:spacing/>'])
def test_duplicate_dimension_groups_fail_before_mutation(properties):
    document = with_direct_properties(properties)
    before = document.to_bytes()
    prop = "left_indent" if "ind" in properties else "space_before"
    with pytest.raises(ValueError, match="Duplicate"):
        setattr(paragraph(document).paragraph_format, prop, 12)
    assert document.to_bytes() == before


def test_character_units_are_preserved_but_not_guessed():
    document = with_direct_properties('<w:ind w:firstLineChars="100"/>')
    before = document.to_bytes()
    with pytest.raises(NotImplementedError, match="font-aware"):
        _ = paragraph(document).effective_paragraph_format
    assert document.to_bytes() == before


RENDER_REPORT = json.loads((BENCHMARKS / "paragraph-dimensions-rendering-26.9.json").read_text())


@pytest.mark.parametrize("row", RENDER_REPORT["records"], ids=lambda row: row["output"])
def test_dimension_style_edit_reaches_native_pdf_geometry(row):
    import runpy

    from aspose.words_foss.pdf_writer import LdmPdfWriter

    probe = runpy.run_path(str(ROOT / "docs/probes/pagination_rendering.py"))
    with ZipFile(BENCHMARKS / RENDER_REPORT["corpus"]) as archive:
        document = DocxDocument(BytesIO(archive.read(row["input"])))
    setattr(document.styles.get_by_name("Target").paragraph_format, row["property"], row["value"])
    before = document.to_bytes()
    raw = LdmPdfWriter().write_to_bytes(document.to_light_document())
    assert document.to_bytes() == before
    actual, expected = probe["pdf_snapshot"](raw), row["native_snapshot"]
    assert actual["page_sizes"] == expected["page_sizes"] and set(actual["lines"]) == set(expected["lines"])
    for label, native in expected["lines"].items():
        result = actual["lines"][label]
        assert result["page"] == native["page"]
        if label != "BEFORE":
            assert result["origin"] == pytest.approx(native["origin"], abs=0.02)
            assert result["advance"] == pytest.approx(native["advance"], abs=0.02)
            assert result["size"] == native["size"]
    with ZipFile(BENCHMARKS / RENDER_REPORT["commercial_outputs"]) as archive:
        reference = archive.read(row["output"])
    assert max(probe["ink_difference"](reference, raw)) <= 0.01
