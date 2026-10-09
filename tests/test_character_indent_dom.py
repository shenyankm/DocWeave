"""Character units can be read without inventing font-dependent point values."""

import json
import runpy
import traceback
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from aspose.words_foss import DocxDocument, NodeType

ROOT = Path(__file__).parents[1]
BENCHMARKS = ROOT / "docs/benchmarks"
PROPERTIES = ("character_unit_left_indent", "character_unit_right_indent", "character_unit_first_line_indent")
READS = json.loads((BENCHMARKS / "paragraph-character-reads-26.9.json").read_text())
NATIVE = [json.loads((BENCHMARKS / name).read_text()) for name in
          ("paragraph-character-indents-26.9.json", "paragraph-character-setters-26.9.json")]
source = runpy.run_path(str(ROOT / "tests/test_character_indent_diagnostics.py"))["source"]
with_direct_properties = runpy.run_path(str(ROOT / "tests/test_paragraph_dimensions.py"))["with_direct_properties"]
INHERITANCE_EDITS = json.loads((BENCHMARKS / "paragraph-character-inheritance-edits-26.9.json").read_text())


@pytest.mark.parametrize("row", INHERITANCE_EDITS["records"], ids=lambda row: row["output"])
def test_native_edited_style_character_values_survive_dom_read_and_save(row):
    with ZipFile(BENCHMARKS / INHERITANCE_EDITS["outputs"]) as corpus:
        document = DocxDocument(BytesIO(corpus.read(row["output"])))
    before = document.to_bytes()
    for target in ("base", "derived"):
        fmt = document.styles.get_by_name(target.capitalize()).paragraph_format
        assert {prop: getattr(fmt, prop) for prop in PROPERTIES} == {
            prop: row["after_reopen"][target][prop] for prop in PROPERTIES}
    assert document.to_bytes() == before


@pytest.mark.parametrize("row", READS["records"], ids=lambda row: row["input"])
def test_character_style_getters_match_native_inheritance_and_keep_direct_layer(row):
    with ZipFile(BENCHMARKS / READS["corpus"]) as corpus:
        document = DocxDocument(BytesIO(corpus.read(row["input"])))
    before = document.to_bytes()
    layers = runpy.run_path(str(ROOT / "docs/probes/paragraph_character_reads.py"))["LAYERS"]
    values = layers[int(Path(row["input"]).stem)]
    prop = row["input"].split("/")[0]
    for target, index in (("base", 1), ("derived", 2)):
        style = document.styles.get_by_name(target.capitalize())
        assert {name: getattr(style.paragraph_format, name) for name in PROPERTIES} == row["values"][target]
        assert getattr(style.direct_paragraph_format, prop) == (values[index] / 100 if values[index] is not None else None)
    paragraph = next(p for p in document.body.paragraphs if "IMPORT" in p.text)
    assert getattr(paragraph.paragraph_format, prop) == (values[3] / 100 if values[3] is not None else None)
    assert document.to_bytes() == before


@pytest.mark.parametrize("report_index,row", [(index, row) for index, report in enumerate(NATIVE) for row in report["records"]],
                         ids=lambda value: str(value) if isinstance(value, int) else value["output"])
def test_native_saved_character_values_can_be_read_without_resolving_points(report_index, row):
    report = NATIVE[report_index]
    with ZipFile(BENCHMARKS / report["outputs"]) as outputs:
        document = DocxDocument(BytesIO(outputs.read(row["output"])))
    before = document.to_bytes()
    paragraph = next(p for p in document.body.paragraphs if "IMPORT" in p.text)
    expected = row["after_reopen"] if report_index == 0 else row["after_reopen"]["paragraph"]
    assert {name: getattr(paragraph.paragraph_format, name) or 0.0 for name in PROPERTIES} == {name: expected[name] for name in PROPERTIES}
    if report_index == 1 and row["target"] == "style":
        style = document.styles.get_by_name("P")
        assert {name: getattr(style.paragraph_format, name) for name in PROPERTIES} == {
            name: row["after_reopen"]["target"][name] for name in PROPERTIES}
    assert document.to_bytes() == before


@pytest.mark.parametrize("attribute", ["leftChars", "startChars", "rightChars", "endChars", "firstLineChars", "hangingChars"])
@pytest.mark.parametrize("kind", ["body", "table", "header", "footer", "style", "default"])
@pytest.mark.parametrize("value", ["0", "113"])
def test_character_getters_preserve_aliases_and_zero_in_supported_stories(attribute, kind, value):
    document = DocxDocument(BytesIO(source(attribute, kind, value)))
    before = document.to_bytes()
    prop = PROPERTIES[2] if attribute in {"firstLineChars", "hangingChars"} else PROPERTIES[0] if attribute in {"leftChars", "startChars"} else PROPERTIES[1]
    if kind in {"style", "default"}:
        fmt = document.styles.get_by_name("Normal").paragraph_format
    else:
        story = document.story(f"word/{kind}1.xml") if kind in {"header", "footer"} else document.body
        paragraphs = story.get_child_nodes(NodeType.PARAGRAPH, deep=True)
        fmt = (paragraphs[-1] if kind == "table" else paragraphs[0]).paragraph_format
    assert getattr(fmt, prop) == int(value) / (-100 if attribute == "hangingChars" else 100)
    assert document.to_bytes() == before


@pytest.mark.parametrize("attrs,prop,expected", [
    ('w:leftChars="100" w:startChars="200"', PROPERTIES[0], 2),
    ('w:startChars="200" w:leftChars="100"', PROPERTIES[0], 1),
    ('w:rightChars="100" w:endChars="200"', PROPERTIES[1], 2),
    ('w:endChars="200" w:rightChars="100"', PROPERTIES[1], 1),
    ('w:firstLineChars="100" w:hangingChars="200"', PROPERTIES[2], -2),
    ('w:hangingChars="200" w:firstLineChars="100"', PROPERTIES[2], 1)])
def test_last_character_alias_wins_without_changing_other_values(attrs, prop, expected):
    document = with_direct_properties(f'<w:ind w:left="240" {attrs}/>')
    paragraph = next(p for p in document.body.paragraphs if "IMPORT" in p.text)
    before = document.to_bytes()
    assert getattr(paragraph.paragraph_format, prop) == expected
    assert document.to_bytes() == before


@pytest.mark.parametrize("prop", PROPERTIES)
@pytest.mark.parametrize("raw", ["PRIVATE ATTRIBUTE", "1.13", "9" * 5000, "1_13", "١١٣", "1e2"])
def test_malformed_character_getter_errors_do_not_disclose_source(prop, raw):
    attribute = dict(zip(PROPERTIES, ("leftChars", "rightChars", "firstLineChars")))[prop]
    document = DocxDocument(BytesIO(source(attribute, "body", raw)))
    before = document.to_bytes()
    with pytest.raises(ValueError, match="integer hundredths") as caught:
        getattr(document.body.paragraphs[0].paragraph_format, prop)
    assert raw not in "".join(traceback.format_exception(caught.value))
    assert document.to_bytes() == before


@pytest.mark.parametrize("prop", PROPERTIES)
def test_character_getters_remain_read_only_until_point_coupling_is_implemented(prop):
    document = DocxDocument(BytesIO(source("leftChars", "body", "0")))
    before = document.to_bytes()
    for fmt in (document.body.paragraphs[0].paragraph_format, document.styles.get_by_name("Normal").paragraph_format):
        with pytest.raises(AttributeError):
            setattr(fmt, prop, 1.13)
    assert document.to_bytes() == before


@pytest.mark.parametrize("xml", ['<w:ind w:leftChars="100"/><w:ind w:leftChars="200"/>',
                                '<w:ind w:leftChars="PRIVATE" w:startChars="100"/>'])
def test_ambiguous_or_malformed_character_groups_fail_without_mutation(xml):
    document = with_direct_properties(xml)
    before = document.to_bytes()
    with pytest.raises(ValueError):
        _ = next(p for p in document.body.paragraphs if "IMPORT" in p.text).paragraph_format.character_unit_left_indent
    assert document.to_bytes() == before
