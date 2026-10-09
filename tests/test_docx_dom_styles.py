"""Supported effective properties follow defaults, basedOn chains and OOXML toggles."""

import hashlib
import json
from dataclasses import FrozenInstanceError
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss.dom.nodes import W

from .test_docx_dom import package, payloads

REL = "http://schemas.openxmlformats.org/package/2006/relationships"
STYLES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
BENCHMARKS = Path(__file__).parents[1] / "docs" / "benchmarks"
TOGGLE_BASELINE = json.loads((BENCHMARKS / "style-toggles-26.9.json").read_text())


@pytest.mark.parametrize("record", TOGGLE_BASELINE["records"], ids=lambda record: record["input"])
def test_style_toggle_getters_match_recorded_commercial_inputs(record):
    with ZipFile(BENCHMARKS / "corpus" / "style-toggles-26.9.zip") as archive:
        data = archive.read(record["input"])
    assert hashlib.sha256(data).hexdigest() == record["sha256"]
    doc = aw.DocxDocument(BytesIO(data))
    run = doc.body.paragraphs[0].runs[0]
    font = run.effective_font
    assert (font.bold, font.italic) == (record["bold"], record["italic"])
    assert payloads(doc.to_bytes()) == payloads(data)


def style_doc(tmp_path, styles="", ppr="", rpr="", body=None, target="styles.xml", extras=None):
    path = tmp_path / "styles.docx"
    if body is None:
        body = f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t>sample</w:t></w:r></w:p>'
    parts = {
        "word/styles.xml": f'<w:styles xmlns:w="{W}">{styles}</w:styles>'.encode(),
        "word/_rels/document.xml.rels": f'<Relationships xmlns="{REL}"><Relationship Id="s" Type="{STYLES_REL}" Target="{target}"/></Relationships>'.encode(),
        **(extras or {}),
    }
    package(path, body, parts)
    return aw.DocxDocument(path)


def test_nearest_style_values_combine_between_paragraph_and_character_levels(tmp_path):
    styles = (
        '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:styleId="Base" w:default="1"><w:rPr><w:b/><w:i/></w:rPr><w:pPr><w:jc w:val="center"/></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/><w:rPr><w:b/><w:i w:val="0"/><w:sz w:val="26"/></w:rPr></w:style>'
        '<w:style w:type="character" w:styleId="CharBase"><w:rPr><w:b/></w:rPr></w:style>'
        '<w:style w:type="character" w:styleId="CharChild"><w:basedOn w:val="CharBase"/><w:rPr><w:i/><w:sz w:val="29"/></w:rPr></w:style>'
    )
    doc = style_doc(tmp_path, styles, '<w:pStyle w:val="Child"/>', '<w:rStyle w:val="CharChild"/>')
    p = doc.body.paragraphs[0]
    run = p.runs[0]
    original = payloads(doc.to_bytes())
    assert run.font.bold is None and run.font.size is None
    effective = run.effective_font
    assert effective.bold is False and effective.italic is True and effective.size == 14.5
    assert p.effective_paragraph_format.alignment == "center"
    assert p.paragraph_format.alignment is None
    assert payloads(doc.to_bytes()) == original  # Read-only resolution never materializes direct formatting.
    with pytest.raises(FrozenInstanceError):
        effective.bold = False
    run.font.bold = False
    run.font.italic = True
    run.font.size = 10.5
    p.paragraph_format.alignment = "both"
    assert run.effective_font.bold is False and run.effective_font.italic is True
    assert run.effective_font.size == 10.5 and p.effective_paragraph_format.alignment == "both"
    run.font.bold = run.font.italic = run.font.size = None
    p.paragraph_format.alignment = None
    assert run.effective_font == effective
    assert p.effective_paragraph_format.alignment == "center"


def test_default_styles_docdefaults_and_paragraph_mark_not_applied_to_run(tmp_path):
    styles = (
        '<w:docDefaults><w:rPrDefault><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:styleId="DefaultP" w:default="true"><w:rPr><w:b/><w:i/></w:rPr></w:style>'
        '<w:style w:type="character" w:styleId="DefaultC" w:default="on"><w:rPr><w:i/><w:sz w:val="70"/></w:rPr></w:style>'
    )
    doc = style_doc(tmp_path, styles, '<w:rPr><w:b/><w:i/><w:sz w:val="90"/></w:rPr>')
    run = doc.body.paragraphs[0].runs[0]
    assert run.effective_font.bold is True and run.effective_font.italic is True
    assert run.effective_font.size == 12
    assert doc.body.paragraphs[0].effective_paragraph_format.alignment == "left"


def test_character_style_setter_validates_type_and_effective_values_are_fresh(tmp_path):
    styles = ('<w:style w:type="paragraph" w:styleId="P"><w:rPr><w:b/></w:rPr></w:style>'
              '<w:style w:type="character" w:styleId="C"><w:rPr><w:i/></w:rPr></w:style>')
    doc = style_doc(tmp_path, styles)
    p = doc.body.paragraphs[0]
    run = p.runs[0]
    assert not run.effective_font.bold and not run.effective_font.italic
    run.font.style_id = "C"
    p.paragraph_format.style_id = "P"
    assert run.font.style_id == "C" and run.effective_font.bold and run.effective_font.italic
    before = payloads(doc.to_bytes())
    with pytest.raises(ValueError):
        run.font.style_id = "P"
    with pytest.raises(ValueError):
        p.paragraph_format.style_id = "C"
    with pytest.raises(ValueError):
        run.font.style_id = "missing"
    assert payloads(doc.to_bytes()) == before
    run.font.style_id = None
    assert run.font.style_id is None and not run.effective_font.italic
    reopened = aw.DocxDocument(BytesIO(doc.to_bytes()))
    assert reopened.body.paragraphs[0].runs[0].effective_font.bold


@pytest.mark.parametrize("target,name", [("../formatting/styles%20custom.xml", "formatting/styles custom.xml"), ("/formatting/styles.xml", "formatting/styles.xml")])
def test_styles_relationship_not_hardcoded_part_path(tmp_path, target, name):
    actual = f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Actual" w:default="1"><w:rPr><w:b/></w:rPr></w:style></w:styles>'.encode()
    doc = style_doc(tmp_path, target=target, extras={name: actual})
    run = doc.body.paragraphs[0].runs[0]
    before = payloads(doc.to_bytes())
    assert run.effective_font.bold is True
    doc.body.paragraphs[0].paragraph_format.style_id = "Actual"
    assert payloads(doc.to_bytes())[name] == before[name]
    assert payloads(doc.to_bytes())["word/styles.xml"] == before["word/styles.xml"]


@pytest.mark.parametrize("styles,reference", [
    ('<w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="B"/></w:style><w:style w:type="paragraph" w:styleId="B"><w:basedOn w:val="A"/></w:style>', "A"),
    ('<w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="missing"/></w:style>', "A"),
    ('<w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="C"/></w:style><w:style w:type="character" w:styleId="C"/>', "A"),
    ('<w:style w:type="paragraph" w:styleId="A"/><w:style w:type="paragraph" w:styleId="A"/>', "A"),
    ('', "missing"),
])
def test_broken_style_graph_is_not_silently_flattened(tmp_path, styles, reference):
    doc = style_doc(tmp_path, styles, f'<w:pStyle w:val="{reference}"/>')
    original = payloads(doc.to_bytes())
    with pytest.raises(ValueError):
        _ = doc.body.paragraphs[0].runs[0].effective_font
    with pytest.raises(ValueError):
        _ = doc.body.paragraphs[0].effective_paragraph_format
    assert payloads(doc.to_bytes()) == original


@pytest.mark.parametrize("styles", [
    '<w:style w:styleId="A" w:default="1"/><w:style w:styleId="B" w:default="true"/>',
    '<w:style w:styleId="A" w:default="invalid"/>',
])
def test_invalid_default_styles_fail_without_mutation(tmp_path, styles):
    doc = style_doc(tmp_path, styles)
    original = payloads(doc.to_bytes())
    with pytest.raises(ValueError):
        _ = doc.body.paragraphs[0].runs[0].effective_font
    assert payloads(doc.to_bytes()) == original


@pytest.mark.parametrize("rpr", ['<w:b w:val="invalid"/>', '<w:sz w:val="invalid"/>', '<w:sz w:val="0"/>'])
def test_invalid_direct_values_are_not_reported_as_valid_effective_values(tmp_path, rpr):
    doc = style_doc(tmp_path, rpr=rpr)
    with pytest.raises(ValueError):
        _ = doc.body.paragraphs[0].runs[0].effective_font


@pytest.mark.parametrize("style,ppr,body", [
    ('', '<w:numPr><w:numId w:val="1"/></w:numPr>', None),
    ('<w:style w:type="paragraph" w:styleId="P" w:default="1"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>', '', None),
    ('', '', '<w:tbl><w:tblPr><w:tblStyle w:val="T"/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'),
    ('<w:style w:type="table" w:styleId="T" w:default="1"/>', '', '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'),
])
def test_unsupported_context_is_explicit_not_a_guessed_effective_value(tmp_path, style, ppr, body):
    doc = style_doc(tmp_path, style, ppr, body=body)
    p = doc.get_child_nodes(aw.NodeType.PARAGRAPH, deep=True)[0]
    original = payloads(doc.to_bytes())
    with pytest.raises(NotImplementedError):
        _ = p.runs[0].effective_font
    with pytest.raises(NotImplementedError):
        _ = p.effective_paragraph_format
    assert payloads(doc.to_bytes()) == original
    p.runs[0].font.bold = True  # Direct-format edits do not require guessing inherited values.
    assert p.runs[0].font.bold is True


def test_unstyled_plain_table_and_header_use_document_styles(tmp_path):
    styles = '<w:style w:type="paragraph" w:styleId="P" w:default="1"><w:rPr><w:i/></w:rPr></w:style>'
    body = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    doc = style_doc(tmp_path, styles, body=body)
    assert doc.body.tables[0].rows[0].cells[0].paragraphs[0].runs[0].effective_font.italic
    assert doc.story("word/header1.xml").paragraphs[0].runs[0].effective_font.italic


@pytest.mark.parametrize("target", ["https://example.com/styles.xml", "../../../outside.xml", "missing.xml", "styles.xml#fragment", "styles.xml?query"])
def test_invalid_styles_relationship_target_is_rejected_before_style_assignment(tmp_path, target):
    doc = style_doc(tmp_path, '<w:style w:styleId="P"/>', target=target)
    original = payloads(doc.to_bytes())
    with pytest.raises(ValueError):
        doc.body.paragraphs[0].paragraph_format.style_id = "P"
    assert payloads(doc.to_bytes()) == original


def test_absent_relationship_does_not_apply_unrelated_styles_part(tmp_path):
    doc = style_doc(tmp_path, '<w:style w:styleId="P" w:default="1"><w:rPr><w:b/></w:rPr></w:style>', extras={
        "word/_rels/document.xml.rels": f'<Relationships xmlns="{REL}"/>'.encode(),
    })
    run = doc.body.paragraphs[0].runs[0]
    assert run.effective_font.bold is False and run.effective_font.size == 11.0
    with pytest.raises(ValueError):
        doc.body.paragraphs[0].paragraph_format.style_id = "P"


@pytest.mark.parametrize("extras", [
    {"word/_rels/document.xml.rels": f'<Relationships xmlns="{REL}"><Relationship Id="s" Type="{STYLES_REL}" Target="styles.xml" TargetMode="External"/></Relationships>'.encode()},
    {"word/_rels/document.xml.rels": f'<Relationships xmlns="{REL}"><Relationship Id="a" Type="{STYLES_REL}" Target="styles.xml"/><Relationship Id="b" Type="{STYLES_REL}" Target="styles.xml"/></Relationships>'.encode()},
    {"word/_rels/document.xml.rels": b'<Relationships xmlns="urn:wrong"/>'},
    {"word/styles.xml": b'<styles xmlns="urn:wrong"/>'},
])
def test_malformed_or_ambiguous_styles_relationship_is_not_used(tmp_path, extras):
    doc = style_doc(tmp_path, '<w:style w:styleId="P"/>', extras=extras)
    original = payloads(doc.to_bytes())
    with pytest.raises(ValueError):
        doc.body.paragraphs[0].paragraph_format.style_id = "P"
    assert payloads(doc.to_bytes()) == original
