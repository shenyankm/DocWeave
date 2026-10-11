"""Zero direct/style/default sizes remain distinct from clearing a declaration."""
from io import BytesIO

import pytest

import dockiln as dk
from .test_docx_dom_styles import style_doc
from .test_docx_dom_ranges import paragraph_doc


@pytest.mark.parametrize("scope", ["run", "style", "default"])
@pytest.mark.parametrize("format_name", ["DOCX", "FLAT_OPC"])
def test_zero_size_survives_edit_and_two_saves(tmp_path, scope, format_name):
    default = 0 if scope == "default" else 40
    declared = "" if scope == "default" else '<w:rPr><w:sz w:val="40"/></w:rPr>'
    styles = (f'<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="{default}"/></w:rPr>'
              '</w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Owned">'
              f'<w:name w:val="Owned"/>{declared}</w:style>')
    document = style_doc(tmp_path, styles, '<w:pStyle w:val="Owned"/>')
    content_types = document._package.payload("[Content_Types].xml")
    document._package.set_parts({"[Content_Types].xml": content_types.replace(
        b"</Types>", b'<Default Extension="bin" ContentType="application/octet-stream"/></Types>')})
    run = document.body.paragraphs[0].runs[0]
    if scope == "run":
        run.font.size = 0
    elif scope == "style":
        document.styles.get_by_name("Owned").font.size = 0
    else:
        run.font.size = 1
        run.font.size = None
    assert run.effective_font.size == 0
    for generation in range(2):
        path = tmp_path / f"{generation}.bin"
        (document.save if format_name == "DOCX" else document.save_flat_opc)(path)
        document = dk.DocxDocument(path)
        run = document.body.paragraphs[0].runs[0]
        assert run.effective_font.size == 0
        assert run.font.size == (0 if scope == "run" else None)
        assert document.styles.get_by_name("Owned").font.size == (20 if scope == "run" else 0)
    if scope == "run":
        run.font.size = None
        assert run.font.size is None and run.effective_font.size == 20


def test_zero_range_font_splits_only_the_selection(tmp_path):
    document = paragraph_doc(tmp_path, '<w:r><w:rPr><w:sz w:val="24"/></w:rPr>'
                                      '<w:t>abcdef</w:t></w:r>')
    paragraph = document.body.paragraphs[0]
    paragraph.range(1, 4).apply_font(size=0)
    assert [run.text for run in paragraph.runs] == ["a", "bcd", "ef"]
    assert [run.font.size for run in paragraph.runs] == [12, 0, 12]
    cold = dk.DocxDocument(BytesIO(document.to_bytes())).body.paragraphs[0]
    assert [run.effective_font.size for run in cold.runs] == [12, 0, 12]
