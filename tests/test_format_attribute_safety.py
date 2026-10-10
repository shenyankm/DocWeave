"""Unsupported formatting writes fail instead of disappearing at save time."""

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from aspose.words_foss import DocxDocument

from .test_docx_dom import payloads


def document():
    root = Path(__file__).parents[1] / 'docs/benchmarks'
    with ZipFile(root / 'corpus/font-boolean-contexts-26.9.zip') as archive:
        return DocxDocument(BytesIO(archive.read('all_caps-n-n-n-n.docx')))


def formats(doc):
    paragraph = doc.body.paragraphs[0]
    style = doc.styles.get_by_id('P')
    return (paragraph.runs[0].font, paragraph.paragraph_format, style.font, style.direct_font,
            style.paragraph_format, style.direct_paragraph_format)


@pytest.mark.parametrize('index', range(6))
@pytest.mark.parametrize('name', ['color', 'nonexistent_format', 'order', 'property_name', '_set'])
def test_unknown_properties_and_method_or_constant_overrides_fail_before_mutation(index, name):
    doc = document()
    original = payloads(doc.to_bytes())
    format = formats(doc)[index]
    with pytest.raises(AttributeError):
        setattr(format, name, 'FF0000')
    assert payloads(doc.to_bytes()) == original
    assert not hasattr(format, '__dict__')


def test_supported_properties_still_use_setters_and_saved_inheritance():
    doc = document()
    run_font, paragraph, style_font, direct_font, style_paragraph, direct_paragraph = formats(doc)
    run_font.hidden = True
    run_font.size = 12.5
    paragraph.keep_with_next = True
    style_font.all_caps = True
    direct_font.italic = True
    style_paragraph.alignment = 'center'
    direct_paragraph.space_before = 7.5
    for data in (doc.to_bytes(), doc.to_flat_opc()):
        reopened = DocxDocument(BytesIO(data))
        p = reopened.body.paragraphs[0]
        assert p.runs[0].font.hidden is True and p.runs[0].font.size == 12.5
        assert p.runs[0].effective_font.all_caps is True
        assert p.runs[0].effective_font.italic is True
        assert p.paragraph_format.keep_with_next is True
        assert p.effective_paragraph_format.alignment == 'center'
        assert p.effective_paragraph_format.space_before == 7.5
