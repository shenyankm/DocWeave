"""Replay official live Boolean edits and retain direct declarations through saving."""

import hashlib
import json
import runpy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from aspose.words_foss import DocxDocument, NodeType
from aspose.words_foss.light_document_model import FONT_BOOLEAN_FIELDS

from .test_docx_dom import package, payloads
from .test_docx_dom_ranges import paragraph_doc

BENCHMARKS = Path(__file__).parents[1] / 'docs/benchmarks'
with ZipFile(BENCHMARKS / 'corpus/dom-font-boolean-observations-26.9.zip') as archive:
    NATIVE = json.loads(archive.read('live.json'))
with ZipFile(BENCHMARKS / 'corpus/font-boolean-contexts-26.9.zip') as archive:
    INPUTS = {name: archive.read(name) for name in archive.namelist()}
ERRORS = json.loads((BENCHMARKS / 'dom-font-boolean-errors-26.9.json').read_text())['native']


def snapshot(document, field):
    fonts = (document.body.paragraphs[0].runs[0].effective_font,
             document.styles.get_by_id('P').font, document.styles.get_by_id('C').font)
    return {name: getattr(font, field) for name, font in zip(('run', 'paragraph', 'character'), fonts)}


@pytest.mark.parametrize('row', NATIVE['records'], ids=lambda r: f"{r['input']}-{r['target']}-{r['value']}")
def test_live_boolean_edit_matches_native_and_survives_both_formats(row):
    raw = INPUTS[row['input']]
    assert hashlib.sha256(raw).hexdigest() == row['source_sha256']
    document = DocxDocument(BytesIO(raw))
    assert snapshot(document, row['field']) == row['before']
    run = document.body.paragraphs[0].runs[0]
    font = run.font if row['target'] == 'run' else document.styles.get_by_id(
        'P' if row['target'] == 'paragraph' else 'C').font
    setattr(font, row['field'], row['value'])
    assert snapshot(document, row['field']) == row['after']
    for saved in (document.to_bytes(), document.to_flat_opc()):
        reopened = DocxDocument(BytesIO(saved))
        assert snapshot(reopened, row['field']) == row['after']
        direct = reopened.body.paragraphs[0].runs[0].font if row['target'] == 'run' else reopened.styles.get_by_id(
            'P' if row['target'] == 'paragraph' else 'C').direct_font
        assert getattr(direct, row['field']) is row['value']


@pytest.mark.parametrize('field', FONT_BOOLEAN_FIELDS)
def test_range_boolean_boundaries_clear_and_invalid_value_are_atomic(tmp_path, field):
    document = paragraph_doc(tmp_path, '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>abcdef</w:t></w:r>')
    paragraph = document.body.paragraphs[0]
    before = payloads(document.to_bytes())
    with pytest.raises(TypeError):
        paragraph.range(1, 5).apply_font(**{field: 1})
    assert payloads(document.to_bytes()) == before
    selected = paragraph.range(1, 5)
    selected.apply_font(**{field: True})
    assert [r.text for r in paragraph.runs] == ['a', 'bcde', 'f']
    assert [getattr(r.font, field) for r in paragraph.runs] == [None, True, None]
    assert all('FF0000' in r.xml for r in paragraph.runs)
    selected.apply_font(**{field: False})
    assert getattr(paragraph.runs[1].font, field) is False
    selected.apply_font(**{field: None})
    assert getattr(paragraph.runs[1].font, field) is None
    assert getattr(paragraph.runs[1].effective_font, field) is False


@pytest.mark.parametrize('field', FONT_BOOLEAN_FIELDS)
def test_style_clearing_and_invalid_assignments_preserve_package(field):
    raw = next(value for name, value in INPUTS.items() if name.startswith('all_caps-'))
    document = DocxDocument(BytesIO(raw))
    style = document.styles.get_by_id('P')
    setattr(style.font, field, True)
    assert getattr(style.direct_font, field) is True
    before = payloads(document.to_bytes())
    for font, value, error in ((style.font, None, TypeError), (style.direct_font, 1, TypeError),
                               (document.body.paragraphs[0].runs[0].font, 'false', TypeError)):
        with pytest.raises(error):
            setattr(font, field, value)
        assert payloads(document.to_bytes()) == before
    setattr(style.direct_font, field, None)
    assert getattr(style.direct_font, field) is None


@pytest.mark.parametrize('field', FONT_BOOLEAN_FIELDS)
def test_boolean_edit_under_content_control_rejects_before_mutation(tmp_path, field):
    source = tmp_path / 'protected.docx'
    package(source, '<w:sdt><w:sdtContent><w:p><w:r><w:t>protected</w:t></w:r></w:p></w:sdtContent></w:sdt>')
    document = DocxDocument(source)
    run = document.get_child_nodes(NodeType.RUN, deep=True)[0]
    before = payloads(document.to_bytes())
    with pytest.raises(NotImplementedError):
        setattr(run.font, field, True)
    assert payloads(document.to_bytes()) == before


def test_frozen_native_live_and_cold_evidence_verifies():
    verify = runpy.run_path(str(BENCHMARKS.parents[1] / 'scripts/verify_commercial_baseline.py'))
    assert verify['verify_dom_font_booleans'](BENCHMARKS) == 10692
    assert verify['verify_dom_font_boolean_errors'](BENCHMARKS) == 297


@pytest.mark.parametrize('row', ERRORS)
def test_native_boolean_type_errors_are_atomic_and_none_clearing_is_explicit(row):
    document = DocxDocument(BytesIO(INPUTS[f"{row['field']}-1-1-n-n.docx"]))
    run = document.body.paragraphs[0].runs[0]
    font = run.font if row['target'] == 'run' else document.styles.get_by_id(
        'P' if row['target'] == 'paragraph' else 'C').font
    before = payloads(document.to_bytes())
    assert getattr(run.effective_font, row['field']) is row['before']
    assert row['error'] == 'TypeError'
    if row['target'] == 'run' and row['input'] is None:
        setattr(font, row['field'], None)
        assert getattr(font, row['field']) is None
    else:
        with pytest.raises(TypeError):
            setattr(font, row['field'], row['input'])
    assert getattr(run.effective_font, row['field']) is row['after']
    assert payloads(document.to_bytes()) == before
