"""Boolean direct declarations survive conversion, model JSON and explicit edits."""

import json
import runpy
import shutil
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from aspose.words_foss import Document, SaveFormat
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.saving import OoxmlSaveOptions

ROOT = Path(__file__).parents[1]
BENCHMARKS = ROOT / 'docs/benchmarks'
REPORT = json.loads((BENCHMARKS / 'font-boolean-contexts-26.9.json').read_text())
ROWS = [r for r in REPORT['records'] if r['field'] != 'hidden']
TOOLS = runpy.run_path(str(ROOT / 'scripts/verify_commercial_baseline.py'))
TAGS = runpy.run_path(str(ROOT / 'docs/probes/font_boolean_contexts.py'))['FIELDS']
FIELDS = [f for f in TAGS if f != 'hidden']


@pytest.mark.parametrize('row', ROWS, ids=[r['input'] for r in ROWS])
@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
@pytest.mark.parametrize('json_roundtrip', [False, True])
def test_boolean_save_keeps_native_getter_and_direct_declaration(row, format, json_roundtrip):
    with ZipFile(BENCHMARKS / REPORT['corpus']) as archive:
        raw = archive.read(row['input'])
    model = Document(BytesIO(raw)).light_document_model
    if json_roundtrip:
        model = ldm.Document.model_validate_json(model.model_dump_json())
    font = model.sections[0].body.paragraphs[0].runs[0].font
    assert getattr(font, row['field'] + '_explicit') is (row['values'][3] is not None)
    saved = LdmDocxWriter(OoxmlSaveOptions(format)).write_to_bytes(model)
    before = TOOLS['saved_hidden_state'](raw, TAGS[row['field']])
    after = TOOLS['saved_hidden_state'](saved, TAGS[row['field']])
    assert after[0] is row['value'] and after[2:] == before[2:]
    reopened = Document(BytesIO(saved)).light_document_model.sections[0].body.paragraphs[0].runs[0].font
    assert getattr(reopened, row['field']) is row['value']


@pytest.mark.parametrize('field', FIELDS)
@pytest.mark.parametrize('copy', [False, True])
@pytest.mark.parametrize('value', [False, True])
def test_explicit_boolean_edit_is_preserved_even_when_value_is_unchanged(field, copy, value):
    row = next(r for r in ROWS if r['field'] == field and r['values'] == [None, value, None, None])
    with ZipFile(BENCHMARKS / REPORT['corpus']) as archive:
        model = Document(BytesIO(archive.read(row['input']))).light_document_model
    run = model.sections[0].body.paragraphs[0].runs[0]
    if copy:
        run.font = run.font.model_copy(update={field: value})
    else:
        setattr(run.font, field, value)
    model = ldm.Document.model_validate_json(model.model_dump_json())
    font = model.sections[0].body.paragraphs[0].runs[0].font
    assert getattr(font, field + '_explicit') is True
    raw = LdmDocxWriter().write_to_bytes(model)
    state = TOOLS['saved_hidden_state'](raw, TAGS[field])
    assert state[0] is state[2] is value


@pytest.mark.parametrize('field', ldm.FONT_BOOLEAN_FIELDS)
def test_model_copy_honors_explicit_origin_override(field):
    font = ldm.Font().model_copy(update={field: True, field + '_explicit': False})
    assert getattr(font, field) is True and getattr(font, field + '_explicit') is False


def test_boolean_roundtrip_verifier_rejects_native_getter_forgery(tmp_path):
    report = json.loads((BENCHMARKS / 'font-boolean-roundtrip-26.9.json').read_text())
    (tmp_path / 'corpus').mkdir()
    for name in ('font-boolean-contexts-26.9.json', REPORT['corpus'],
                 report['outputs'], report['before_origin_fix']['outputs']):
        shutil.copyfile(BENCHMARKS / name, tmp_path / name)
    row = report['native_reread']['records'][0]
    row['value'] = not row['value']
    (tmp_path / 'font-boolean-roundtrip-26.9.json').write_text(json.dumps(report))
    with pytest.raises(AssertionError):
        TOOLS['verify_font_boolean_roundtrip'](tmp_path)


def test_byte_identical_outputs_share_archive_storage():
    report = json.loads((BENCHMARKS / 'font-boolean-roundtrip-26.9.json').read_text())
    for phase in (report, report['before_origin_fix']):
        with ZipFile(BENCHMARKS / phase['outputs']) as archive:
            assert len(archive.namelist()) < len(phase['records'])
            assert set(archive.namelist()) == {r['storage'] for r in phase['records']}
    assert TOOLS['verify_font_boolean_roundtrip'](BENCHMARKS) == 3240


def test_boolean_roundtrip_verifier_rejects_wrong_storage_alias(tmp_path):
    report = json.loads((BENCHMARKS / 'font-boolean-roundtrip-26.9.json').read_text())
    (tmp_path / 'corpus').mkdir()
    for name in ('font-boolean-contexts-26.9.json', REPORT['corpus'],
                 report['outputs'], report['before_origin_fix']['outputs']):
        shutil.copyfile(BENCHMARKS / name, tmp_path / name)
    row = report['records'][0]
    row['storage'] = next(r['storage'] for r in report['records'] if r['output_sha256'] != row['output_sha256'])
    (tmp_path / 'font-boolean-roundtrip-26.9.json').write_text(json.dumps(report))
    with pytest.raises(AssertionError):
        TOOLS['verify_font_boolean_roundtrip'](tmp_path)
