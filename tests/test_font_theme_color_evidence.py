"""Theme diagnostics freeze getter and PDF color separately, without acceptance."""

import hashlib
import json
import runpy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pymupdf


def test_owned_theme_color_observations_are_reproducible():
    root = Path(__file__).parents[1]
    base = root / 'docs/benchmarks'
    report = json.loads((base / 'font-theme-colors-26.9.json').read_text())
    generated = {n: (case, raw) for n, case, raw in runpy.run_path(
        str(root / 'docs/probes/font_theme_colors.py'))['inputs']()}
    assert len(generated) == 32
    assert report['implementation_acceptance'] is report['rendering_acceptance'] is False
    for key in ('corpus', 'native_outputs', 'current_before_outputs'):
        assert hashlib.sha256((base / report[key]).read_bytes()).hexdigest() == report[key + '_sha256']
    with ZipFile(base / report['corpus']) as archive:
        assert set(archive.namelist()) == set(generated)
        for name, (_, rebuilt) in generated.items():
            with ZipFile(BytesIO(rebuilt)) as expected, ZipFile(BytesIO(archive.read(name))) as actual:
                assert {n: expected.read(n) for n in expected.namelist()} == {
                    n: actual.read(n) for n in actual.namelist()}
    for observation, output_key in (('native', 'native_outputs'), ('current_before_theme_fix', 'current_before_outputs')):
        rows = report[observation]['records']
        assert len(rows) == 32 and {r['input'] for r in rows} == set(generated)
        with ZipFile(base / report['corpus']) as sources, ZipFile(base / report[output_key]) as outputs:
            for row in rows:
                assert hashlib.sha256(sources.read(row['input'])).hexdigest() == row['sha256']
                assert all(row[k] == v for k, v in generated[row['input']][0].items())
                raw = outputs.read(row['output'])
                assert hashlib.sha256(raw).hexdigest() == row['output_sha256']
                with pymupdf.open(stream=raw, filetype='pdf') as pdf:
                    colors = [s['color'] for p in pdf for b in p.get_text('dict')['blocks']
                              for line in b.get('lines', []) for s in line['spans'] if s['text'] == 'IMPORT']
                assert colors == [row['pdf_target_color']]
    native = report['native']['records']
    before = report['current_before_theme_fix']['records']
    assert sum(a['getter'] != b['getter'] for a, b in zip(native, before)) == report['getter_differences_before'] == 16
    assert sum(a['pdf_target_color'] != b['pdf_target_color'] for a, b in zip(native, before)) == report['pdf_target_color_differences_before'] == 20
