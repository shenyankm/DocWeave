"""Replay frozen theme getters, declarations, resource bytes and target colors."""

import hashlib
import json
import math
import re
import runpy
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pymupdf

from aspose.words_foss import Document
from aspose.words_foss._flat_opc import decode
from aspose.words_foss._opc import (
    related_part_snapshot,
    relationships_path,
    resolve_target,
)


def verify_theme_colors(base):
    root = Path(__file__).parents[1]
    report = json.loads((base / 'font-theme-colors-26.9.json').read_text())
    generated = {n: (case, raw) for n, case, raw in runpy.run_path(
        str(root / 'docs/probes/font_theme_colors.py'))['inputs']()}
    assert len(generated) == 32
    assert report['implementation_acceptance'] is report['rendering_acceptance'] is False
    for key in ('corpus', 'native_outputs', 'current_before_outputs', 'current_after_outputs'):
        assert hashlib.sha256((base / report[key]).read_bytes()).hexdigest() == report[key + '_sha256']
    with ZipFile(base / report['corpus']) as archive:
        assert set(archive.namelist()) == set(generated)
        for name, (_, rebuilt) in generated.items():
            with ZipFile(BytesIO(rebuilt)) as expected, ZipFile(BytesIO(archive.read(name))) as actual:
                assert {n: expected.read(n) for n in expected.namelist()} == {
                    n: actual.read(n) for n in actual.namelist()}
    for observation, output_key in (('native', 'native_outputs'), ('current_before_theme_fix', 'current_before_outputs'),
                                    ('current_after_loading_fix', 'current_after_outputs')):
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

    after = report['current_after_loading_fix']['records']
    assert sum(a['getter'] != b['getter'] for a, b in zip(native, after)) == report['getter_differences_after_loading_fix'] == 0
    assert sum(a['pdf_target_color'] != b['pdf_target_color'] for a, b in zip(native, after)) == report['pdf_target_color_differences_after_loading_fix'] == 0

    measure = runpy.run_path(str(Path(__file__).parents[1] / 'docs/probes/font_theme_colors.py'))['pdf_observation']
    for key, output_key in (('native', 'native_outputs'), ('current_before_theme_fix', 'current_before_outputs'),
                            ('current_after_loading_fix', 'current_after_outputs')):
        with ZipFile(base / report[output_key]) as archive:
            for row in report[key]['records']:
                assert all(row[k] == value for k, value in measure(archive.read(row['output'])).items())
    assert report['pixel_channel_tolerance'] == 1
    assert sum(max(abs(x - y) for x, y in zip(a['painted_target_rgb'], b['painted_target_rgb'])) > 1
               for a, b in zip(native, after)) == report['painted_color_differences_after_fix'] == 0

    return 32


def verify_theme_contexts(base):
    root = Path(__file__).parents[1]
    report = json.loads((base / 'font-theme-combined-contexts-26.9.json').read_text())
    generated = {name: (case, raw) for name, case, raw in runpy.run_path(
        str(root / 'docs/probes/font_theme_combined_contexts.py'))['inputs']()}
    assert len(generated) == 16
    assert report['full_Font_acceptance'] is report['rendering_acceptance'] is False
    for key in ('corpus', 'native_outputs', 'current_outputs'):
        assert hashlib.sha256((base / report[key]).read_bytes()).hexdigest() == report[key + '_sha256']
    measure = runpy.run_path(str(root / 'docs/probes/font_theme_colors.py'))['pdf_observation']
    with ZipFile(base / report['corpus']) as sources:
        assert set(sources.namelist()) == set(generated)
        for name, (_, rebuilt) in generated.items():
            with ZipFile(BytesIO(rebuilt)) as expected, ZipFile(BytesIO(sources.read(name))) as actual:
                assert set(expected.namelist()) == set(actual.namelist())
                for part in expected.namelist():
                    a, b = actual.read(part), expected.read(part)
                    assert ET.tostring(ET.fromstring(a)) == ET.tostring(ET.fromstring(b))
        for key, output_key in (('native', 'native_outputs'), ('current', 'current_outputs')):
            rows = report[key]['records']
            assert len(rows) == 16 and {r['input'] for r in rows} == set(generated)
            with ZipFile(base / report[output_key]) as outputs:
                assert set(outputs.namelist()) == {r['output'] for r in rows}
                for row in rows:
                    assert hashlib.sha256(sources.read(row['input'])).hexdigest() == row['sha256']
                    assert all(row[k] == v for k, v in generated[row['input']][0].items())
                    raw = outputs.read(row['output'])
                    assert hashlib.sha256(raw).hexdigest() == row['output_sha256']
                    assert all(row[k] == v for k, v in measure(raw).items())
    assert report['pixel_channel_tolerance'] == 1
    for native, current in zip(report['native']['records'], report['current']['records']):
        assert native['input'] == current['input'] and native['getter'] == current['getter']
        expected = 0 if native['fallback'] == 'auto' else 0xFF0000
        assert native['pdf_target_color'] == current['pdf_target_color'] == expected
        assert max(abs(a - b) for a, b in zip(native['painted_target_rgb'], current['painted_target_rgb'])) <= 1

    return 16


def verify_theme_roundtrips(base):
    report = json.loads((base / 'font-theme-roundtrip-26.9.json').read_text())
    original = json.loads((base / 'font-theme-colors-26.9.json').read_text())
    expected = {r['input']: r for r in original['native']['records']}
    for key in ('corpus', 'outputs'):
        assert hashlib.sha256((base / report[key]).read_bytes()).hexdigest() == report[key + '_sha256']
    assert report['version'] == '26.9.0' and report['licensed'] is False
    assert report['full_Font_acceptance'] is report['rendering_acceptance'] is False
    rows = report['records']
    assert len(rows) == 128
    assert {(r['input'], r['json_roundtrip'], r['format']) for r in rows} == {
        (name, phase, fmt) for name in expected for phase in (False, True) for fmt in ('docx', 'flat_opc')}
    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    relns = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    with ZipFile(base / report['outputs']) as outputs, ZipFile(base / report['corpus']) as corpus:
        assert set(outputs.namelist()) == {r['output'] for r in rows}
        for row in rows:
            raw = outputs.read(row['output'])
            assert hashlib.sha256(raw).hexdigest() == row['output_sha256']
            assert raw.startswith(b'PK') is (row['format'] == 'docx')
            source = expected[row['input']]
            assert row['getter'] == source['getter'] and row['theme_color'] == 4
            tint = re.search('themeTint="([0-9A-F]+)"', source['modifiers'])
            shade = re.search('themeShade="([0-9A-F]+)"', source['modifiers'])
            value = (1 - int(tint[1], 16) / 255) if tint else -(1 - int(shade[1], 16) / 255) if shade else 0
            assert math.isclose(row['tint_and_shade'], value, abs_tol=1e-15)
            with ZipFile(BytesIO(corpus.read(row['input']))) as original, ZipFile(BytesIO(
                    raw if row['format'] == 'docx' else decode(raw))) as saved:
                def color(package):
                    document = ET.fromstring(package.read('word/document.xml'))
                    run = next(r for r in document.iter(w + 'r') if ''.join(t.text or '' for t in r.iter(w + 't')) == 'IMPORT')
                    return run.find(w + 'rPr/' + w + 'color').attrib
                assert color(original) == color(saved)
                relationships = ET.fromstring(saved.read('word/_rels/document.xml.rels'))
                themes = [r for r in relationships.findall(relns + 'Relationship') if r.get('Type').endswith('/theme')]
                assert len(themes) == 1
                target = resolve_target('word/document.xml', themes[0].get('Target'))
                actual = saved.read(target)
                expected_theme = original.read('word/theme/theme1.xml')
                assert ET.tostring(ET.fromstring(actual)) == ET.tostring(ET.fromstring(expected_theme))
                if row['format'] == 'docx':
                    assert actual == expected_theme
                types = ET.fromstring(saved.read('[Content_Types].xml'))
                assert any(t.get('PartName') == '/' + target and t.get('ContentType') ==
                           'application/vnd.openxmlformats-officedocument.theme+xml' for t in types)

    return 128


def verify_theme_resources(base):
    import hashlib
    import json

    from aspose.words_foss._flat_opc import decode
    from aspose.words_foss.docx_writer import LdmDocxWriter

    report = json.loads((base / 'font-theme-resources-26.9.json').read_text())
    assert report['version'] == '26.9.0' and report['license'] == 'official trial, no license'
    assert report['implementation_acceptance'] is report['rendering_acceptance'] is False
    corpus = base / report['corpus']
    assert hashlib.sha256(corpus.read_bytes()).hexdigest() == report['corpus_sha256']
    assert len(report['records']) == 6
    assert {row['file'] for row in report['records']} == {
        'input.docx', 'current.docx', 'current.xml', 'native-input.docx.docx',
        'native-current.docx.docx', 'native-current.xml.docx'}
    with ZipFile(corpus) as archive:
        generated = runpy.run_path(str(Path(__file__).parents[1] / 'docs/probes/font_theme_resources.py'))['generate_input']()
        with ZipFile(BytesIO(generated)) as expected, ZipFile(BytesIO(archive.read('input.docx'))) as actual:
            assert set(expected.namelist()) == set(actual.namelist())
            for name in expected.namelist():
                a, b = expected.read(name), actual.read(name)
                if name.endswith(('.xml', '.rels')):
                    assert ET.tostring(ET.fromstring(a)) == ET.tostring(ET.fromstring(b))
                else:
                    assert a == b
        for row in report['records']:
            raw = archive.read(row['file'])
            assert hashlib.sha256(raw).hexdigest() == row['sha256']
            package = decode(raw) if row['file'] == 'current.xml' else raw
            with ZipFile(BytesIO(package)) as inner:
                graph = related_part_snapshot(inner, row['theme_part'])
                theme = ET.fromstring(inner.read(row['theme_part']))
                blips = list(theme.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip'))
                assert len(blips) == 1
                rid = blips[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                links = ET.fromstring(inner.read(relationships_path(row['theme_part'])))
                link = next(item for item in links if item.get('Id') == rid)
                assert link.get('Type').endswith('/image') and link.get('TargetMode') != 'External'
                image = resolve_target(row['theme_part'], link.get('Target'))
                assert image == row['theme_images'][0]['part']
            images = [{'part': name, 'sha256': hashlib.sha256(data).hexdigest()}
                      for name, kind, data in graph if kind == 'image/png']
            assert images == row['theme_images']
            assert images[0]['sha256'] == 'de33ddc09a0ba9b83128b6b59461e3be59299575eb68ccf2c03984b504f9fa9c'
        model = Document(BytesIO(archive.read('input.docx'))).light_document_model
        saved = LdmDocxWriter().write_to_bytes(model)
        current = Document(BytesIO(saved)).light_document_model.source_theme
        assert any(part.content_type == 'image/png' and hashlib.sha256(part.data).hexdigest() ==
                   report['records'][0]['theme_images'][0]['sha256'] for part in current.parts)

    return 6
