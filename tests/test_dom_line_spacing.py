"""Direct/warm line spacing and independent cold projections follow owned inputs."""
from io import BytesIO
from zipfile import ZipFile

import pytest

from aspose.words_foss import DocxDocument, LineSpacingRule as Rule
from aspose.words_foss import light_document_model as ldm
from .test_docx_dom import W, package, payloads


def document(tmp_path, direct='', defaults=None, style='', based=''):
    styles = '<w:styles xmlns:w="' + W + '">'
    if defaults is not None:
        styles += '<w:docDefaults><w:pPrDefault><w:pPr>' + defaults + '</w:pPr></w:pPrDefault></w:docDefaults>'
    styles += '<w:style w:type="paragraph" w:default="1" w:styleId="Base"><w:pPr>' + style + '</w:pPr></w:style>'
    styles += '<w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/><w:pPr>' + based + '</w:pPr></w:style></w:styles>'
    path = tmp_path / 'source.docx'
    parts = package(path, '<w:p><w:pPr><w:pStyle w:val="Child"/>' + direct + '</w:pPr><w:r><w:t>OWNED</w:t></w:r></w:p>', {'word/styles.xml': styles.encode()})
    parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(
        b'</Types>', b'<Default Extension="bin" ContentType="application/octet-stream"/></Types>')
    with ZipFile(path, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return DocxDocument(path)


def spacing(value, token='auto'):
    return f'<w:spacing w:line="{value}" w:lineRule="{token}"/>'


@pytest.mark.parametrize('defaults,style,based,expected', [
    (None, '', '', (12.95, Rule.MULTIPLE)),
    ('', '', '', (12, Rule.MULTIPLE)),
    (spacing(480), '', '', (24, Rule.MULTIPLE)),
    (spacing(480), spacing(360, 'exact'), '', (18, Rule.EXACTLY)),
    (spacing(480), spacing(360, 'exact'), spacing(0, 'atLeast'), (0, Rule.AT_LEAST)),
    (spacing(480), '<w:spacing w:lineRule="exact"/>', '', (24, Rule.MULTIPLE)),
])
def test_inheritance_is_read_only_and_direct_unset_stays_distinct(tmp_path, defaults, style, based, expected):
    doc = document(tmp_path, defaults=defaults, style=style, based=based)
    before = payloads(doc.to_bytes())
    p = doc.body.paragraphs[0]
    assert p.paragraph_format.line_spacing is p.paragraph_format.line_spacing_rule is None
    effective = p.effective_paragraph_format
    assert (effective.line_spacing, effective.line_spacing_rule) == expected
    assert payloads(doc.to_bytes()) == before
    model = doc.to_light_document()
    pf = model.sections[0].body.paragraphs[0].paragraph_format
    assert (pf.line_spacing, pf.line_spacing_rule) == expected
    restored = ldm.Document.model_validate_json(model.model_dump_json())
    pf2 = restored.sections[0].body.paragraphs[0].paragraph_format
    assert (pf2.line_spacing, pf2.line_spacing_rule) == expected
    assert ('line_spacing' in pf2.model_fields_set) == ('line_spacing' in pf.model_fields_set)


@pytest.mark.parametrize('rule', list(Rule))
@pytest.mark.parametrize('value,warm,cold,exact', [
    (0, 0, 0, False), (1, 1, 1, False),
    (13.025, 13, 13, False), (13.075, 13.1, 13.1, False),
    (-12, -12, 12, True),
    (107374182.35, 107374182.35, .05, True),
    (107374183, 107374182.35, .05, True),
    (1e308, -107374182.4, 0, False),
    (float('nan'), -107374182.4, 0, False),
    (float('inf'), -107374182.4, 0, False),
    (float('-inf'), -107374182.4, 0, False),
])
@pytest.mark.parametrize('flat', [False, True])
def test_warm_held_values_two_saves_and_cold_json(tmp_path, rule, value, warm, cold, exact, flat):
    doc = document(tmp_path)
    p = doc.body.paragraphs[0]
    held = p.paragraph_format
    held.line_spacing_rule = rule
    held.line_spacing = value
    assert (held.line_spacing, held.line_spacing_rule) == (warm, rule)
    assert p.paragraph_format.line_spacing == warm
    cold_rule = Rule.EXACTLY if exact else Rule.AT_LEAST if cold == 0 and rule == Rule.EXACTLY else rule
    before = payloads(doc.to_bytes())
    for number in range(2):
        output = tmp_path / f'{number}.out'
        (doc.save_flat_opc if flat else doc.save)(output)
        fresh = DocxDocument(output)
        actual = fresh.body.paragraphs[0].paragraph_format
        assert (actual.line_spacing, actual.line_spacing_rule) == (cold, cold_rule)
        assert (held.line_spacing, held.line_spacing_rule) == (warm, rule)
        model = fresh.to_light_document()
        restored = ldm.Document.model_validate_json(model.model_dump_json())
        pf = restored.sections[0].body.paragraphs[0].paragraph_format
        assert (pf.line_spacing, pf.line_spacing_rule) == (cold, cold_rule)
        assert 'line_spacing' in pf.model_fields_set
    assert payloads(doc.to_bytes()) == before


@pytest.mark.parametrize('rule', list(Rule))
@pytest.mark.parametrize('twips,value,force_exact', [
    (32767, 1638.35, False), (32768, 1638.4, True), (32769, 1638.35, True),
    (65535, .05, True), (65536, 0, False), (65537, .05, False),
    (1073741823, .05, True), (1073741824, 0, False), (1073741825, .05, False),
    (2147483600, 2.4, True), (2147483645, .15, True), (2147483646, .1, True),
    (2147483647, .05, True), (-2147483648, 0, False),
])
def test_loaded_signed_sixteen_bit_measure(tmp_path, rule, twips, value, force_exact):
    doc = document(tmp_path, spacing(twips, {Rule.AT_LEAST: 'atLeast', Rule.EXACTLY: 'exact', Rule.MULTIPLE: 'auto'}[rule]))
    p = doc.body.paragraphs[0]
    kind = Rule.EXACTLY if force_exact else Rule.AT_LEAST if value == 0 and rule == Rule.EXACTLY else rule
    assert (p.paragraph_format.line_spacing, p.paragraph_format.line_spacing_rule) == (value, kind)
    assert (p.effective_paragraph_format.line_spacing, p.effective_paragraph_format.line_spacing_rule) == (value, kind)


def test_rule_assignment_materializes_local_twelve_and_spacing_inherits_rule(tmp_path):
    doc = document(tmp_path, style=spacing(480, 'exact'))
    p = doc.body.paragraphs[0]
    p.paragraph_format.line_spacing = 18
    assert p.paragraph_format.line_spacing_rule == Rule.EXACTLY
    p.paragraph_format.line_spacing = None  # Existing DOM direct-None clearing contract.
    assert p.effective_paragraph_format.line_spacing == 24
    p.paragraph_format.line_spacing_rule = Rule.EXACTLY
    assert p.paragraph_format.line_spacing == 12
    p.paragraph_format.line_spacing_rule = None
    assert p.paragraph_format.line_spacing_rule == Rule.MULTIPLE  # Missing rule with a line means auto.


def test_style_edits_update_held_paragraph_and_rebinding_has_no_stale_cache(tmp_path):
    doc = document(tmp_path, style=spacing(480))
    p = doc.body.paragraphs[0]
    held = p.paragraph_format
    base = doc.styles.get_by_id('Base')
    base.paragraph_format.line_spacing = 18
    assert p.effective_paragraph_format.line_spacing == 18
    assert held.line_spacing is None
    held.line_spacing = -12
    held.space_before = 3
    assert held.line_spacing == -12  # Unrelated formatting cannot turn warm into cold.
    doc2 = DocxDocument(BytesIO(doc.to_bytes()))
    assert doc2.body.paragraphs[0].paragraph_format.line_spacing == 12
    held.line_spacing = None
    assert p.effective_paragraph_format.line_spacing == 18


@pytest.mark.parametrize('property,value', [
    ('line_spacing', True), ('line_spacing', '12'), ('line_spacing', []),
    ('line_spacing_rule', 0), ('line_spacing_rule', 1), ('line_spacing_rule', 2),
    ('line_spacing_rule', True), ('line_spacing_rule', 'exact'),
])
def test_invalid_setter_is_atomic(tmp_path, property, value):
    doc = document(tmp_path)
    before = payloads(doc.to_bytes())
    with pytest.raises(TypeError):
        setattr(doc.body.paragraphs[0].paragraph_format, property, value)
    assert payloads(doc.to_bytes()) == before


@pytest.mark.parametrize('value', ['240foo', '12 cm', '240 ', '12PT'])
def test_invalid_source_measure_has_explicit_error(tmp_path, value):
    doc = document(tmp_path, spacing(value))
    before = payloads(doc.to_bytes())
    with pytest.raises(RuntimeError, match='measurement unit'):
        _ = doc.body.paragraphs[0].paragraph_format.line_spacing
    assert payloads(doc.to_bytes()) == before


@pytest.mark.parametrize('value,expected', [
    ('', (0, Rule.MULTIPLE)), ('bad', (0, Rule.MULTIPLE)),
    ('1.5', (.05, Rule.MULTIPLE)), ('2e3', (100, Rule.MULTIPLE)),
    ('12pt', (12, Rule.MULTIPLE)), ('12.53pt', (12.55, Rule.MULTIPLE)),
    ('12.525pt', (12.5, Rule.MULTIPLE)), ('12.575pt', (12.6, Rule.MULTIPLE)),
    ('1mm', (2.85, Rule.MULTIPLE)), ('1.99mm', (5.65, Rule.MULTIPLE)),
    ('1cm', (28.35, Rule.MULTIPLE)), ('1in', (72, Rule.MULTIPLE)),
    ('1pc', (12, Rule.MULTIPLE)), ('1pi', (12, Rule.MULTIPLE)),
    ('-12pt', (12, Rule.EXACTLY)), ('9007199254740993', (0, Rule.MULTIPLE)),
    ('1e999', (.05, Rule.EXACTLY)), ('-1e308', (0, Rule.MULTIPLE)),
    ('+ 240', (0, Rule.MULTIPLE)), ('１２', (0, Rule.MULTIPLE)),
])
def test_loaded_numeric_and_universal_measure_lexical_forms(tmp_path, value, expected):
    doc = document(tmp_path, spacing(value))
    pf = doc.body.paragraphs[0].paragraph_format
    assert (pf.line_spacing, pf.line_spacing_rule) == expected
    model = doc.to_light_document()
    actual = model.sections[0].body.paragraphs[0].paragraph_format
    assert (actual.line_spacing, actual.line_spacing_rule) == expected


@pytest.mark.parametrize('flat', [False, True])
@pytest.mark.parametrize('failure', ['publish', 'budget'])
def test_failed_save_preserves_existing_output_and_warm_node(tmp_path, monkeypatch, flat, failure):
    from aspose.words_foss import _io
    doc = document(tmp_path)
    p = doc.body.paragraphs[0]
    held = p.paragraph_format
    held.line_spacing = -12
    before = p.xml
    output = tmp_path / 'existing.out'
    output.write_bytes(b'original')
    if failure == 'publish':
        def fail(*args):
            raise OSError('publish failed')
        monkeypatch.setattr(_io.os, 'replace', fail)
    else:
        monkeypatch.setattr(_io, 'MAX_INPUT_BYTES', 10)
    with pytest.raises((OSError, ValueError)):
        (doc.save_flat_opc if flat else doc.save)(output)
    assert output.read_bytes() == b'original'
    assert p.xml == before and held.line_spacing == -12
