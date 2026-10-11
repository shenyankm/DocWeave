"""Owned row-format direct/style provenance, warm values and atomic saving."""
from io import BytesIO
from zipfile import ZipFile

import pytest

import dockiln as dk
from dockiln import light_document_model as ldm
from dockiln.dom import DocxDocument
from dockiln.model.enums import HeightRule
from .test_docx_dom import package


def row_document(tmp_path, properties='', style_properties='', *, conditional='', default=False):
    path = tmp_path / 'input.docx'
    style = (f'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
             f'<w:style w:type="table" w:styleId="Base" w:default="{int(default)}">'
             f'<w:name w:val="Base"/>{style_properties}{conditional}</w:style>'
             '<w:style w:type="table" w:styleId="Child"><w:name w:val="Child"/>'
             '<w:basedOn w:val="Base"/></w:style></w:styles>')
    body = ('<w:tbl><w:tblPr>' + ('' if default else '<w:tblStyle w:val="Child"/>')
            + '</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr>'
            + properties + '</w:trPr><w:tc><w:p><w:r><w:t>Owned</w:t></w:r></w:p>'
            '</w:tc></w:tr></w:tbl><w:sectPr/>')
    original = package(path, body=body, extras={'word/styles.xml': style.encode()})
    original['[Content_Types].xml'] = original['[Content_Types].xml'].replace(
        b'</Types>', b'<Default Extension="bin" ContentType="application/octet-stream"/></Types>')
    with ZipFile(path, 'w') as archive:
        for name, data in original.items():
            archive.writestr(name, data)
    return path, original


def values(properties):
    return properties.height, int(properties.height_rule), properties.allow_break_across_pages


def ldm_row(document):
    return document.get_child_nodes(ldm.NodeType.ROW, True)[0]


@pytest.mark.parametrize('value,expected', [
    ('0', 0), ('1', .05), ('-20', -1), ('2147483647', 107374182.35),
    ('2147483648', -107374182.4), ('4294967295', -.05), ('4294967296', 0),
    ('9223372036854775808', -.05), ('1_13', 0), ('1.5', .05), ('1e1', .5),
    ('NaN', 0), ('0.025pt', 0), ('0.026pt', .05), ('0.075pt', .1),
    ('0.125pt', .1), ('1in', 72), ('1cm', 28.35), ('1mm', 2.85), ('1pc', 12), ('1pi', 12),
])
@pytest.mark.parametrize('rule', ['auto', 'atLeast', 'exact', 'unknown'])
def test_source_measure_shared_between_dom_and_ldm(tmp_path, value, expected, rule):
    path, _ = row_document(tmp_path, f'<w:trHeight w:val="{value}" w:hRule="{rule}"/>')
    expected_rule = 2 if expected == 0 else 1 if rule == 'exact' else 0
    assert values(DocxDocument(path).body.tables[0].rows[0].row_format) == (expected, expected_rule, True)
    assert values(ldm_row(dk.Document(path)).row_format) == (expected, expected_rule, True)


@pytest.mark.parametrize('value,allowed', [('', False), ('on', False), ('1', False), ('true', False),
                                        ('off', True), ('0', True), ('false', True)])
def test_split_value_not_element_presence(tmp_path, value, allowed):
    path, _ = row_document(tmp_path, f'<w:cantSplit w:val="{value}"/>')
    assert DocxDocument(path).body.tables[0].rows[0].row_format.allow_break_across_pages is allowed
    assert ldm_row(dk.Document(path)).row_format.allow_break_across_pages is allowed


@pytest.mark.parametrize('declaration', ['<w:cantSplit w:val="TRUE"/>', '<w:cantSplit w:val="unknown"/>',
                                       '<w:trHeight w:val="1PT"/>', '<w:trHeight w:val="1 pt"/>',
                                       '<w:cantSplit w:val="unknown"/><w:cantSplit w:val="1"/>'])
def test_invalid_source_is_rejected_without_echoing_value(tmp_path, declaration):
    path, _ = row_document(tmp_path, declaration)
    for constructor in (DocxDocument, dk.Document):
        with pytest.raises(RuntimeError, match='Invalid row-') as error:
            constructor(path)
        assert str(path) not in str(error.value)
        assert declaration not in str(error.value)


@pytest.mark.parametrize('property_name,invalid', [('height', None), ('height', True), ('height', '2'),
    ('height_rule', None), ('height_rule', 1), ('height_rule', True), ('height_rule', 'exact'),
    ('allow_break_across_pages', None), ('allow_break_across_pages', 0), ('allow_break_across_pages', 'true')])
def test_setter_errors_preserve_package(tmp_path, property_name, invalid):
    path, _ = row_document(tmp_path)
    document = DocxDocument(path)
    properties = document.body.tables[0].rows[0].row_format
    before = document.to_bytes()
    with pytest.raises(TypeError):
        setattr(properties, property_name, invalid)
    assert values(properties) == (0, 2, True)
    assert document.to_bytes() == before


@pytest.mark.parametrize('number,expected', [(0, 0), (.025, 0), (.075, .1), (.125, .1), (-.026, -.05),
    (float('nan'), -107374182.4), (float('inf'), -107374182.4), (float('-inf'), -107374182.4),
    (1e100, 107374182.35), (-1e100, -107374182.4)])
def test_numeric_setter_quantizes_and_keeps_enum(tmp_path, number, expected):
    path, _ = row_document(tmp_path)
    properties = DocxDocument(path).body.tables[0].rows[0].row_format
    properties.height = number
    assert values(properties) == (expected, 2 if expected == 0 else 0, True)
    assert isinstance(properties.height_rule, HeightRule)


@pytest.mark.parametrize('rule', [HeightRule.EXACTLY, HeightRule.AT_LEAST])
@pytest.mark.parametrize('flat', [False, True])
def test_zero_rule_warm_is_not_cold_projection(tmp_path, rule, flat):
    path, _ = row_document(tmp_path)
    document = DocxDocument(path)
    row = document.body.tables[0].rows[0]
    properties = row.row_format
    properties.height_rule = rule
    destination = tmp_path / ('out.xml' if flat else 'out.docx')
    save = document.save_flat_opc if flat else document.save
    for _ in range(2):
        save(destination)
        assert properties.height == 0
        assert properties.height_rule == rule
        assert document.body.tables[0].rows[0] is row
        assert values(DocxDocument(destination).body.tables[0].rows[0].row_format) == (0, 2, True)
    row._element.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'trHeight')[0].setAttributeNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:val', '20')
    assert values(properties) == (1, int(rule), True)


@pytest.mark.parametrize('default', [False, True])
@pytest.mark.parametrize('height', [20, 2000])
def test_inherited_style_preserved_without_direct_height(tmp_path, default, height):
    direct = f'<w:trPr><w:cantSplit/><w:trHeight w:val="{height * 20}" w:hRule="exact"/></w:trPr>'
    path, _ = row_document(tmp_path, style_properties=direct, default=default)
    document = dk.Document(path)
    row = ldm_row(document)
    assert values(row.row_format) == (height, 1, False)
    assert row.row_format.height_explicit is False
    cloned = ldm.Document.model_validate_json(document.light_document_model.model_dump_json())
    assert values(ldm_row(cloned).row_format) == (height, 1, False)
    for format_ in (dk.SaveFormat.DOCX, dk.SaveFormat.FLAT_OPC):
        destination = tmp_path / ('style.docx' if format_ == dk.SaveFormat.DOCX else 'style.xml')
        document.save(destination, format_)
        assert values(row.row_format) == (height, 1, False)
        assert values(DocxDocument(destination).body.tables[0].rows[0].row_format) == (height, 1, False)
        assert values(ldm_row(dk.Document(destination)).row_format) == (height, 1, False)
    row.row_format.height = 3
    row.row_format.allow_break_across_pages = True
    assert row.row_format.height_explicit is True
    assert row.row_format.allow_break_explicit is True


@pytest.mark.parametrize('flat', [False, True])
def test_direct_overrides_and_conditional_metadata_survive(tmp_path, flat):
    conditional = '<w:tblStylePr w:type="firstRow"><w:trPr><w:trHeight w:val="1400" w:hRule="exact"/></w:trPr></w:tblStylePr>'
    path, original = row_document(tmp_path, '<w:trHeight w:val="0" w:hRule="exact"/><w:cantSplit w:val="false"/>',
        '<w:trPr><w:trHeight w:val="400" w:hRule="exact"/><w:cantSplit/></w:trPr>', conditional=conditional)
    document = DocxDocument(path)
    properties = document.body.tables[0].rows[0].row_format
    assert values(properties) == (0, 2, True)
    properties.height = -1
    properties.height_rule = HeightRule.EXACTLY
    properties.allow_break_across_pages = False
    destination = tmp_path / ('edited.xml' if flat else 'edited.docx')
    (document.save_flat_opc if flat else document.save)(destination)
    assert values(DocxDocument(destination).body.tables[0].rows[0].row_format) == (-1, 1, False)
    assert values(ldm_row(dk.Document(destination)).row_format) == (-1, 1, False)
    with ZipFile(BytesIO(document.to_bytes())) as archive:
        assert archive.read('customXml/item.xml') == original['customXml/item.xml']
    snapshot = document.to_light_document()
    style = next(style for style in snapshot.styles if style.name == 'Base')
    assert style.row_format.height == 20
    assert style.table_style_properties[0].row_format.height == 70


@pytest.mark.parametrize('legacy', [False, True])
@pytest.mark.parametrize('flat', [False, True])
def test_save_clamp_is_success_only_and_identity_preserving(tmp_path, monkeypatch, legacy, flat):
    path, _ = row_document(tmp_path, '<w:trHeight w:val="40" w:hRule="exact"/>')
    document = dk.Document(path) if legacy else DocxDocument(path)
    properties = ldm_row(document).row_format if legacy else document.body.tables[0].rows[0].row_format
    properties.height = 2000
    before = values(properties)
    destination = tmp_path / ('result.xml' if flat else 'result.docx')
    destination.write_bytes(b'old output')
    save = (lambda p: document.save(p, dk.SaveFormat.FLAT_OPC if flat else dk.SaveFormat.DOCX)) if legacy else (document.save_flat_opc if flat else document.save)
    from dockiln import _io
    replace = _io.os.replace
    monkeypatch.setattr(_io.os, 'replace', lambda *args: (_ for _ in ()).throw(OSError('owned failure')))
    with pytest.raises(OSError):
        save(destination)
    assert values(properties) == before
    assert destination.read_bytes() == b'old output'
    monkeypatch.setattr(_io.os, 'replace', replace)
    for _ in range(2):
        save(destination)
        assert properties.height == 1584
        assert values(DocxDocument(destination).body.tables[0].rows[0].row_format) == (1584, 1, True)
        if legacy:
            assert ldm_row(document).row_format is properties


def test_duplicate_height_last_wins_and_setter_canonicalizes(tmp_path):
    path, _ = row_document(tmp_path, '<w:trHeight w:val="40" w:hRule="exact"/><w:trHeight w:val="80" w:hRule="exact"/>')
    document = DocxDocument(path)
    properties = document.body.tables[0].rows[0].row_format
    assert properties.height == 4
    assert ldm_row(dk.Document(path)).row_format.height == 4
    properties.height = 5
    assert properties.height == 5
    assert document.part_xml('word/document.xml').count('trHeight') == 1


@pytest.mark.parametrize('legacy', [False, True])
@pytest.mark.parametrize('flat', [False, True])
def test_pure_export_and_budget_failure_keep_original_owner(tmp_path, monkeypatch, legacy, flat):
    path, _ = row_document(tmp_path, '<w:trHeight w:val="40" w:hRule="exact"/>')
    document = dk.Document(path) if legacy else DocxDocument(path)
    properties = ldm_row(document).row_format if legacy else document.body.tables[0].rows[0].row_format
    properties.height = 2000
    format_ = dk.SaveFormat.FLAT_OPC if flat else dk.SaveFormat.DOCX
    data = document.to_bytes(format_) if legacy else (document.to_flat_opc() if flat else document.to_bytes())
    assert properties.height == 2000
    assert DocxDocument(BytesIO(data)).body.tables[0].rows[0].row_format.height == 1584
    source = document.light_document_model.source_font_table if legacy else document.part_xml('word/document.xml')
    destination = tmp_path / 'budget-result.docx'
    destination.write_bytes(b'original')
    from dockiln import _io
    monkeypatch.setattr(_io, 'MAX_PART_BYTES', 32)
    with pytest.raises(ValueError):
        document.save(destination, format_) if legacy else (document.save_flat_opc(destination) if flat else document.save(destination))
    assert destination.read_bytes() == b'original'
    assert properties.height == 2000
    assert (document.light_document_model.source_font_table is source) if legacy else document.part_xml('word/document.xml') == source


def test_legacy_json_and_copy_edits_keep_internal_integer_rules():
    original = ldm.RowFormat.model_validate({'height': 12, 'height_rule': 1})
    assert original.height_explicit is None
    assert original.allow_break_explicit is None
    inherited = ldm.RowFormat(height=20, height_rule=1, height_explicit=False, allow_break_explicit=False)
    changed = inherited.model_copy(update={'height': 7, 'allow_break_across_pages': False})
    assert changed.height_explicit is True
    assert changed.allow_break_explicit is True
    assert values(inherited) == (20, 1, True)


@pytest.mark.parametrize('format_', [dk.SaveFormat.DOCX, dk.SaveFormat.FLAT_OPC])
def test_bad_resource_graph_does_not_commit_font_or_height(tmp_path, format_):
    from .test_source_font_table import archive, owned_parts, FONT
    document = dk.Document(BytesIO(archive(owned_parts())))
    model = document.light_document_model
    properties = ldm.RowFormat(height=2000, height_rule=1)
    model.sections[0].body.children.append(ldm.Table(rows=[ldm.Row(row_format=properties,
        cells=[ldm.Cell(paragraphs=[ldm.Paragraph(children=[ldm.Run(text='Owned')])])])]))
    source = model.source_font_table
    broken = source.model_copy(update={'parts': tuple(part for part in source.parts if part.name != FONT)})
    model.source_font_table = broken
    held = document.font_infos
    registrations = tuple(font.name for font in held)
    before = model.model_dump_json()
    destination = tmp_path / 'bad-graph.docx'
    destination.write_bytes(b'original')
    with pytest.raises(ValueError):
        document.save(destination, format_)
    assert properties.height == 2000
    assert model.model_dump_json() == before
    assert model.source_font_table is broken
    assert tuple(font.name for font in held) == registrations
    assert tuple(font.name for font in document.font_infos) == registrations
    assert destination.read_bytes() == b'original'
