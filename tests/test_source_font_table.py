"""Owned font-table graphs survive conversion without private font APIs."""
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from aspose.words_foss import Document, SaveFormat
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._opc import resolve_target, relationships_path
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.saving import OoxmlSaveOptions

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'
TABLE = 'owned/tables/fonts.xml'
FONT = 'owned/fonts/tiny.odttf'


def tiny_font():
    from fontTools.fontBuilder import FontBuilder
    from fontTools.pens.ttGlyphPen import TTGlyphPen
    builder = FontBuilder(1000, isTTF=True)
    builder.setupGlyphOrder(['.notdef', 'A', 'B'])
    builder.setupCharacterMap({65: 'A', 66: 'B'})
    glyphs = {}
    for name in ['.notdef', 'A', 'B']:
        pen = TTGlyphPen(None)
        pen.moveTo((0, 0)); pen.lineTo((300, 0)); pen.lineTo((150, 500)); pen.closePath()
        glyphs[name] = pen.glyph()
    builder.setupGlyf(glyphs)
    builder.setupHorizontalMetrics({name: (500, 0) for name in glyphs})
    builder.setupHorizontalHeader(ascent=800, descent=-200)
    builder.setupNameTable({'familyName': 'Owned Tiny', 'styleName': 'Regular',
                           'uniqueFontIdentifier': 'Owned Tiny Regular',
                           'fullName': 'Owned Tiny Regular', 'psName': 'OwnedTiny-Regular'})
    builder.setupOS2(sTypoAscender=800, sTypoDescender=-200, usWinAscent=800, usWinDescent=200)
    builder.setupPost(); builder.setupMaxp()
    output = BytesIO(); builder.save(output)
    return output.getvalue()


def rels(*edges):
    root = ET.Element(f'{{{P}}}Relationships')
    for rid, kind, target in edges:
        ET.SubElement(root, f'{{{P}}}Relationship', Id=rid, Type=R + '/' + kind, Target=target)
    return ET.tostring(root)


def owned_parts(flag=True):
    table = (f'<w:fonts xmlns:w="{W}" xmlns:r="{R}"><w:font w:name="Owned Tiny">'
             '<w:altName w:val="Owned Alias"/><w:charset w:val="00"/>'
             '<w:family w:val="swiss"/><w:pitch w:val="variable"/>'
             '<w:panose1 w:val="020B0604020202020204"/>'
             '<w:embedRegular r:id="font" w:fontKey="{00000000-0000-0000-0000-000000000000}"/>'
             '</w:font></w:fonts>').encode()
    settings = f'<w:settings xmlns:w="{W}">'
    if flag is not None:
        settings += '<w:embedTrueTypeFonts' + ('/>' if flag else ' w:val="0"/>')
    settings += '<w:doNotEmbedSystemFonts w:val="0"/><w:saveSubsetFonts/></w:settings>'
    parts = {'word/document.xml': (f'<w:document xmlns:w="{W}"><w:body><w:p><w:r>'
              '<w:rPr><w:rFonts w:ascii="Owned Tiny" w:hAnsi="Owned Tiny"/></w:rPr>'
              '<w:t>AB</w:t></w:r></w:p><w:sectPr/></w:body></w:document>').encode(),
             '_rels/.rels': rels(('doc', 'officeDocument', 'word/document.xml')),
             'word/_rels/document.xml.rels': rels(('fonts', 'fontTable', '../' + TABLE),
                                                 ('settings', 'settings', 'settings.xml')),
             'word/settings.xml': settings.encode(), TABLE: table, FONT: tiny_font(),
             relationships_path(TABLE): rels(('font', 'font', '../fonts/tiny.odttf'))}
    types = ET.Element(f'{{{C}}}Types')
    ET.SubElement(types, f'{{{C}}}Default', Extension='rels', ContentType='application/vnd.openxmlformats-package.relationships+xml')
    for name, kind in [('word/document.xml', 'document.main'), (TABLE, 'fontTable'), ('word/settings.xml', 'settings')]:
        ET.SubElement(types, f'{{{C}}}Override', PartName='/' + name,
                      ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.' + kind + '+xml')
    ET.SubElement(types, f'{{{C}}}Override', PartName='/' + FONT,
                  ContentType='application/vnd.openxmlformats-officedocument.obfuscatedFont')
    parts['[Content_Types].xml'] = ET.tostring(types)
    return parts


def archive(parts):
    output = BytesIO()
    with ZipFile(output, 'w') as package:
        for name, data in parts.items():
            package.writestr(name, data)
    return output.getvalue()


@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
@pytest.mark.parametrize('flag', [None, False, True])
def test_metadata_and_embedded_graph_survive_two_cold_saves(format, flag):
    parts = owned_parts(flag)
    model = Document(BytesIO(archive(parts))).light_document_model
    before = model.model_dump_json()
    assert model.source_font_table.part_name == TABLE
    assert model.font_embedding.embed_true_type_fonts is flag
    model = ldm.Document.model_validate_json(before)
    for _ in range(2):
        saved = LdmDocxWriter(OoxmlSaveOptions(format)).write_to_bytes(model)
        assert model.model_dump_json() == before
        model = Document(BytesIO(saved)).light_document_model
        snapshot = model.source_font_table
        if format == SaveFormat.DOCX:
            assert snapshot.data == parts[TABLE]
        else:
            assert ET.canonicalize(snapshot.data, rewrite_prefixes=True) == ET.canonicalize(parts[TABLE], rewrite_prefixes=True)
        font = next(part for part in snapshot.parts if part.content_type.endswith('obfuscatedFont'))
        assert font.data == parts[FONT]
        assert model.font_embedding.model_dump() == {'embed_true_type_fonts': flag,
            'do_not_embed_system_fonts': False, 'save_subset_fonts': True}
        captured = {part.name for part in snapshot.parts}
        for part in snapshot.parts:
            if part.name.endswith('.rels'):
                parent = part.name.replace('/_rels/', '/').removesuffix('.rels')
                for edge in ET.fromstring(part.data):
                    assert resolve_target(parent, edge.get('Target')) in captured
        before = model.model_dump_json()


@pytest.mark.parametrize('damage', ['missing', 'duplicate', 'escape', 'dangling', 'unrelated', 'wrong-root'])
def test_invalid_snapshot_keeps_existing_output_and_model(tmp_path, damage):
    model = Document(BytesIO(archive(owned_parts()))).light_document_model
    source = model.source_font_table
    parts = list(source.parts)
    if damage == 'missing':
        parts = [part for part in parts if part.name != FONT]
    elif damage == 'duplicate':
        parts.append(parts[0])
    elif damage in {'escape', 'unrelated'}:
        parts.append(ldm.SourceThemePart(name='../bad.bin' if damage == 'escape' else 'extra.bin',
                                         content_type='application/octet-stream', data=b'owned'))
    else:
        data = source.data.replace(b'r:id="font"', b'r:id="missing"') if damage == 'dangling' else b'<wrong/>'
        source = source.model_copy(update={'data': data})
        parts = [part.model_copy(update={'data': data}) if part.name == TABLE else part for part in parts]
    model.source_font_table = source.model_copy(update={'parts': tuple(parts)})
    before = model.model_dump_json()
    output = tmp_path / 'existing.docx'; output.write_bytes(b'existing')
    with pytest.raises(ValueError):
        LdmDocxWriter().write(model, output)
    assert output.read_bytes() == b'existing'
    assert model.model_dump_json() == before


def test_budget_checked_before_parsing(monkeypatch, tmp_path):
    from aspose.words_foss import _io
    model = Document(BytesIO(archive(owned_parts()))).light_document_model
    monkeypatch.setattr(_io, 'MAX_INPUT_BYTES', 10)
    output = tmp_path / 'existing.docx'; output.write_bytes(b'existing')
    with pytest.raises(ValueError, match='Input exceeds'):
        LdmDocxWriter().write(model, output)
    assert output.read_bytes() == b'existing'


def test_old_json_and_flag_validation():
    assert ldm.Document.model_validate({}).source_font_table is None
    assert ldm.Document.model_validate({}).font_embedding is None
    with pytest.raises(ValueError):
        ldm.FontEmbeddingSettings(embed_true_type_fonts=1)
    with pytest.raises(TypeError):
        ldm.FontEmbeddingSettings().model_copy(update={'save_subset_fonts': 1})


@pytest.mark.parametrize('format', [SaveFormat.DOCX, SaveFormat.FLAT_OPC])
def test_font_and_theme_resources_have_independent_graphs(format):
    parts = owned_parts()
    theme = 'owned/theme/theme.xml'
    image = 'owned/theme/resource.bin'
    parts[theme] = b'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + R.encode() + b'" r:id="resource"/>'
    parts[image] = b'owned theme resource'
    parts[relationships_path(theme)] = rels(('resource', 'image', 'resource.bin'))
    types = ET.fromstring(parts['[Content_Types].xml'])
    for name, kind in [(theme, 'application/xml'), (image, 'application/octet-stream')]:
        ET.SubElement(types, f'{{{C}}}Override', PartName='/' + name, ContentType=kind)
    parts['[Content_Types].xml'] = ET.tostring(types)
    root = ET.fromstring(parts['word/_rels/document.xml.rels'])
    ET.SubElement(root, f'{{{P}}}Relationship', Id='theme', Type=R + '/theme', Target='../' + theme)
    parts['word/_rels/document.xml.rels'] = ET.tostring(root)
    model = Document(BytesIO(archive(parts))).light_document_model
    saved = LdmDocxWriter(OoxmlSaveOptions(format)).write_to_bytes(model)
    cold = Document(BytesIO(saved)).light_document_model
    assert any(part.data == parts[image] for part in cold.source_theme.parts)
    assert any(part.data == parts[FONT] for part in cold.source_font_table.parts)
    assert not {part.name for part in cold.source_theme.parts} & {part.name for part in cold.source_font_table.parts}
