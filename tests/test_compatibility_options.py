"""Self-contained compatibility preferences and version-profile regressions."""
import io
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from defusedxml.common import DefusedXmlException

import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer.settings_part import render_settings_xml
from aspose.words_foss.settings import MsWordVersion

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def document(compat='', ui=''):
    out=io.BytesIO()
    with ZipFile(out,'w') as z:
        z.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
        z.writestr('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr('word/document.xml','<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>OWNED</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')
        z.writestr('word/_rels/document.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
        z.writestr('word/settings.xml','<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat>'+compat+'</w:compat>'+ui+'</w:settings>')
    return aw.Document(io.BytesIO(out.getvalue()))


@pytest.mark.parametrize('fmt',[aw.SaveFormat.DOCX,aw.SaveFormat.FLAT_OPC])
@pytest.mark.parametrize('version',list(MsWordVersion))
def test_version_profiles_preserve_trailing_underlining_and_roundtrip(tmp_path,fmt,version):
    doc=document()
    held=doc.compatibility_options
    held.ul_trail_space=True
    held.optimize_for(version)
    assert doc.compatibility_options is held
    assert held.ul_trail_space
    assert held.do_not_break_wrapped_tables is (version == MsWordVersion.WORD2000)
    assert held.grow_autofit is (version in (MsWordVersion.WORD2000,MsWordVersion.WORD2002))
    source=doc.light_document_model.source_settings
    for i in range(2):
        path=tmp_path/str(i)
        doc.save(path,fmt)
        cold=aw.Document(path,load_options=_format(fmt))
        assert cold.compatibility_options.ul_trail_space
        assert cold.compatibility_options.do_not_break_wrapped_tables == held.do_not_break_wrapped_tables
        assert doc.light_document_model.source_settings is source


def _format(fmt):
    options=aw.LoadOptions()
    options.load_format=aw.LoadFormat.DOCX if fmt == aw.SaveFormat.DOCX else aw.LoadFormat.FLAT_OPC
    return options


@pytest.mark.parametrize('fragment,ui',[
    ('<w:doNotExpandShiftReturn w:val="0"/>',''),
    ('','<w:uiCompat97To2003 w:val="0"/>'),
])
@pytest.mark.parametrize('fmt',[aw.SaveFormat.DOCX,aw.SaveFormat.FLAT_OPC])
def test_false_source_declarations_keep_warm_defaults_and_recover_cold(tmp_path,fragment,ui,fmt):
    doc=document(fragment,ui)
    held=doc.compatibility_options
    source=doc.light_document_model.source_settings
    assert not held.disable_open_type_font_formatting_features
    for i in range(2):
        path=tmp_path/str(i)
        doc.save(path,fmt)
        cold=aw.Document(path,load_options=_format(fmt))
        assert cold.compatibility_options.disable_open_type_font_formatting_features
        assert not held.disable_open_type_font_formatting_features
        assert doc.light_document_model.source_settings is source


@pytest.mark.parametrize('name,tag',[
    ('shape_layout_like_ww8','shapeLayoutLikeWW8'),
    ('do_not_use_html_paragraph_auto_spacing','doNotUseHTMLParagraphAutoSpacing'),
    ('use_single_borderfor_contiguous_cells','useSingleBorderforContiguousCells'),
    ('suppress_spacing_at_top_of_page','suppressSpacingAtTopOfPage'),
    ('do_not_expand_shift_return','doNotExpandShiftReturn'),
    ('swap_borders_facing_pgs','swapBordersFacingPages'),
])
def test_public_preferences_json_and_exact_wordml_names(name,tag):
    doc=document()
    setattr(doc.compatibility_options,name,True)
    model=ldm.Document.model_validate_json(doc.light_document_model.model_dump_json())
    assert ET.fromstring(render_settings_xml(model)).find(W+'compat/'+W+tag) is not None
    if name == 'do_not_expand_shift_return':
        assert model.do_not_expand_shift_return


@pytest.mark.parametrize('value',[0,1,-1,8,True,False,1.0,None,'WORD2010',[]])
def test_optimize_requires_real_word_version_and_does_not_mutate(value):
    doc=aw.Document()
    options=doc.compatibility_options
    before=doc.light_document_model.model_dump_json()
    with pytest.raises(TypeError): options.optimize_for(value)
    assert doc.light_document_model.model_dump_json() == before


@pytest.mark.parametrize('name',['do_not_expand_shift_return','ul_trail_space','ui_compat_97_to_2003','transparent_metafiles'])
@pytest.mark.parametrize('value',[0,1,None,'false',1.0])
def test_bool_setters_reject_other_values_without_mutation(name,value):
    doc=aw.Document()
    options=doc.compatibility_options
    before=doc.light_document_model.model_dump_json()
    with pytest.raises(TypeError):setattr(options,name,value)
    assert doc.light_document_model.model_dump_json() == before


@pytest.mark.parametrize('fmt',[aw.SaveFormat.DOCX,aw.SaveFormat.FLAT_OPC])
def test_profile_rebuilds_typed_settings_keeps_foreign_siblings_and_failed_publication_atomic(tmp_path,monkeypatch,fmt):
    doc=document('<x:extension xmlns:x="urn:owned" x:value="keep"/><w:compatSetting w:name="Owned" w:uri="urn:owned" w:val="keep"/>')
    options=doc.compatibility_options
    options.optimize_for(MsWordVersion.WORD2019)
    original=doc.light_document_model.source_settings
    before=doc.light_document_model.model_dump_json()
    path=tmp_path/'destination';path.write_bytes(b'original')
    def failed(*args):raise OSError('owned failure')
    with monkeypatch.context() as m:
        m.setattr('os.replace',failed)
        with pytest.raises(OSError):doc.save(path,fmt)
    assert path.read_bytes() == b'original'
    assert doc.light_document_model.model_dump_json() == before
    assert doc.light_document_model.source_settings is original
    with pytest.warns(aw.ContentLossWarning, match='normalized'):
        doc.save(path,fmt)
    cold=aw.Document(path,load_options=_format(fmt))
    source=ET.fromstring(cold.light_document_model.source_settings.data)
    assert source.find('{urn:owned}extension').get('{urn:owned}value') == 'keep'
    assert not any(c.get(W+'uri') == 'urn:owned' for c in source)
    assert b'Owned' in original.data


def test_defaults_are_distinct_and_constructor_readonly_contract():
    new=aw.Document()
    loaded=document()
    assert new.compatibility_options.grow_autofit
    assert not new.compatibility_options.disable_open_type_font_formatting_features
    assert not loaded.compatibility_options.grow_autofit
    assert loaded.compatibility_options.disable_open_type_font_formatting_features
    with pytest.raises(TypeError):aw.settings.CompatibilityOptions()
    with pytest.raises(AttributeError):new.compatibility_options=loaded.compatibility_options
    with pytest.raises(AttributeError):new.compatibility_options.owned_unknown=True
    with pytest.raises(AttributeError):getattr(new.compatibility_options,'owned_unknown')
    new.compatibility_options.optimize_for(version=aw.settings.MsWordVersion.WORD2019)


def test_invalid_state_and_source_cannot_serialize_or_expose_contents():
    doc=document()
    doc.compatibility_options
    obj=doc.light_document_model.model_dump()
    obj['compatibility_options_state']['flags']['owned-secret']=True
    with pytest.raises(ValueError,match='supported preferences') as error:
        ldm.Document.model_validate(obj)
    assert 'owned-secret' not in str(error.value)


@pytest.mark.parametrize('name,tag',[('do_not_expand_shift_return','doNotExpandShiftReturn'),
    ('grow_autofit','growAutofit'),('use_fe_layout','useFELayout')])
@pytest.mark.parametrize('first,last',[('0','1'),('1','0')])
def test_repeated_known_flags_use_last_declaration_with_immutable_source(name,tag,first,last):
    doc=document('<w:'+tag+' w:val="'+first+'"/><w:'+tag+' w:val="'+last+'"/>')
    original=doc.light_document_model.source_settings
    assert getattr(doc.compatibility_options,name) == (last == '1')
    root=ET.fromstring(render_settings_xml(doc.light_document_model)).find(W+'compat')
    assert len(root.findall(W+tag)) == int(last == '1')
    assert doc.light_document_model.source_settings is original
    assert len(ET.fromstring(original.data).findall(W+tag)) == 2


def test_positive_known_records_keep_foreign_extensions_when_edited():
    doc=document('<w:growAutofit xmlns:x="urn:owned" x:value="keep" w:val="0"><x:extension/></w:growAutofit>',
        '<w:uiCompat97To2003 xmlns:x="urn:owned" x:value="keep" w:val="0"><x:extension/></w:uiCompat97To2003>')
    doc.compatibility_options.grow_autofit=True
    doc.compatibility_options.ui_compat_97_to_2003=True
    model=ldm.Document.model_validate_json(doc.light_document_model.model_dump_json())
    xml=ET.fromstring(render_settings_xml(model))
    for path in (W+'compat/'+W+'growAutofit',W+'uiCompat97To2003'):
        node=xml.find(path)
        assert node.get('{urn:owned}value') == 'keep'
        assert node.find('{urn:owned}extension') is not None


@pytest.mark.parametrize('data',[b'<broken',b'<foreign/>',b'<!DOCTYPE x><x/>'])
def test_bad_ui_snapshot_cannot_publish_even_when_false(tmp_path,data):
    doc=document()
    source=ldm.SourceSettings(ui_compat_97_to_2003=False,ui_compat_data=data)
    doc.light_document_model.source_settings=source
    target=tmp_path/'output';target.write_bytes(b'previous')
    with pytest.raises((ValueError,ET.ParseError,DefusedXmlException)):doc.save(target,aw.SaveFormat.DOCX)
    assert target.read_bytes() == b'previous'
    assert doc.light_document_model.source_settings is source


@pytest.mark.parametrize('new_document', [False, True])
@pytest.mark.parametrize('first', [MsWordVersion.WORD2000, MsWordVersion.WORD2003])
@pytest.mark.parametrize('fmt', [aw.SaveFormat.DOCX, aw.SaveFormat.FLAT_OPC])
def test_sequential_profiles_restore_implicit_defaults_without_losing_explicit_edits(tmp_path, new_document, first, fmt):
    doc = aw.Document() if new_document else document()
    held = doc.compatibility_options
    held.optimize_for(first)
    assert not held.disable_open_type_font_formatting_features
    held.optimize_for(MsWordVersion.WORD2019)
    assert held.disable_open_type_font_formatting_features is (not new_document)
    model = ldm.Document.model_validate_json(doc.light_document_model.model_dump_json())
    restored = aw.Document()
    restored._document = model
    assert restored.compatibility_options.disable_open_type_font_formatting_features is (not new_document)
    for i in range(2):
        path = tmp_path / str(i)
        doc.save(path, fmt)
        assert doc.compatibility_options is held
        assert held.disable_open_type_font_formatting_features is (not new_document)
        assert aw.Document(path, load_options=_format(fmt)).compatibility_options.disable_open_type_font_formatting_features
    held.disable_open_type_font_formatting_features = True
    held.optimize_for(first)
    held.optimize_for(MsWordVersion.WORD2019)
    assert not held.disable_open_type_font_formatting_features


@pytest.mark.parametrize('version,mode,override', [
    (MsWordVersion.WORD2000, None, '1'), (MsWordVersion.WORD2003, '11', '1'),
    (MsWordVersion.WORD2007, '12', '0'), (MsWordVersion.WORD2010, '14', '1'),
])
@pytest.mark.parametrize('fmt', [aw.SaveFormat.DOCX, aw.SaveFormat.FLAT_OPC])
def test_existing_typed_settings_profile_rebuild_includes_legacy_mode_and_defaults(tmp_path, version, mode, override, fmt):
    doc = document('<w:compatSetting xmlns:o="urn:owned" w:name="Owned" '
                   'w:uri="urn:owned" w:val="private" o:extra="private"><o:child/></w:compatSetting>')
    held = doc.compatibility_options
    source = doc.light_document_model.source_settings
    held.optimize_for(version)
    for i in range(2):
        path = tmp_path / str(i)
        with pytest.warns(aw.ContentLossWarning):
            doc.save(path, fmt)
        cold = aw.Document(path, load_options=_format(fmt))
        root = ET.fromstring(cold.light_document_model.source_settings.data)
        records = root.findall(W+'compatSetting')
        values = {n.get(W+'name'): n.get(W+'val') for n in records}
        assert values.get('compatibilityMode') == mode
        assert values['overrideTableStyleFontSizeAndJustification'] == override
        assert all(values[name] == '1' for name in ('enableOpenTypeFeatures', 'doNotFlipMirrorIndents', 'differentiateMultirowTableHeaders'))
        assert 'Owned' not in values
        assert all(n.get('{urn:owned}extra') is None and len(n) == 0 for n in records)
        assert doc.compatibility_options is held and doc.light_document_model.source_settings is source
    assert b'private' in source.data


@pytest.mark.parametrize('mode,override', [(12, '0'), (14, '0'), (11, '1'), (15, '1')])
def test_plain_save_keeps_unknown_typed_records_and_completes_defaults_by_name(mode, override):
    doc = document('<w:compatSetting w:name="Owned" w:uri="urn:owned" w:val="1"/>')
    doc.light_document_model.compatibility_mode = mode
    root = ET.fromstring(render_settings_xml(doc.light_document_model)).find(W+'compat')
    assert root.find(W+'compatSetting').get(W+'uri') == 'urn:owned'
    values = {n.get(W+'name'): n.get(W+'val') for n in root.findall(W+'compatSetting')}
    assert values['overrideTableStyleFontSizeAndJustification'] == override
    foreign = document('<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="urn:owned" w:val="0"/>')
    root = ET.fromstring(render_settings_xml(foreign.light_document_model)).find(W+'compat')
    assert len([n for n in root if n.get(W+'name') == 'enableOpenTypeFeatures']) == 1


@pytest.mark.parametrize('mode,override', [('11', '1'), ('14', '0'), ('16', '0'), (' 15 ', '1'),
    ('1.5e1', '1'), ('4294967311', '1'), ('-4294967281', '1'), ('9007199254741007', '0'),
    ('15pt', '0'), ('1_5', '0'), ('15\u00a0', '0'), ('١٥', '0')])
def test_newer_default_uses_last_named_mode_number_even_with_foreign_uri(mode, override):
    doc = document('<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/>'
                   '<w:compatSetting w:name="compatibilityMode" w:uri="urn:owned" w:val="'+mode+'"/>')
    root = ET.fromstring(render_settings_xml(doc.light_document_model)).find(W+'compat')
    modes = [n for n in root if n.get(W+'name') == 'compatibilityMode']
    assert len(modes) == 2 and modes[-1].get(W+'val') == mode
    records = {n.get(W+'name'): n.get(W+'val') for n in root.findall(W+'compatSetting')}
    assert records['overrideTableStyleFontSizeAndJustification'] == override
