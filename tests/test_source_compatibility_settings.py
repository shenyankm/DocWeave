"""Compatibility declarations retain source order instead of adding template flags."""
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as E
from defusedxml.common import DefusedXmlException
import pytest
import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.docx_writer.settings_part import render_settings_xml
from aspose.words_foss.saving import OoxmlSaveOptions
from tests.test_settings_relationship_loading import settings_package
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def document(compat):
    raw=settings_package('internal')
    out=BytesIO()
    with ZipFile(BytesIO(raw))as source,ZipFile(out,'w')as target:
        for name in source.namelist():
            data=source.read(name)
            if name=='word/settings.xml':
                data=('<w:settings xmlns:w="'+W[1:-1]+'">'+compat+'</w:settings>').encode()
            target.writestr(name,data)
    return aw.Document(BytesIO(out.getvalue())).light_document_model


def declarations(model):
    root=E.fromstring(render_settings_xml(model));compat=root.find(W+'compat')
    return None if compat is None else [(e.tag,dict(e.attrib))for e in compat]


@pytest.mark.parametrize('compat', ['', '<w:compat/>', '<w:compat><w:useFELayout/><w:growAutofit w:val="0"/><w:compatSetting w:name="unknown" w:uri="owned" w:val="a"/><w:compatSetting w:name="unknown" w:uri="owned" w:val="b"/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>'])
def test_loaded_source_json_and_two_cold_saves_keep_compatibility_declarations(compat):
    model=document(compat);before=model.model_dump_json();expected=declarations(model)
    assert expected is None or not any(e[0]==W+'useWord2002TableStyleRules'for e in expected)
    for value in [model,ldm.Document.model_validate_json(before)]:
        for fmt in [aw.SaveFormat.DOCX,aw.SaveFormat.FLAT_OPC]:
            for _ in range(2):
                raw=LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(value)
                value=aw.Document(BytesIO(raw)).light_document_model
                assert declarations(value)==expected
    assert model.model_dump_json()==before


def test_public_edits_change_only_tracked_source_settings():
    model=document('<w:compat><w:useFELayout/><w:compatSetting w:name="unknown" w:uri="owned" w:val="a"/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="13"/><w:doNotExpandShiftReturn/></w:compat>')
    original=model.source_settings
    model.compatibility_mode=15;model.do_not_expand_shift_return=False
    values=declarations(model)
    assert values[:2]==[(W+'useFELayout',{}),(W+'compatSetting',{W+'name':'unknown',W+'uri':'owned',W+'val':'a'})]
    assert values[2:]==[(W+'compatSetting',{W+'name':'compatibilityMode',W+'uri':'http://schemas.microsoft.com/office/word',W+'val':'15'}),(W+'doNotExpandShiftReturn',{W+'val':'0'})]
    assert model.source_settings is original


@pytest.mark.parametrize('data',[b'<w:settings xmlns:w="'+W[1:-1].encode()+b'"/>',b'<broken',b'<!DOCTYPE x [<!ENTITY e "bad">]><x/>'])
def test_invalid_compatibility_payload_is_rejected_at_save(data):
    with pytest.raises((ValueError, E.ParseError, DefusedXmlException)):
        render_settings_xml(ldm.Document(source_settings=ldm.SourceSettings(data=data)))


def test_programmatic_document_retains_its_separate_template_baseline():
    assert W+'useWord2002TableStyleRules' in [tag for tag,attrs in declarations(ldm.Document())]


@pytest.mark.parametrize('fmt', [aw.SaveFormat.DOCX, aw.SaveFormat.FLAT_OPC])
@pytest.mark.parametrize('failure', ['malformed', 'publication'])
def test_failed_compatibility_save_preserves_source_model_and_destination(tmp_path, monkeypatch, fmt, failure):
    model=document('<w:compat><w:useFELayout/></w:compat>')
    if failure == 'malformed':
        model.source_settings=ldm.SourceSettings(data=b'<broken')
    source=model.source_settings
    before=model.model_dump_json()
    doc=aw.Document();doc._document=model
    destination=tmp_path/'owned.bin'
    destination.write_bytes(b'owned previous output')
    if failure == 'publication':
        import os
        def fail(*args):
            raise OSError('owned publication failure')
        monkeypatch.setattr(os,'replace',fail)
    with pytest.raises((E.ParseError,OSError)):
        doc.save(destination,fmt)
    assert destination.read_bytes()==b'owned previous output'
    assert model.source_settings is source and model.model_dump_json()==before


def test_compatibility_snapshot_is_immutable_and_old_json_has_no_source():
    from pydantic import ValidationError
    source=ldm.SourceSettings(data=b'<w:compat xmlns:w="'+W[1:-1].encode()+b'"/>')
    with pytest.raises(ValidationError):
        source.data=None
    model=ldm.Document(source_settings=source)
    assert ldm.Document.model_validate_json(model.model_dump_json()).source_settings==source
    assert ldm.Document.model_validate_json('{"_type":"Document"}').source_settings is None


@pytest.mark.parametrize('fmt', [aw.SaveFormat.DOCX, aw.SaveFormat.FLAT_OPC])
def test_unknown_compatibility_extension_survives_json_and_two_saves(fmt):
    model=document('<w:compat xmlns:o="urn:owned" o:source="owned"><o:extra o:key="1">'
                   '<o:child>Owned retained</o:child></o:extra></w:compat>')
    expected=E.fromstring(model.source_settings.data)
    model=ldm.Document.model_validate_json(model.model_dump_json())
    for _ in range(2):
        raw=LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(model)
        model=aw.Document(BytesIO(raw)).light_document_model
        actual=E.fromstring(model.source_settings.data)
        assert actual.attrib==expected.attrib
        assert E.tostring(actual)==E.tostring(expected)
