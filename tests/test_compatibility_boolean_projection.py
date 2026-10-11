"""Native-compatible boolean projection keeps immutable source declarations."""
from contextlib import nullcontext
from io import BytesIO
import warnings
from xml.etree import ElementTree as E

import pytest
import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.docx_writer.settings_part import _COMPAT_BOOLEAN_TAGS
from aspose.words_foss.saving import OoxmlSaveOptions
from tests.test_source_compatibility_settings import document

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
O = '{urn:owned}'


@pytest.mark.parametrize('value', ['0', 'false', 'off'])
@pytest.mark.parametrize('fmt', [aw.SaveFormat.DOCX, aw.SaveFormat.FLAT_OPC])
def test_false_compatibility_flags_are_omitted_without_mutating_source(value, fmt):
    flags = ''.join('<w:'+tag+' w:val="'+value+'"/>' for tag in _COMPAT_BOOLEAN_TAGS)
    model = document('<w:compat xmlns:o="urn:owned" o:source="kept">'+flags+
                     '<o:extra o:value="kept"/><w:compatSetting w:name="owned" '
                     'w:uri="urn:owned" w:val="kept"/></w:compat>')
    source = model.source_settings
    before = model.model_dump_json()
    current = ldm.Document.model_validate_json(before)
    for _ in range(2):
        with pytest.warns(aw.ContentLossWarning) if _ == 0 else nullcontext():
            raw = LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(current)
        current = aw.Document(BytesIO(raw)).light_document_model
        root = E.fromstring(current.source_settings.data)
        assert all(root.find(W+tag) is None for tag in _COMPAT_BOOLEAN_TAGS)
        assert root.get(O+'source') == 'kept'
        assert root.find(O+'extra').get(O+'value') == 'kept'
        assert root.find(W+'compatSetting').get(W+'val') == 'kept'
    assert model.source_settings is source and model.model_dump_json() == before


@pytest.mark.parametrize('tag', ['doNotExpandShiftReturn', 'growAutofit', 'useFELayout'])
@pytest.mark.parametrize('values', [('0', '1'), ('1', '0')])
@pytest.mark.parametrize('fmt', [aw.SaveFormat.DOCX, aw.SaveFormat.FLAT_OPC])
def test_last_boolean_declaration_controls_saved_value_and_keeps_winner_extensions(tag, values, fmt):
    model = document('<w:compat xmlns:o="urn:owned"><w:'+tag+' w:val="'+values[0]+'"/>'
                     '<w:'+tag+' w:val="'+values[1]+'" o:winner="kept"><o:extra/></w:'+tag+
                     '><o:unrelated/></w:compat>')
    source = model.source_settings
    before = model.model_dump_json()
    if tag == 'doNotExpandShiftReturn':
        assert model.do_not_expand_shift_return == (values[-1] == '1')
    current = model
    for _ in range(2):
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', aw.ContentLossWarning)
            raw = LdmDocxWriter(OoxmlSaveOptions(fmt)).write_to_bytes(current)
        cold = aw.Document(BytesIO(raw)).light_document_model
        root = E.fromstring(cold.source_settings.data)
        nodes = root.findall(W+tag)
        assert len(nodes) == (values[-1] == '1')
        if nodes:
            assert nodes[0].get(O+'winner') == 'kept'
            assert nodes[0].find(O+'extra') is not None
        assert root.find(O+'unrelated') is not None
        if tag == 'doNotExpandShiftReturn':
            assert cold.do_not_expand_shift_return == (values[-1] == '1')
        current = cold
    assert model.source_settings is source and model.model_dump_json() == before
