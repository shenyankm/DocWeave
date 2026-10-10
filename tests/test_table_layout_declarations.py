"""Owned table layout declarations, effective defaults and native recovery."""
from io import BytesIO
from xml.etree import ElementTree as E
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.saving import OoxmlSaveOptions
from .test_table_cell_margins import source_docx

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def layout_input(values):
    with ZipFile(BytesIO(source_docx())) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    tree = E.fromstring(parts['word/document.xml'])
    props = tree.find('.//' + W + 'tblPr')
    for old in list(props.findall(W + 'tblLayout')):
        props.remove(old)
    index = next((i for i, child in enumerate(props) if child.tag == W + 'tblLook'), len(props))
    for value in values:
        layout = E.Element(W + 'tblLayout')
        if value is not None:
            layout.set(W + 'type', value)
        props.insert(index, layout)
        index += 1
    parts['word/document.xml'] = E.tostring(tree)
    output = BytesIO()
    with ZipFile(output, 'w') as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return output.getvalue()


def table(model):
    return model.sections[0].body.tables[0]


@pytest.mark.parametrize('values, expected', [
    ([], None), ([None], True), (['fixed'], False), (['autofit'], True),
    ([''], True), (['invalid'], True), (['Fixed'], True), (['AUTOFIT'], True),
    ([' fixed '], True), (['fixed', None], False), (['autofit', None], True),
    (['fixed', 'autofit'], True), (['autofit', 'fixed'], False),
    ([None, 'fixed'], False), (['fixed', 'invalid'], True),
])
@pytest.mark.parametrize('flat', [False, True])
def test_layout_declarations_preserve_defaults_recovery_json_and_cold_saves(values, expected, flat):
    document = aw.Document(BytesIO(layout_input(values)))
    model = document.light_document_model
    diagnostics = document.diagnostics
    assert table(model).allow_auto_fit is expected
    normalized = len(values) > 1 or any(value is not None and value not in {'fixed', 'autofit'} for value in values)
    assert bool(diagnostics) == normalized
    assert all(item.code == 'docx.normalized_table_layout' for item in diagnostics)
    assert all('invalid' not in item.message for item in diagnostics)
    before = model.model_dump_json()
    for value in (model, ldm.Document.model_validate_json(before)):
        for _ in range(2):
            raw = LdmDocxWriter(OoxmlSaveOptions(aw.SaveFormat.FLAT_OPC if flat else aw.SaveFormat.DOCX)).write_to_bytes(value)
            value = aw.Document(BytesIO(raw)).light_document_model
            assert table(value).allow_auto_fit is expected
    assert model.model_dump_json() == before


@pytest.mark.parametrize('initial', [None, False, True])
@pytest.mark.parametrize('edited', [None, False, True])
@pytest.mark.parametrize('deep', [False, True])
def test_layout_assignment_and_copy_preserve_declared_state_and_validate(initial, edited, deep):
    owner = ldm.Table(allow_auto_fit=initial)
    copied = owner.model_copy(update={'allow_auto_fit': edited}, deep=deep)
    assert copied.allow_auto_fit is edited and owner.allow_auto_fit is initial
    owner.allow_auto_fit = edited
    before = owner.model_dump_json()
    with pytest.raises(ValueError):
        owner.allow_auto_fit = 'invalid'
    assert owner.model_dump_json() == before
    with pytest.raises(ValueError):
        owner.model_copy(update={'allow_auto_fit': 'invalid'}, deep=deep)
    assert owner.model_dump_json() == before
