"""DOM ordinary-size fallback keeps disabled Normal metadata intact."""
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest
import dockiln as dk
from dockiln.writers.docx import LdmDocxWriter
from .test_nondefault_normal_paint import owned_model, paint

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def dom(flag, size, defaults):
    data = LdmDocxWriter().write_to_bytes(owned_model(flag, size, defaults))
    return dk.DocxDocument(BytesIO(data))


def assert_declarations(document):
    with ZipFile(BytesIO(document.to_bytes())) as archive:
        styles = ET.fromstring(archive.read('word/styles.xml'))
        normal = next(s for s in styles if s.get(W+'styleId') == 'Normal')
        assert normal.get(W+'default') != '1'
        content = ET.fromstring(archive.read('word/document.xml'))
        assert content.find('.//'+W+'pStyle') is None
        assert content.find('.//'+W+'r/'+W+'rPr/'+W+'sz') is None


@pytest.mark.parametrize('flag', ['0', 'false', 'off'])
@pytest.mark.parametrize('size', [0, 16])
@pytest.mark.parametrize('defaults', [False, True])
def test_dom_disabled_normal_size_fallback_is_live_and_preserves_source(flag, size, defaults):
    document = dom(flag, size, defaults)
    run = document.body.paragraphs[0].runs[0]
    assert run.font.size is None
    assert run.effective_font.size == size
    paint(document.to_light_document(), size)
    held = document.styles.get_by_name('Normal').font
    held.size = 20
    assert held.size == run.effective_font.size == 20
    assert_declarations(document)
    for flat in [False, True]:
        current = document
        for _ in range(2):
            data = current.to_flat_opc() if flat else current.to_bytes()
            current = dk.DocxDocument(BytesIO(data))
            assert current.body.paragraphs[0].runs[0].effective_font.size == 20
        assert_declarations(current)
        paint(current.to_light_document(), 20)
    assert run.effective_font.size == 20


@pytest.mark.parametrize('flag', ['0', 'false', 'off'])
def test_dom_failed_save_keeps_live_size_and_raw_declarations(flag, tmp_path, monkeypatch):
    document = dom(flag, 16, False)
    held = document.styles.get_by_name('Normal').font
    held.size = 20
    before = document.to_bytes()

    from pathlib import Path

    def fail_write(path, data):
        raise OSError('owned failure')

    destination = tmp_path / 'failed.docx'
    monkeypatch.setattr(Path, 'write_bytes', fail_write)
    with pytest.raises(OSError, match='owned failure'):
        document.save(destination)
    assert not destination.exists()
    assert document.to_bytes() == before
    assert held.size == document.body.paragraphs[0].runs[0].effective_font.size == 20
    assert_declarations(document)
