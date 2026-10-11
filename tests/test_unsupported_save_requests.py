"""Valid but unimplemented save requests must produce actionable diagnostics."""

import warnings
from io import BytesIO
from zipfile import ZipFile

import pytest

from ._word_fixtures import word_document

import dockiln as dk
from dockiln.diagnostics import ConversionWarning, collect_diagnostics
from dockiln.writers.docx import LdmDocxWriter


def requested_options():
    options = dk.saving.OoxmlSaveOptions()
    options.zip_64_mode = dk.saving.Zip64Mode.ALWAYS
    return options


@pytest.mark.parametrize("path_output", [False, True])
def test_unsupported_request_is_recorded_even_when_warnings_ignored(tmp_path, path_output):
    doc = word_document(b"CONTENT_SENTINEL")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", ConversionWarning)
        if path_output:
            target = tmp_path / "result.output"
            doc.save(target, requested_options())
            raw = target.read_bytes()
        else:
            raw = doc.to_bytes(requested_options())
    assert raw
    assert len(doc.diagnostics) == 1
    assert doc.diagnostics[0].code == "docx.unsupported_option"
    assert "zip_64_mode" in doc.diagnostics[0].message
    assert dk.Document(BytesIO(raw)).get_text().strip() == "CONTENT_SENTINEL"
    with ZipFile(BytesIO(raw)) as archive:
        assert all(info.extract_version < 45 for info in archive.infolist())


def test_rejected_warning_keeps_original_output(tmp_path):
    doc = word_document(b"CONTENT_SENTINEL")
    target = tmp_path / "original.output"
    target.write_bytes(b"ORIGINAL")
    with warnings.catch_warnings():
        warnings.simplefilter("error", ConversionWarning)
        with pytest.raises(ConversionWarning):
            doc.save(target, requested_options())
    assert target.read_bytes() == b"ORIGINAL"
    assert list(tmp_path.iterdir()) == [target]


def test_default_options_do_not_warn():
    doc = word_document(b"CONTENT_SENTINEL")
    options = dk.saving.OoxmlSaveOptions()
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        doc.to_bytes(options)
    assert not caught and not doc.diagnostics


def test_direct_writers_share_unsupported_request_diagnostics():
    model = word_document(b"CONTENT_SENTINEL").light_document_model
    diagnostics = []
    with collect_diagnostics(diagnostics), warnings.catch_warnings():
        warnings.simplefilter("ignore", ConversionWarning)
        LdmDocxWriter(requested_options()).write_to_bytes(model)
    assert len(diagnostics) == 1 and diagnostics[0].code.endswith(".unsupported_option")
