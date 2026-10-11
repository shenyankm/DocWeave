"""Observed 26.9 public setter values and DOCX/Flat OPC round trips."""
from io import BytesIO
import pytest
import dockiln as dk


@pytest.mark.parametrize('value,expected', [
    (0, 0), (.01, 0), (.025, 0), (.026, .05), (.075, .1), (.125, .1),
    (36, 36), (61.7, 61.7), (107374182.375, 107374182.35),
    (1e300, 107374182.35), (2**64-1, 107374182.35), (float('nan'), -107374182.4),
    (float('inf'), -107374182.4),
])
@pytest.mark.parametrize('fmt', [dk.SaveFormat.DOCX, dk.SaveFormat.FLAT_OPC])
def test_default_tab_stop_set_save_reload(value, expected, fmt):
    doc = dk.Document()
    assert doc.default_tab_stop == 36
    doc.default_tab_stop = value
    assert doc.default_tab_stop == expected
    assert doc.light_document_model.default_tab_stop == expected
    for _ in range(2):
        raw = doc.to_bytes(fmt)
        assert doc.default_tab_stop == expected
        doc = dk.Document(BytesIO(raw))
        assert doc.default_tab_stop == expected


@pytest.mark.parametrize('value,error', [
    (-.001, RuntimeError), (float('-inf'), RuntimeError),
    (None, TypeError), ('36', TypeError), (True, TypeError),
    (10**100, OverflowError), (2**64, OverflowError), (-2**63-1, OverflowError),
])
def test_invalid_default_tab_stop_does_not_mutate(value, error):
    doc = dk.Document()
    doc.default_tab_stop = 61.7
    with pytest.raises(error):
        doc.default_tab_stop = value
    assert doc.default_tab_stop == 61.7
