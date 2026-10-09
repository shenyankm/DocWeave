"""Edited packages must stay within the loader's limits and save atomically."""

from io import BytesIO

import pytest

from aspose.words_foss import DocxDocument, _io

from .test_docx_dom import payloads
from .test_style_serialization_projection import documents


@pytest.mark.parametrize("limit", ["part", "expanded", "archive", "entries"])
def test_serialized_output_limits_preserve_existing_destination(monkeypatch, tmp_path, limit):
    _, document = documents()
    original = document.to_bytes()
    parts = payloads(original)
    document.body.paragraphs[0].runs[0].text = "x" * 6000
    ceilings = {
        "part": ("MAX_PART_BYTES", max(map(len, parts.values()))),
        "expanded": ("MAX_EXPANDED_BYTES", sum(map(len, parts.values()))),
        "archive": ("MAX_INPUT_BYTES", len(original)),
        "entries": ("MAX_ZIP_ENTRIES", len(parts) - 1),
    }
    name, ceiling = ceilings[limit]
    monkeypatch.setattr(_io, name, ceiling)
    path = tmp_path / "existing.docx"
    path.write_bytes(b"keep")
    with pytest.raises(ValueError):
        document.to_bytes()
    with pytest.raises(ValueError):
        document.save(path)
    assert path.read_bytes() == b"keep"
    assert document.body.paragraphs[0].runs[0].text == "x" * 6000


def test_output_at_exact_limits_can_be_reopened(monkeypatch):
    _, document = documents()
    document.body.paragraphs[0].runs[0].text = "x" * 6000
    saved = document.to_bytes()
    parts = payloads(saved)
    monkeypatch.setattr(_io, "MAX_PART_BYTES", max(map(len, parts.values())))
    monkeypatch.setattr(_io, "MAX_EXPANDED_BYTES", sum(map(len, parts.values())))
    monkeypatch.setattr(_io, "MAX_INPUT_BYTES", len(saved))
    monkeypatch.setattr(_io, "MAX_ZIP_ENTRIES", len(parts))
    assert document.to_bytes() == saved
    assert DocxDocument(BytesIO(saved)).body.paragraphs[0].runs[0].text == "x" * 6000


def test_style_projection_reappearance_is_measured_at_save(monkeypatch, tmp_path):
    source, target = documents()
    target.import_node(source.body.paragraphs[0], True)
    projected = payloads(target.to_bytes())
    assert len(projected["word/styles.xml"]) < len(target._package._payloads["word/styles.xml"])
    monkeypatch.setattr(_io, "MAX_PART_BYTES", max(map(len, projected.values())))
    target.body.paragraphs[0].runs[0].font.style_id = None
    path = tmp_path / "existing.docx"
    path.write_bytes(b"keep")
    with pytest.raises(ValueError, match="part exceeds"):
        target.save(path)
    assert path.read_bytes() == b"keep"
