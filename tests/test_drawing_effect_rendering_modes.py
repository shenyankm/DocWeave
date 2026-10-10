"""Effect omission is an explicit PDF policy, without changing source XML."""
from io import BytesIO
import warnings

import pytest

from aspose.words_foss import Document, SaveFormat, light_document_model as ldm
from aspose.words_foss.docx_writer import LdmDocxWriter
from aspose.words_foss.pdf_writer.diagnostics import PdfContentLossWarning
from aspose.words_foss.saving import DmlEffectsRenderingMode, PdfSaveOptions
from tests.test_drawing_effect_sources import EFFECTS, effect_docx


def test_effect_mode_values_and_default():
    assert [(m.name, int(m)) for m in DmlEffectsRenderingMode] == [
        ("SIMPLIFIED", 0), ("NONE", 1), ("FINE", 2)]
    assert PdfSaveOptions().dml_effects_rendering_mode is DmlEffectsRenderingMode.SIMPLIFIED


@pytest.mark.parametrize("value", [0, 1, 2, -1, 3, None, "FINE", True])
def test_effect_mode_rejects_non_enum_without_changing_policy(value):
    options = PdfSaveOptions()
    options.dml_effects_rendering_mode = DmlEffectsRenderingMode.NONE
    with pytest.raises(TypeError, match="DmlEffectsRenderingMode"):
        options.dml_effects_rendering_mode = value
    assert options.dml_effects_rendering_mode is DmlEffectsRenderingMode.NONE


@pytest.mark.parametrize("mode", [DmlEffectsRenderingMode.SIMPLIFIED, DmlEffectsRenderingMode.FINE])
def test_unimplemented_modes_can_abort_pdf_save_atomically(tmp_path, mode):
    owner = Document(BytesIO(effect_docx(EFFECTS[0])))
    before = owner.light_document_model.model_dump_json()
    options = PdfSaveOptions()
    options.dml_effects_rendering_mode = mode
    output = tmp_path / "existing.pdf"
    output.write_bytes(b"existing")
    with warnings.catch_warnings():
        warnings.simplefilter("error", PdfContentLossWarning)
        with pytest.raises(PdfContentLossWarning, match="shape effects"):
            owner.save(output, options)
    assert output.read_bytes() == b"existing"
    assert owner.light_document_model.model_dump_json() == before


@pytest.mark.parametrize("effect", EFFECTS)
@pytest.mark.parametrize("entrypoint", ["path", "bytes"])
def test_none_renders_base_shape_and_keeps_effect_source(tmp_path, effect, entrypoint):
    fitz = pytest.importorskip("pymupdf")
    model = Document(BytesIO(effect_docx(effect))).light_document_model
    model = ldm.Document.model_validate_json(model.model_dump_json())
    owner = Document(BytesIO(LdmDocxWriter().write_to_bytes(model)))
    before = owner.light_document_model.model_dump_json()
    options = PdfSaveOptions()
    options.dml_effects_rendering_mode = DmlEffectsRenderingMode.NONE
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        output = tmp_path / "none.pdf"
        if entrypoint == "path":
            owner.save(output, options)
        else:
            output.write_bytes(owner.to_bytes(options))
    assert not any("shape effects" in str(w.message) for w in caught)
    assert owner.light_document_model.model_dump_json() == before
    control = tmp_path / "control.pdf"
    Document(BytesIO(effect_docx(""))).save(control, SaveFormat.PDF)
    actual = fitz.open(output)[0].get_pixmap()
    expected = fitz.open(control)[0].get_pixmap()
    assert actual.samples == expected.samples
