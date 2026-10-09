"""Style edits reach native-observed page assignment, text origins and black ink."""

import json
import runpy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from aspose.words_foss import DocxDocument
from aspose.words_foss.pdf_writer import LdmPdfWriter

ROOT = Path(__file__).parents[1]
BENCHMARKS = ROOT / "docs/benchmarks"
REPORT = json.loads((BENCHMARKS / "pagination-rendering-26.9.json").read_text())
PROBE = runpy.run_path(str(ROOT / "docs/probes/pagination_rendering.py"))


@pytest.mark.parametrize("row", REPORT["records"], ids=lambda row: row["output"])
def test_style_pagination_edits_match_frozen_native_pdf(row):
    with ZipFile(BENCHMARKS / REPORT["corpus"]) as archive:
        document = DocxDocument(BytesIO(archive.read(row["input"])))
    setattr(document.styles.get_by_name("Target").paragraph_format, row["property"], row["value"])
    before = document.to_bytes()
    raw = LdmPdfWriter().write_to_bytes(document.to_light_document())
    assert document.to_bytes() == before
    actual, expected = PROBE["pdf_snapshot"](raw), row["native_snapshot"]
    assert actual["page_sizes"] == expected["page_sizes"]
    assert set(actual["lines"]) == set(expected["lines"])
    for label, native in expected["lines"].items():
        result = actual["lines"][label]
        assert result["page"] == native["page"]
        if label != "BEFORE":
            assert result["origin"] == pytest.approx(native["origin"], abs=REPORT["tolerances"]["text_origin_pt"])
            assert result["advance"] == pytest.approx(native["advance"], abs=REPORT["tolerances"]["text_advance_pt"])
            assert result["size"] == native["size"]
    with ZipFile(BENCHMARKS / REPORT["commercial_outputs"]) as archive:
        reference = archive.read(row["output"])
    assert max(PROBE["ink_difference"](reference, raw)) <= REPORT["tolerances"]["black_ink_difference_ratio"]
