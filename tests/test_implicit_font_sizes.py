"""Implicit ordinary size getters match fixed official 26.9 observations."""

import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from aspose.words_foss import DocxDocument

from .test_docx_dom import payloads

BENCHMARKS = Path(__file__).parents[1] / "docs/benchmarks"
CASES = [(report, row) for name in ("font-defaults-26.9.json", "font-default-matrix-26.9.json")
         for report in [json.loads((BENCHMARKS / name).read_text())] for row in report["records"]]


@pytest.mark.parametrize("report,row", CASES, ids=[report["corpus"] + ":" + row["input"] for report, row in CASES])
def test_implicit_size_reads_edits_and_clears_match_baseline(report, row):
    with ZipFile(BENCHMARKS / report["corpus"]) as archive:
        raw = archive.read(row["input"])
    document = DocxDocument(BytesIO(raw))
    run = document.body.paragraphs[0].runs[0]
    style = document.styles.get_by_name("Derived")
    assert run.font.size is None and style.direct_font.size is None
    assert run.effective_font.size == row["run_size"]
    assert style.font.size == row["style_size"]
    assert payloads(document.to_bytes()) == payloads(raw)
    run.font.size = 15
    style.font.size = 17.5
    assert run.effective_font.size == 15
    assert DocxDocument(BytesIO(document.to_bytes())).body.paragraphs[0].runs[0].effective_font.size == 15
    run.font.size = None
    assert run.effective_font.size == 17.5
    style.direct_font.size = None
    assert run.font.size is None and style.direct_font.size is None
    assert run.effective_font.size == row["run_size"]
    assert style.font.size == row["style_size"]
    reopened = DocxDocument(BytesIO(document.to_bytes()))
    assert reopened.body.paragraphs[0].runs[0].effective_font.size == row["run_size"]
    assert reopened.styles.get_by_name("Derived").font.size == row["style_size"]
