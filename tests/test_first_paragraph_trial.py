"""Trial observations cannot justify discarding the source's explicit pagination."""

import json
import runpy
from io import BytesIO
from pathlib import Path
from shutil import copyfile
from zipfile import ZipFile

import pytest

from aspose.words_foss import DocxDocument
from aspose.words_foss.light_document_model import NodeType

from .test_docx_dom import payloads

ROOT = Path(__file__).parents[1]
BENCHMARKS = ROOT / "docs/benchmarks"
REPORT = json.loads((BENCHMARKS / "first-paragraph-trial-26.9.json").read_text())
VERIFY = runpy.run_path(str(ROOT / "scripts/verify_commercial_baseline.py"))["verify_first_paragraph_trial"]


@pytest.mark.parametrize("row", [r for r in REPORT["records"] if r["mode"] == "load"], ids=lambda r: r["input"])
def test_preserve_explicit_flag_across_paragraph_table_and_section_positions(row):
    with ZipFile(BENCHMARKS / REPORT["corpus"]) as archive:
        raw = archive.read(row["input"])
    document = DocxDocument(BytesIO(raw))
    paragraph = next(p for p in document.get_child_nodes(NodeType.PARAGRAPH, True) if "IMPORT" in p.text)
    assert paragraph.paragraph_format.page_break_before is True
    assert paragraph.effective_paragraph_format.page_break_before is True
    saved = document.to_bytes()
    assert payloads(saved) == payloads(raw)
    reopened = DocxDocument(BytesIO(saved))
    paragraph = next(p for p in reopened.get_child_nodes(NodeType.PARAGRAPH, True) if "IMPORT" in p.text)
    assert paragraph.effective_paragraph_format.page_break_before is True


def test_trial_evidence_is_independently_verifiable():
    assert VERIFY(BENCHMARKS) == 18


@pytest.mark.parametrize("mutation", ["flag", "banner", "licensed", "missing"])
def test_trial_evidence_rejects_forged_observation(tmp_path, mutation):
    (tmp_path / "corpus").mkdir()
    for name in (REPORT["corpus"], REPORT["outputs"]):
        copyfile(BENCHMARKS / name, tmp_path / name)
    report = json.loads(json.dumps(REPORT))
    if mutation == "flag":
        report["records"][15]["saved_live"]["owned"] = [True]
    elif mutation == "banner":
        report["records"][15]["saved_live"]["trial_banner_count"] = 0
    elif mutation == "licensed":
        report["licensed_behavior_confirmed"] = True
    else:
        report["records"].pop()
    (tmp_path / "first-paragraph-trial-26.9.json").write_text(json.dumps(report))
    with pytest.raises(AssertionError):
        VERIFY(tmp_path)
