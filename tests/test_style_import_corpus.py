"""Replay owned inputs against the installed package and independently inspect output."""

import hashlib
import json
import runpy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

import aspose.words_foss as aw

from .test_docx_dom import payloads

ROOT = Path(__file__).parents[1]
BENCHMARKS = ROOT / "docs" / "benchmarks"
REPORTS = [json.loads((BENCHMARKS / name).read_text()) for name in ("style-import-default-on.json", "paragraph-style-defaults-26.9.json")]
INSPECT = runpy.run_path(str(ROOT / "docs" / "probes" / "inspect_style_imports.py"))["inspect_document"]
FORMATS = runpy.run_path(str(ROOT / "scripts" / "verify_commercial_baseline.py"))["saved_story_formats"]
BASELINE = json.loads((BENCHMARKS / "style-import-conflicts-26.9.json").read_text())["reports"]["commercial"]


@pytest.mark.parametrize("report,row", [(report, row) for report in REPORTS for row in report["reports"]["candidate"]["records"]],
                         ids=[row["case"] for report in REPORTS for row in report["reports"]["candidate"]["records"]])
def test_style_import_saved_layers_and_atomic_refusals(report, row):
    checks = {item["case"]: item for item in report["independent_checks"]["candidate"]}
    with ZipFile(BENCHMARKS / report["corpus"]["archive"]) as archive:
        data = {phase: archive.read(row["case"] + "/" + phase + ".docx") for phase in ("source", "destination")}
    assert all(hashlib.sha256(value).hexdigest() == row["inputs"][phase] for phase, value in data.items())
    source, target = (aw.DocxDocument(BytesIO(data[phase])) for phase in ("source", "destination"))
    node = source.body.paragraphs[0]
    if row["outcome"] == "raised":
        with pytest.raises(NotImplementedError):
            target.import_node(node, True)
        assert payloads(target.to_bytes()) == payloads(data["destination"])
    else:
        copied = target.import_node(node, True)
        assert copied.owner_document is target and copied.parent_node is None
        native = report["reports"].get("commercial", BASELINE)
        expected_warm = next(item["imported_format"] for item in native["records"] if item["case"] == row["case"])
        font = copied.runs[0].effective_font
        alignment = copied.effective_paragraph_format.alignment
        assert {"bold": font.bold, "italic": font.italic, "size": font.size,
                "alignment": {"both": "JUSTIFY"}.get(alignment, (alignment or "left").upper())} == expected_warm
        target.body.append_child(copied)
        assert INSPECT(BytesIO(target.to_bytes())) == {key: value for key, value in checks[row["case"]].items()
                                                     if key not in {"case", "output"}}
        expected = next(item["formats"] for item in report["story_rereads"]["commercial"]["records"] if item["case"] == row["case"])
        assert FORMATS(target.to_bytes()) == expected
    assert node.parent_node is source.body and payloads(source.to_bytes()) == payloads(data["source"])
