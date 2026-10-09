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
REPORT = json.loads((BENCHMARKS / "style-import-translated.json").read_text())
INSPECT = runpy.run_path(str(ROOT / "docs" / "probes" / "inspect_style_imports.py"))["inspect_document"]
CHECKS = {row["case"]: row for row in REPORT["independent_checks"]["candidate"]}


@pytest.mark.parametrize("row", REPORT["reports"]["candidate"]["records"], ids=lambda row: row["case"])
def test_style_import_saved_layers_and_atomic_refusals(row):
    with ZipFile(BENCHMARKS / REPORT["corpus"]["archive"]) as archive:
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
        target.body.append_child(copied)
        assert INSPECT(BytesIO(target.to_bytes())) == {key: value for key, value in CHECKS[row["case"]].items()
                                                     if key not in {"case", "output"}}
    assert node.parent_node is source.body and payloads(source.to_bytes()) == payloads(data["source"])
