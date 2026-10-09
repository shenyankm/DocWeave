"""Owned finite indent boundaries: clamping, style/direct edits and reopening."""

import hashlib
import json
import runpy
import sys
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

VALUES = (
    -1e308,
    -1e100,
    -107374182.45,
    -107374182.4,
    107374182.35,
    107374182.4,
    1e100,
    1e308,
    12.375,
)


def inputs():
    probe = runpy.run_path(str(Path(__file__).with_name("paragraph_dimensions.py")))
    names = {
        f"{prop}/5.docx"
        for prop in ("left_indent", "right_indent", "first_line_indent")
    }
    selected = [row for row in probe["inputs"]() if row[0] in names]
    assert len(selected) == 3
    yield from selected


if __name__ == "__main__":
    probe = runpy.run_path(str(Path(__file__).with_name("paragraph_dimensions.py")))
    probe["main"](inputs, VALUES)
    import aspose.words as aw

    output = Path(sys.argv[1])
    report = json.loads((output / "report.json").read_text())
    with (
        ZipFile(output / "outputs.zip") as source,
        ZipFile(output / "normalized.zip", "w", ZIP_DEFLATED) as normalized,
    ):
        for row in report["records"]:
            raw = source.read(row["output"])
            with ZipFile(BytesIO(raw)) as package:
                parts = {name: package.read(name) for name in package.namelist()}
            affected = [
                name
                for name in ("word/document.xml", "word/styles.xml")
                if b'w:hanging="-2147483648"' in parts[name]
            ]
            if not affected:
                continue
            assert row["property"] == "first_line_indent"
            for name in affected:
                parts[name] = parts[name].replace(
                    b'w:hanging="-2147483648"', b'w:hanging="2147483648"'
                )
            stream = BytesIO()
            with ZipFile(stream, "w", ZIP_DEFLATED) as package:
                for name, data in parts.items():
                    package.writestr(name, data)
            data = stream.getvalue()
            normalized.writestr(row["output"], data)
            snapshot = probe["snapshot"](aw.Document(BytesIO(data)), aw)
            assert snapshot == row["after_reopen"]
            row["normalized_sha256"] = hashlib.sha256(data).hexdigest()
            row["normalized_reopen"] = snapshot
            row["native_xml_issue"] = "negative unsigned hanging at signed twip minimum"
    report["schema_reference"] = {
        "url": "https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.indentation.hanging?view=openxml-3.0.1",
        "attribute": "w:hanging",
        "documented_numeric_type": "UInt64Value",
        "verified_on": "2026-10-10",
    }
    report["normalized_outputs"] = "normalized.zip"
    report["normalized_outputs_sha256"] = hashlib.sha256(
        (output / "normalized.zip").read_bytes()
    ).hexdigest()
    report["scope"] = (
        "finite point indent clamping and multiplication overflow; negative unsigned native hanging independently normalized and reopened; no rendering acceptance"
    )
    report["full_paragraph_formatting_acceptance"] = False
    assert len([row for row in report["records"] if "native_xml_issue" in row]) == 10
    (output / "report.json").write_text(json.dumps(report, indent=2) + "\n")
