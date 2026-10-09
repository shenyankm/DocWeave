"""Independent raw style/run layers for the owned style-conflict corpus outputs."""

import argparse
import hashlib
import json
import runpy
from pathlib import Path

from docx import Document

font_values = runpy.run_path(str(Path(__file__).with_name("inspect_import_outputs.py")))["font_values"]

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def inspect_document(source):
    doc = Document(source)
    styles = []
    for style in doc.styles:
        if style.name not in {"Ancestor", "Base", "Derived", "P", "C"}:
            continue
        alignment = style.element.xpath("w:pPr/w:jc")
        styles.append({"id": style.style_id, "name": style.name, "kind": int(style.type),
                       "based_on": style.base_style.style_id if style.base_style is not None else None,
                       "font": font_values(style.font),
                       "alignment": alignment[0].get(f"{{{W}}}val") if alignment else None})
    paragraphs = []
    for paragraph in doc.paragraphs:
        if paragraph.text not in {"IMPORT", "DESTINATION"}:
            continue
        runs = []
        for run in paragraph.runs:
            reference = run._r.xpath("w:rPr/w:rStyle")
            runs.append({"text": run.text, "style_id": reference[0].get(f"{{{W}}}val") if reference else None,
                         "font": font_values(run.font)})
        paragraphs.append({"text": paragraph.text, "style_id": paragraph.style.style_id, "runs": runs})
    return {"styles": styles, "paragraphs": paragraphs}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("folder", type=Path)
    args = parser.parse_args()
    report = json.loads((args.folder / "observations.json").read_text())
    records = []
    for row in report["records"]:
        path = args.folder / row["output"]
        assert hashlib.sha256(path.read_bytes()).hexdigest() == row["output_sha256"]
        inspected = inspect_document(path)
        assert any(p["text"] == "DESTINATION" for p in inspected["paragraphs"])
        assert any(p["text"] == "IMPORT" for p in inspected["paragraphs"]) == (row["outcome"] == "returned")
        records.append({"case": row["case"], "output": row["output"], **inspected})
    (args.folder / "independent.json").write_text(json.dumps(records, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"checked_outputs": len(records)}))


if __name__ == "__main__":
    main()
