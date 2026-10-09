"""Observe live and reopened Run/Style font state around an official save."""

import argparse
import hashlib
import json
import platform
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile


def snapshot(document, aw):
    paragraphs, styles = {}, {}
    for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True):
        label = node.get_text().strip()
        if label not in {"IMPORT", "DESTINATION"}:
            continue
        assert label not in paragraphs, "Duplicate owned paragraph"
        paragraph = node.as_paragraph()
        font = paragraph.runs[0].as_run().font
        paragraphs[label] = {"bold": font.bold, "italic": font.italic, "size": font.size,
                             "alignment": paragraph.paragraph_format.alignment.name}
    for name in ("Ancestor", "Base", "Derived", "P", "C"):
        style = document.styles.get_by_name(name)
        if style is not None:
            styles[name] = {"bold": style.font.bold, "italic": style.font.italic, "size": style.font.size}
    return {"paragraphs": paragraphs, "styles": styles}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    import aspose.words as aw

    assert version("aspose-words") == "26.9.0"
    args.output.mkdir(parents=True, exist_ok=True)
    records = []
    raw = args.corpus.read_bytes()
    with ZipFile(BytesIO(raw)) as archive:
        keys = sorted({name.split("/")[0] for name in archive.namelist()})
        assert keys and set(archive.namelist()) == {key + "/" + phase + ".docx" for key in keys for phase in ("source", "destination")}
        for index, key in enumerate(keys):
            data = {phase: archive.read(key + "/" + phase + ".docx") for phase in ("source", "destination")}
            source, destination = (aw.Document(BytesIO(data[phase])) for phase in ("source", "destination"))
            node = next(node for node in source.get_child_nodes(aw.NodeType.PARAGRAPH, True) if node.get_text().strip() == "IMPORT")
            copied = destination.import_node(node, True)
            destination.first_section.body.append_child(copied)
            before = snapshot(destination, aw)
            assert set(before["paragraphs"]) == {"IMPORT", "DESTINATION"}
            path = args.output / f"{index:03d}.docx"
            destination.save(str(path))
            records.append({"case": key, "inputs": {name: hashlib.sha256(value).hexdigest() for name, value in data.items()},
                            "before_save": before, "after_save_live": snapshot(destination, aw),
                            "after_reopen": snapshot(aw.Document(str(path)), aw),
                            "output": path.name, "output_sha256": hashlib.sha256(path.read_bytes()).hexdigest()})
    report = {"module": "aspose.words", "version": version("aspose-words"), "licensed": False,
              "python": platform.python_version(), "platform": platform.platform(),
              "corpus_sha256": hashlib.sha256(raw).hexdigest(), "records": records,
              "scope": "default-mode deep paragraph import; live/cold Run and owned Style b/i/size getters",
              "full_format_acceptance": False, "rendering_acceptance": False}
    (args.output / "observations.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"observations": len(records)}))


if __name__ == "__main__":
    main()
