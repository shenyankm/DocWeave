"""Owned paragraph/style character edits across font contexts; no layout acceptance."""

import argparse
import hashlib
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

VALUES = (0.0, 1.235, -1.235)
TARGETS = ("paragraph", "style")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    import aspose.words as aw

    assert version("aspose-words") == "26.9.0"
    probe = runpy.run_path(str(Path(__file__).with_name("paragraph_character_indents.py")))
    properties = probe["PROPERTIES"]

    def snapshots(document, target):
        paragraph = next(node.as_paragraph() for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                         if "IMPORT" in node.get_text())
        node = paragraph if target == "paragraph" else document.styles.get_by_name("P")
        return node, {"target": {prop: getattr(node.paragraph_format, prop) for prop in properties},
                      "paragraph": {prop: getattr(paragraph.paragraph_format, prop) for prop in properties}}

    args.output.mkdir(parents=True, exist_ok=True)
    rows, errors = [], []
    output_path = args.output / "outputs.zip"
    with ZipFile(args.corpus) as corpus, ZipFile(output_path, "w", compression=ZIP_DEFLATED) as outputs:
        for name in corpus.namelist():
            if not name.startswith("fonts/"):
                continue
            raw = corpus.read(name)
            for target in TARGETS:
                for prop in properties[3:]:
                    for value in VALUES:
                        document = aw.Document(BytesIO(raw))
                        node, loaded = snapshots(document, target)
                        setattr(node.paragraph_format, prop, value)
                        after_edit = snapshots(document, target)[1]
                        stream = BytesIO()
                        document.save(stream, aw.SaveFormat.DOCX)
                        saved = stream.getvalue()
                        output = f"{len(rows):03d}.docx"
                        outputs.writestr(output, saved)
                        rows.append({"input": name, "input_sha256": hashlib.sha256(raw).hexdigest(),
                                     "target": target, "property": prop, "value": value,
                                     "loaded": loaded, "after_edit": after_edit,
                                     "after_save_live": snapshots(document, target)[1],
                                     "after_reopen": snapshots(aw.Document(BytesIO(saved)), target)[1],
                                     "output": output, "output_sha256": hashlib.sha256(saved).hexdigest()})
                    for value in (None, True, "bad"):
                        document = aw.Document(BytesIO(raw))
                        node, before = snapshots(document, target)
                        try:
                            setattr(node.paragraph_format, prop, value)
                        except TypeError:
                            error = "TypeError"
                        else:
                            raise AssertionError("Expected a native setter type error")
                        assert snapshots(document, target)[1] == before
                        errors.append({"input": name, "target": target, "property": prop,
                                       "value": value, "error": error})
    assert len(rows) == len(errors) == 144
    report = {"version": version("aspose-words"), "licensed": False,
              "python": platform.python_version(), "platform": platform.platform(),
              "full_format_acceptance": False, "rendering_acceptance": False, "sdk_acceptance": False,
              "corpus": "corpus/paragraph-character-indents-26.9.zip",
              "corpus_sha256": hashlib.sha256(args.corpus.read_bytes()).hexdigest(),
              "outputs": "corpus/paragraph-character-setters-26.9.zip",
              "outputs_sha256": hashlib.sha256(output_path.read_bytes()).hexdigest(), "records": rows,
              "setter_errors": errors,
              "scope": "8 owned font contexts, paragraph/style targets, 3 character setters, zero/positive/negative; no layout update, inheritance-only inputs, complex scripts or rendered geometry"}
    (args.output / "report.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"records": len(rows), "setter_errors": len(errors), "sdk_acceptance": False}))


if __name__ == "__main__":
    main()
