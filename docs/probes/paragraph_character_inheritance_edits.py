"""Owned character edits through defaults and basedOn; no SDK/layout acceptance."""

import argparse
import hashlib
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

TARGETS = ("base", "derived", "paragraph")
VALUES = (0.0, 1.235, -1.235)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    import aspose.words as aw

    assert version("aspose-words") == "26.9.0"
    properties = runpy.run_path(str(Path(__file__).with_name("paragraph_character_indents.py")))["PROPERTIES"]

    def snapshot(document):
        paragraph = next(node.as_paragraph() for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                         if "IMPORT" in node.get_text())
        nodes = {"base": document.styles.get_by_name("Base"),
                 "derived": document.styles.get_by_name("Derived"), "paragraph": paragraph}
        return nodes, {target: {prop: getattr(node.paragraph_format, prop) for prop in properties}
                       for target, node in nodes.items()}

    args.output.mkdir(parents=True, exist_ok=True)
    rows, errors = [], []
    output_path = args.output / "outputs.zip"
    with ZipFile(args.corpus) as corpus, ZipFile(output_path, "w", compression=ZIP_DEFLATED) as outputs:
        for name in corpus.namelist():
            raw = corpus.read(name)
            prop = name.split("/")[0]
            assert prop in properties[3:]
            for target in TARGETS:
                for value in VALUES:
                    document = aw.Document(BytesIO(raw))
                    nodes, loaded = snapshot(document)
                    setattr(nodes[target].paragraph_format, prop, value)
                    after_edit = snapshot(document)[1]
                    stream = BytesIO()
                    document.save(stream, aw.SaveFormat.DOCX)
                    saved = stream.getvalue()
                    output = f"{len(rows):03d}.docx"
                    outputs.writestr(output, saved)
                    rows.append({"input": name, "input_sha256": hashlib.sha256(raw).hexdigest(),
                                 "target": target, "property": prop, "value": value,
                                 "loaded": loaded, "after_edit": after_edit,
                                 "after_save_live": snapshot(document)[1],
                                 "after_reopen": snapshot(aw.Document(BytesIO(saved)))[1],
                                 "output": output, "output_sha256": hashlib.sha256(saved).hexdigest()})
                for value in (None, True, "bad"):
                    document = aw.Document(BytesIO(raw))
                    nodes, before = snapshot(document)
                    try:
                        setattr(nodes[target].paragraph_format, prop, value)
                    except TypeError:
                        error = "TypeError"
                    else:
                        raise AssertionError("Expected a native setter type error")
                    assert snapshot(document)[1] == before
                    errors.append({"input": name, "target": target, "property": prop,
                                   "value": value, "error": error})
    assert len(rows) == len(errors) == 216
    report = {"version": version("aspose-words"), "licensed": False,
              "python": platform.python_version(), "platform": platform.platform(),
              "full_format_acceptance": False, "rendering_acceptance": False, "sdk_acceptance": False,
              "corpus": "corpus/paragraph-character-reads-26.9.zip",
              "corpus_sha256": hashlib.sha256(args.corpus.read_bytes()).hexdigest(),
              "outputs": "corpus/paragraph-character-inheritance-edits-26.9.zip",
              "outputs_sha256": hashlib.sha256(output_path.read_bytes()).hexdigest(),
              "records": rows, "setter_errors": errors,
              "scope": "24 owned docDefaults/basedOn/direct inputs, Base/Derived/paragraph character setters with zero/positive/negative; six point/character getters across edit/save/reopen; no layout update, font changes, lists, tables or rendering"}
    (args.output / "report.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"records": len(rows), "setter_errors": len(errors), "sdk_acceptance": False}))


if __name__ == "__main__":
    main()
