"""Record native/target node import ownership, formatting and saved outputs."""

import argparse
import hashlib
import importlib
import json
from io import BytesIO
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--module", required=True, choices=["aspose.words", "aspose.words_foss"])
    args = parser.parse_args()
    aw = importlib.import_module(args.module)
    commercial = args.module == "aspose.words"
    inputs = {name: (args.corpus / (name + ".docx")).read_bytes() for name in ("source", "destination")}
    args.output.mkdir(parents=True, exist_ok=True)

    def load(data):
        return aw.Document(BytesIO(data)) if commercial else aw.DocxDocument(BytesIO(data))

    def body(doc):
        return doc.first_section.body if commercial else doc.body

    def owner(node):
        return node.document if commercial else node.owner_document

    records = []
    for subject in ("conflict", "derived", "normal", "numbered", "image", "table", "run"):
        for deep in (False, True):
            for mode in ("default", "USE_DESTINATION_STYLES", "KEEP_SOURCE_FORMATTING", "KEEP_DIFFERENT_STYLES"):
                source, destination = (load(inputs[name]) for name in ("source", "destination"))
                if subject == "table":
                    node = body(source).tables[0]
                else:
                    marker = "IMPORT_" + ("CONFLICT" if subject == "run" else subject.upper())
                    node = next(p for p in body(source).paragraphs if (
                        p.get_text() if commercial else p.text).startswith(marker))
                    node = node.as_paragraph() if commercial else node
                    if subject == "run":
                        node = node.runs[0].as_run() if commercial else node.runs[0]
                original_parent = node.parent_node
                name = f"{subject}-{int(deep)}-{mode}"
                try:
                    copied = destination.import_node(node, deep) if mode == "default" else destination.import_node(
                        node, deep, getattr(aw.ImportFormatMode, mode))
                    outcome = {"status": "returned", "detached": copied.parent_node is None,
                               "owner_is_destination": owner(copied) == destination,
                               "source_parent_unchanged": node.parent_node == original_parent}
                    if subject == "run":
                        parent = next(p for p in body(destination).paragraphs if (
                            p.get_text() if commercial else p.text).startswith("DESTINATION"))
                        parent = parent.as_paragraph() if commercial else parent
                    else:
                        parent = body(destination)
                    parent.append_child(copied)
                except (AttributeError, TypeError, ValueError, RuntimeError, NotImplementedError) as error:
                    outcome = {"status": "raised", "exception_type": type(error).__name__, "message": str(error)[:400]}
                path = args.output / (name + ".docx")
                destination.save(str(path))
                records.append({"subject": subject, "deep": deep, "mode": mode, "outcome": outcome,
                                "output": path.name, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()})
    result = {"module": args.module, "source_sha256": {name: hashlib.sha256(data).hexdigest() for name, data in inputs.items()},
              "scope": "seven node/resource cases, two depths, default and three explicit modes; evaluation artifacts retained",
              "records": records}
    (args.output / "observations.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"observations": len(records), "returned": sum(r["outcome"]["status"] == "returned" for r in records)}))


if __name__ == "__main__":
    main()
