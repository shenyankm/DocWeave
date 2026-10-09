"""Observe selected ownership/structure behaviors with a fixed native DOCX."""

import argparse
import hashlib
import importlib
import json
from io import BytesIO
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--module", required=True, choices=["aspose.words", "aspose.words_foss"])
    args = parser.parse_args()
    aw = importlib.import_module(args.module)
    commercial = args.module == "aspose.words"
    source = args.source.read_bytes()
    args.output.mkdir(parents=True, exist_ok=True)

    def load():
        return aw.Document(BytesIO(source)) if commercial else aw.DocxDocument(BytesIO(source))

    def body(doc):
        return doc.first_section.body if commercial else doc.body

    def text(node):
        return node.get_text().rstrip("\r\n\f") if commercial else node.text

    def paragraph(node):
        return node.as_paragraph() if commercial else node

    def alpha(doc):
        return paragraph(next(node for node in body(doc).paragraphs if text(node).startswith("ALPHA")))

    def deep_copy(node):
        return paragraph(node.clone(True)) if commercial else node.clone()

    def owner(node):
        return node.document if commercial else node.owner_document

    def state(doc):
        labels = [text(node) for node in body(doc).paragraphs]
        return {"labelled_paragraphs": [label for label in labels if label.startswith(("ALPHA", "BETA", "CLONE"))]}

    def move(doc, node):
        result = body(doc).append_child(node)
        return {"returned_same_node": result == node, "parent_is_body": node.parent_node == body(doc)}

    def insert(doc, node):
        beta = next(item for item in body(doc).paragraphs if text(item) == "BETA")
        result = body(doc).insert_before(beta, node)
        return {"returned_inserted_node": result == beta}

    def clone(doc, node):
        copied = deep_copy(node)
        detached = copied.parent_node is None
        same_owner = owner(copied) == doc
        run = next(iter(copied.runs))
        if commercial:
            run = run.as_run()
        run.text = "CLONE"
        body(doc).append_child(copied)
        return {"detached_before_insert": detached, "same_owner": same_owner, "source_text": text(node), "copy_text": text(copied)}

    def shallow(doc, node):
        copied = paragraph(node.clone(False))
        children = copied.get_child_nodes(aw.NodeType.ANY, False) if commercial else copied.child_nodes
        return {"copy_text": text(copied), "child_count": len(list(children)),
                "same_owner": owner(copied) == doc, "detached": copied.parent_node is None}

    def remove(doc, node):
        result = node.remove()
        return {"returns_none": result is None, "returns_self": result == node,
                "parent_is_none": node.parent_node is None, "owner_unchanged": owner(node) == doc}

    def detached_remove(doc, node):
        copied = deep_copy(node)
        result = copied.remove()
        return {"returns_none": result is None, "detached": copied.parent_node is None, "same_owner": owner(copied) == doc}

    def foreign_append(doc, node):
        other = load()
        return {"returned": body(doc).append_child(alpha(other)) is not None}

    def foreign_import(doc, node):
        other = load()
        original = alpha(other)
        copied = paragraph(doc.import_node(original, True))
        detached = copied.parent_node is None
        source_still_attached = original.parent_node == body(other)
        body(doc).append_child(copied)
        return {"owner_is_destination": owner(copied) == doc, "detached_before_insert": detached,
                "source_still_attached": source_still_attached, "source_text": text(original)}

    scenarios = {"move_to_end": move, "insert_before": insert, "clone_deep": clone,
                 "clone_shallow": shallow, "remove": remove, "remove_detached": detached_remove,
                 "cross_document_append": foreign_append, "cross_document_import": foreign_import,
                 "illegal_parent": lambda doc, node: {"returned_node": node.append_child(aw.Paragraph(doc) if commercial else doc.create_paragraph()) is not None},
                 "cycle": lambda doc, node: {"returned_node": body(doc).append_child(body(doc)) is not None}}
    records = []
    for name, operation in scenarios.items():
        doc = load()
        node = alpha(doc)
        before = state(doc)
        try:
            result = operation(doc, node)
            outcome = {"status": "returned", "result": result}
        except (TypeError, ValueError, RuntimeError, AttributeError, NotImplementedError) as error:
            outcome = {"status": "raised", "exception_type": type(error).__name__, "message": str(error)[:400]}
        output = args.output / (name + ".docx")
        doc.save(str(output))
        records.append({"scenario": name, "before": before, "after": state(doc), "outcome": outcome,
                        "saved_docx": output.name, "saved_sha256": hashlib.sha256(output.read_bytes()).hexdigest()})
    report = {"schema": 1, "module": args.module, "source_sha256": hashlib.sha256(source).hexdigest(),
              "adapter": "commercial first_section.body/as_paragraph/document; local DocxDocument.body/owner_document; deep clone uses explicit True vs local default",
              "scope": "ten plain-DOCX scenarios, evaluation artifacts retained; no full DOM claim", "records": records}
    (args.output / "observations.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"module": args.module, "scenarios": len(records)}))


if __name__ == "__main__":
    main()
