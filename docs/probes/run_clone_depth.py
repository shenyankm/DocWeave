"""Observe both clone depths on five node types in a fixed, independently owned DOCX."""

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
    parser.add_argument("--module", choices=["aspose.words", "aspose.words_foss"], required=True)
    args = parser.parse_args()
    aw = importlib.import_module(args.module)
    commercial = args.module == "aspose.words"
    source = args.source.read_bytes()
    doc = aw.Document(BytesIO(source)) if commercial else aw.DocxDocument(BytesIO(source))
    body = doc.first_section.body if commercial else doc.body
    # Select our fixture marker; the trial may prepend its own paragraph.
    paragraph = next(node for node in body.paragraphs if (
        node.get_text() if commercial else node.text).startswith("ALPHA"))
    paragraph = paragraph.as_paragraph() if commercial else paragraph
    run = paragraph.runs[0].as_run() if commercial else paragraph.runs[0]
    table = body.tables[0]
    row = table.first_row if commercial else table.rows[0]
    cell = row.first_cell if commercial else row.cells[0]
    records = []
    for kind, node in (("paragraph", paragraph), ("run", run), ("table", table), ("row", row), ("cell", cell)):
        for deep in (False, True):
            copied = node.clone(deep)
            children = None if kind == "run" else (
                copied.as_composite_node().get_child_nodes(aw.NodeType.ANY, False).count
                if commercial else len(copied.child_nodes))
            record = {"kind": kind, "deep": deep, "node_type": int(copied.node_type),
                      "detached": copied.parent_node is None,
                      "same_owner": (copied.document if commercial else copied.owner_document) == doc,
                      "content_child_count": children}
            if kind == "run":
                copied = copied.as_run() if commercial else copied
                record.update(text=copied.text, bold=copied.font.bold)
            records.append(record)
    args.output.write_text(json.dumps({"module": args.module, "source_sha256": hashlib.sha256(source).hexdigest(),
                                     "records": records}, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"module": args.module, "observations": len(records)}))


if __name__ == "__main__":
    main()
