"""Calibrate absent/empty run defaults for paragraph and character style getters."""

import argparse
import hashlib
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo


def inputs():
    helpers = runpy.run_path(str(Path(__file__).with_name("import_style_conflicts.py")))
    defaults = {
        "absent": "", "empty": "<w:docDefaults/>",
        "paragraph": "<w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>",
        "run_group": "<w:docDefaults><w:rPrDefault/></w:docDefaults>",
        "run_empty": "<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>",
        "run_bold": "<w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr></w:rPrDefault></w:docDefaults>",
        "run_complex_size": '<w:docDefaults><w:rPrDefault><w:rPr><w:szCs w:val="34"/></w:rPr></w:rPrDefault></w:docDefaults>',
        "run_size": '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>',
    }
    for kind in ("paragraph", "character"):
        for name, value in sorted(defaults.items()):
            yield kind + "-" + name + ".docx", helpers["document"](
                value + helpers["style"]("Derived", kind, {}),
                '<w:pStyle w:val="Derived"/>' if kind == "paragraph" else "",
                '<w:rStyle w:val="Derived"/>' if kind == "character" else "")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("generate", "commercial"))
    parser.add_argument("corpus", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.action == "generate":
        with ZipFile(args.corpus, "w") as archive:
            for name, data in inputs():
                archive.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), data)
        return
    if args.output is None:
        parser.error("--output is required")
    import aspose.words as aw

    assert version("aspose-words") == "26.9.0"
    records = []
    with ZipFile(args.corpus) as archive:
        for name in sorted(archive.namelist()):
            data = archive.read(name)
            document = aw.Document(BytesIO(data))
            paragraph = next(node.as_paragraph() for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                             if node.get_text().strip() == "IMPORT")
            records.append({"input": name, "sha256": hashlib.sha256(data).hexdigest(),
                            "run_size": paragraph.runs[0].as_run().font.size,
                            "style_size": document.styles.get_by_name("Derived").font.size})
    report = {"module": "aspose.words", "version": version("aspose-words"), "licensed": False,
              "python": platform.python_version(), "platform": platform.platform(),
              "corpus": "corpus/" + args.corpus.name,
              "corpus_sha256": hashlib.sha256(args.corpus.read_bytes()).hexdigest(), "records": records,
              "full_format_acceptance": False, "rendering_acceptance": False,
              "scope": "16 plain-DOCX implicit Latin-size Run/Style getters; no complex-script selection or rendering acceptance"}
    args.output.write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps([{row["input"]: [row["run_size"], row["style_size"]]} for row in records]))


if __name__ == "__main__":
    main()
