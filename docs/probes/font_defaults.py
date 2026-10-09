"""Observe implicit font sizes from owned packages without guessing application defaults."""

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
        "ppr_empty": "<w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>",
        "rpr_empty": "<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>",
        "rpr_bold": "<w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr></w:rPrDefault></w:docDefaults>",
    }
    for name, value in sorted(defaults.items()):
        yield name + ".docx", helpers["document"](
            value + helpers["style"]("Derived", "paragraph", {}), '<w:pStyle w:val="Derived"/>', "")


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
        parser.error("--output is required for observations")
    import aspose.words as aw

    records = []
    with ZipFile(args.corpus) as archive:
        for name in sorted(archive.namelist()):
            data = archive.read(name)
            document = aw.Document(BytesIO(data))
            paragraph = next(node.as_paragraph() for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                             if node.get_text().strip() == "IMPORT")
            run = paragraph.runs[0].as_run()
            records.append({"input": name, "sha256": hashlib.sha256(data).hexdigest(),
                            "run_size": run.font.size, "style_size": document.styles.get_by_name("Derived").font.size})
    report = {"module": "aspose.words", "version": version("aspose-words"), "licensed": False,
              "python": platform.python_version(), "platform": platform.platform(),
              "corpus": "corpus/" + args.corpus.name,
              "corpus_sha256": hashlib.sha256(args.corpus.read_bytes()).hexdigest(), "records": records,
              "full_format_acceptance": False, "scope": "five implicit-size getter observations; rendering unverified"}
    args.output.write_text(json.dumps(report, indent=2) + "\n")


if __name__ == "__main__":
    main()
