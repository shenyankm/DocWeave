"""Owned character-indent inheritance getters; no edit or rendering acceptance."""

import argparse
import hashlib
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

PROPERTIES = ("character_unit_left_indent", "character_unit_right_indent", "character_unit_first_line_indent")
LAYERS = ((None, None, None, None), (125, None, None, None), (None, 125, None, None),
          (125, 200, None, None), (125, 200, 0, None), (125, 200, None, 0),
          (None, 125, 225, 325), (125, 200, 225, None))


def inputs():
    helper = runpy.run_path(str(Path(__file__).with_name("style_toggles.py")))
    for prop, attr in zip(PROPERTIES, ("leftChars", "rightChars", "firstLineChars")):
        for index, (default, base, derived, direct) in enumerate(LAYERS):
            def properties(value, attribute=attr):
                return "" if value is None else f'<w:pPr><w:ind w:{attribute}="{value}"/></w:pPr>'
            styles = '<w:docDefaults><w:pPrDefault>' + properties(default) + '</w:pPrDefault></w:docDefaults>'
            styles += '<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/>' + properties(base) + '</w:style>'
            styles += '<w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="Base"/>' + properties(derived) + '</w:style>'
            ppr = '<w:pStyle w:val="Derived"/>' + (f'<w:ind w:{attr}="{direct}"/>' if direct is not None else "")
            with ZipFile(BytesIO(helper["document"](styles, ppr, ""))) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            parts["word/document.xml"] = parts["word/document.xml"].replace(b"<w:body>", b"<w:body><w:p><w:r><w:t>BEFORE</w:t></w:r></w:p>", 1)
            stream = BytesIO()
            with ZipFile(stream, "w", compression=ZIP_DEFLATED) as archive:
                for name, raw in parts.items():
                    archive.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), raw)
            yield f"{prop}/{index}.docx", stream.getvalue()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    import aspose.words as aw

    assert version("aspose-words") == "26.9.0"
    args.output.mkdir(parents=True, exist_ok=True)
    corpus = args.output / "inputs.zip"
    rows = []
    with ZipFile(corpus, "w", compression=ZIP_DEFLATED) as archive:
        for name, raw in inputs():
            archive.writestr(name, raw)
            document = aw.Document(BytesIO(raw))
            paragraph = next(node.as_paragraph() for node in document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                             if "IMPORT" in node.get_text())
            targets = {"base": document.styles.get_by_name("Base"),
                       "derived": document.styles.get_by_name("Derived"), "paragraph": paragraph}
            rows.append({"input": name, "input_sha256": hashlib.sha256(raw).hexdigest(),
                         "values": {target: {prop: getattr(node.paragraph_format, prop) for prop in PROPERTIES}
                                    for target, node in targets.items()}})
    assert len(rows) == 24
    report = {"version": version("aspose-words"), "licensed": False,
              "python": platform.python_version(), "platform": platform.platform(),
              "full_format_acceptance": False, "rendering_acceptance": False, "sdk_acceptance": False,
              "corpus": "corpus/paragraph-character-reads-26.9.zip",
              "corpus_sha256": hashlib.sha256(corpus.read_bytes()).hexdigest(), "records": rows,
              "scope": "3 getters, 8 docDefaults/basedOn/direct layers including explicit zero; no setters, aliases, list/table contexts, layout or rendering"}
    (args.output / "report.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"records": len(rows), "sdk_acceptance": False}))


if __name__ == "__main__":
    main()
