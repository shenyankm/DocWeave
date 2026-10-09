"""Own style-conflict inputs and actual default-mode paragraph import observations."""

import argparse
import hashlib
import importlib
import itertools
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo

document = runpy.run_path(str(Path(__file__).with_name("style_toggles.py")))["document"]


def style(identifier, kind, properties, base=None):
    run, paragraph = "", ""
    for name, value in properties.items():
        if value is None:
            continue
        value = str(int(value)) if isinstance(value, bool) else str(value)
        element = f'<w:{name} w:val="{value}"/>'
        if name == "jc":
            paragraph += element
        else:
            run += element
    return (f'<w:style w:type="{kind}" w:styleId="{identifier}"><w:name w:val="{identifier}"/>' +
            (f'<w:basedOn w:val="{base}"/>' if base else "") + f'<w:pPr>{paragraph}</w:pPr><w:rPr>{run}</w:rPr></w:style>')


def cases():
    for kind in ("paragraph", "character"):
        for name, values in (("b", (None, False, True)), ("i", (None, False, True)), ("sz", (None, 20, 36))):
            for source, child, destination in itertools.product(values, repeat=3):
                yield kind, name, source, child, destination, False, False, None
        for source, child, destination in itertools.product((None, False, True), repeat=3):
            yield kind, "b", source, child, destination, True, False, None
        yield kind, "b", None, None, None, False, True, None
    for source, child, destination in itertools.product((None, "left", "center"), (None, "right", "both"), (None, "left", "center")):
        yield "paragraph", "jc", source, child, destination, False, False, None
    for character in (False, True):
        yield "paragraph", "b", None, None, True, True, False, character


def inputs():
    for kind, name, source, child, destination, default, ancestor, character in cases():
        key = "-".join(str(value) for value in (kind, name, source, child, destination, default, ancestor, character))
        for phase, value in (("source", source), ("destination", destination)):
            styles = ('<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/>' +
                      ('<w:b/>' if default else "") + '</w:rPr></w:rPrDefault></w:docDefaults>')
            if ancestor:
                styles += style("Ancestor", kind, {"b": phase == "source"})
            styles += style("Base", kind, {name: value}, "Ancestor" if ancestor else None)
            if phase == "source":
                styles += style("Derived", kind, {name: child}, "Base")
            if kind == "character":
                styles += style("P", "paragraph", {"b": False, "i": False})
            if character is not None:
                styles += style("C", "character", {"b": character})
            identifier = "Derived" if phase == "source" else "Base"
            pstyle = f'<w:pStyle w:val="{identifier if kind == "paragraph" else "P"}"/>'
            rstyle = f'<w:rStyle w:val="{identifier}"/>' if kind == "character" else '<w:rStyle w:val="C"/>' if character is not None else ""
            data = document(styles, pstyle, rstyle)
            if phase == "destination":
                # Keep the destination distinguishable without altering its format.
                with ZipFile(BytesIO(data)) as archive:
                    parts = {item.filename: archive.read(item) for item in archive.infolist()}
                parts["word/document.xml"] = parts["word/document.xml"].replace(b">IMPORT<", b">DESTINATION<")
                stream = BytesIO()
                with ZipFile(stream, "w") as archive:
                    for filename, payload in parts.items():
                        archive.writestr(ZipInfo(filename, (2020, 1, 1, 0, 0, 0)), payload)
                data = stream.getvalue()
            yield key, phase, data


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=["generate", "commercial", "current", "commercial-read", "commercial-read-stories", "commercial-roundtrip", "commercial-no-source-getter"])
    parser.add_argument("corpus", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.action == "generate":
        with ZipFile(args.corpus, "w") as archive:
            count = 0
            for key, phase, data in inputs():
                archive.writestr(ZipInfo(key + "/" + phase + ".docx", (2020, 1, 1, 0, 0, 0)), data)
                count += 1
        assert count == 494
        print(json.dumps({"input_pairs": count // 2, "sha256": hashlib.sha256(args.corpus.read_bytes()).hexdigest()}))
        return
    if args.output is None:
        parser.error("--output is required for observations")
    args.output.mkdir(parents=True, exist_ok=True)
    commercial = args.action != "current"
    module = "aspose.words" if commercial else "aspose.words_foss"
    aw = importlib.import_module(module)

    def load(data):
        return aw.Document(BytesIO(data)) if commercial else aw.DocxDocument(BytesIO(data))

    def paragraph(doc, label="IMPORT"):
        if commercial:
            return next(node.as_paragraph() for node in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                        if node.get_text().strip() == label)
        from aspose.words_foss.dom.styles import _style_stories

        for root in _style_stories(doc._package):
            part = next(name for name, tree in doc._package._trees.items() if tree.documentElement is root)
            story = doc.body if part == "word/document.xml" else doc.story(part)
            for node in story.paragraphs:
                if node.text == label:
                    return node
        raise ValueError("Missing owned paragraph: " + label)

    def formatting(node):
        font = node.runs[0].as_run().font if commercial else node.runs[0].effective_font
        alignment = node.paragraph_format.alignment.name if commercial else node.effective_paragraph_format.alignment
        return {"bold": font.bold, "italic": font.italic, "size": font.size, "alignment": alignment}

    records = []
    if args.action in {"commercial-read", "commercial-read-stories"}:
        candidate = json.loads((args.corpus / "observations.json").read_text())
        for row in candidate["records"]:
            if row["outcome"] != "returned" and args.action == "commercial-read":
                continue
            data = (args.corpus / row["output"]).read_bytes()
            assert hashlib.sha256(data).hexdigest() == row["output_sha256"]
            document = load(data)
            if args.action == "commercial-read-stories":
                labels = ("DESTINATION", "IMPORT") if row["outcome"] == "returned" else ("DESTINATION",)
                observed = {"formats": {label: formatting(paragraph(document, label)) for label in labels}}
            else:
                observed = {"format": formatting(paragraph(document))}
            records.append({"case": row["case"], "output_sha256": row["output_sha256"], **observed})
        report = {"module": module, "version": version("aspose-words"), "python": platform.python_version(),
                  "platform": platform.platform(), "records": records,
                  "scope": "official cold getters for owned paragraphs; no render acceptance", "licensed": False}
        (args.output / "observations.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
        print(json.dumps({"checked_candidate_outputs": len(records)}))
        return
    raw = args.corpus.read_bytes()
    with ZipFile(BytesIO(raw)) as archive:
        keys = sorted({name.split("/")[0] for name in archive.namelist()})
        assert keys and set(archive.namelist()) == {key + "/" + phase + ".docx" for key in keys for phase in ("source", "destination")}
        for index, key in enumerate(keys):
            data = {phase: archive.read(key + "/" + phase + ".docx") for phase in ("source", "destination")}
            if args.action == "commercial-roundtrip":
                for phase, payload in data.items():
                    document = load(payload)
                    label = "IMPORT" if phase == "source" else "DESTINATION"
                    before = formatting(paragraph(document, label))
                    path = args.output / (f"{index:03d}-{phase}.docx")
                    document.save(str(path))
                    records.append({"case": key, "phase": phase, "input_sha256": hashlib.sha256(payload).hexdigest(),
                                    "output": path.name, "output_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                                    "before": before, "after": formatting(paragraph(load(path.read_bytes()), label))})
                continue
            source, destination = (load(data[phase]) for phase in ("source", "destination"))
            node = paragraph(source)
            row = {"case": key, "inputs": {phase: hashlib.sha256(payload).hexdigest() for phase, payload in data.items()}}
            if args.action == "commercial-no-source-getter":
                row["no_source_getter_before_import"] = True
            else:
                row["source_format"] = formatting(node)
            before = destination.to_bytes() if not commercial else None
            try:
                copied = destination.import_node(node, True)
                body = destination.first_section.body if commercial else destination.body
                body.append_child(copied)
                row.update({"outcome": "returned", "imported_format": formatting(copied.as_paragraph() if commercial else copied)})
            except (ValueError, RuntimeError, NotImplementedError) as error:
                row.update({"outcome": "raised", "exception_type": type(error).__name__, "message": str(error)[:400]})
                if not commercial:
                    row["destination_unchanged"] = before == destination.to_bytes()
            path = args.output / (f"{index:03d}.docx")
            destination.save(str(path))
            row.update({"output": path.name, "output_sha256": hashlib.sha256(path.read_bytes()).hexdigest()})
            if row["outcome"] == "returned":
                row["reopened_format"] = formatting(paragraph(load(path.read_bytes())))
            records.append(row)
    report = {"module": module, "version": version("aspose-words" if commercial else "aspose-words-foss-enhanced"),
              "python": platform.python_version(), "platform": platform.platform(),
              "corpus_sha256": hashlib.sha256(raw).hexdigest(), "records": records,
              "scope": f"{len(keys)} new paragraph/character style conflicts; default mode, deep paragraph import; b/i/size/alignment getters",
              "licensed": False if commercial else None, "full_import_acceptance": False, "rendering_acceptance": False}
    if args.action == "commercial-roundtrip":
        report["scope"] = "official save-only roundtrips of both owned inputs; no import or rendering acceptance"
    (args.output / "observations.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"observations": len(records), "returned": sum(row.get("outcome") == "returned" for row in records)}))


if __name__ == "__main__":
    main()
