"""Separate trial-banner insertion from observed first-paragraph pagination."""

import argparse
import hashlib
import itertools
import json
import platform
import runpy
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

PREFIXES = {
    "none": "",
    "empty": "<w:p/>",
    "text": "<w:p><w:r><w:t>BEFORE</w:t></w:r></w:p>",
    "table": '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>',
    "section": '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>BEFORE</w:t></w:r></w:p>',
}


def inputs():
    probe = runpy.run_path(str(Path(__file__).with_name("paragraph_pagination.py")))
    raw = next(data for name, _, data in probe["inputs"](first_paragraph=True)
               if name == "page_break_before/None-None-None-True.docx")
    with ZipFile(BytesIO(raw)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    for prefix, section in itertools.product(PREFIXES, ("absent", "continuous", "nextPage")):
        stream = BytesIO()
        with ZipFile(stream, "w") as archive:
            for name, data in parts.items():
                if name == "word/document.xml":
                    data = data.replace(b"<w:body>", b"<w:body>" + PREFIXES[prefix].encode(), 1)
                    if section != "absent":
                        data = data.replace(b"<w:sectPr/>", f'<w:sectPr><w:type w:val="{section}"/></w:sectPr>'.encode())
                archive.writestr(ZipInfo(name, (2020, 1, 1, 0, 0, 0)), data)
        yield f"{prefix}-{section}.docx", stream.getvalue()


def snapshot(document):
    result = {"owned": [], "trial_banner_count": 0}
    for section in document.sections:
        for node in section.as_section().body.paragraphs:
            paragraph = node.as_paragraph()
            text = paragraph.get_text()
            if "Created with an evaluation copy" in text:
                result["trial_banner_count"] += 1
            if "IMPORT" in text:
                result["owned"].append(paragraph.paragraph_format.page_break_before)
    assert len(result["owned"]) == 1
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    import aspose.words as aw
    assert version("aspose-words") == "26.9.0"
    args.output.mkdir(parents=True, exist_ok=True)
    rows = []
    with ZipFile(args.output / "inputs.zip", "w", compression=ZIP_DEFLATED) as corpus, ZipFile(args.output / "outputs.zip", "w", compression=ZIP_DEFLATED) as outputs:
        cases = [("load", name, raw) for name, raw in inputs()]
        cases += [("builder", prefix, None) for prefix in ("none", "empty", "text")]
        for mode, name, raw in cases:
            if mode == "load":
                info = ZipInfo(name, (2020, 1, 1, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                corpus.writestr(info, raw)
                document = aw.Document(BytesIO(raw))
            else:
                document = aw.Document()
                builder = aw.DocumentBuilder(document)
                if name != "none":
                    builder.writeln("BEFORE" if name == "text" else "")
                builder.paragraph_format.page_break_before = True
                builder.write("IMPORT")
            before = snapshot(document)
            stream = BytesIO()
            document.save(stream, aw.SaveFormat.DOCX)
            saved = stream.getvalue()
            output = f"{len(rows):02d}.docx"
            info = ZipInfo(output, (2020, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            outputs.writestr(info, saved)
            rows.append({"mode": mode, "input": name, "input_sha256": hashlib.sha256(raw).hexdigest() if raw else None,
                         "before_save": before, "saved_live": snapshot(document),
                         "reopened": snapshot(aw.Document(BytesIO(saved))),
                         "output": output, "output_sha256": hashlib.sha256(saved).hexdigest()})
    report = {"version": version("aspose-words"), "licensed": False, "python": platform.python_version(),
              "platform": platform.platform(), "licensed_behavior_confirmed": False, "rendering_acceptance": False,
              "status": "trial_confounded_first_paragraph_behavior; licensed control still required",
              "corpus": "corpus/first-paragraph-trial-26.9.zip", "outputs": "corpus/first-paragraph-trial-26.9-outputs.zip",
              "corpus_sha256": hashlib.sha256((args.output / "inputs.zip").read_bytes()).hexdigest(),
              "outputs_sha256": hashlib.sha256((args.output / "outputs.zip").read_bytes()).hexdigest(), "records": rows}
    (args.output / "observations.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"observations": len(rows)}))


if __name__ == "__main__":
    main()
