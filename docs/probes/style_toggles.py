"""Generate owned inputs and measure b/i inheritance without saving trial documents."""

import argparse
import hashlib
import importlib
import itertools
import json
import platform
from importlib.metadata import version
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/package/2006/relationships"
O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def properties(value):
    return "" if value is None else f'<w:b w:val="{int(value)}"/><w:i w:val="{int(value)}"/>'


def style(identifier, kind, value, base=None, default=False):
    return (f'<w:style w:type="{kind}" w:styleId="{identifier}"' + (' w:default="1"' if default else "") +
            f'><w:name w:val="{identifier}"/>' + (f'<w:basedOn w:val="{base}"/>' if base else "") +
            f'<w:rPr>{properties(value)}</w:rPr></w:style>')


def document(styles, pstyle, rstyle, direct=None):
    body = (f'<w:document xmlns:w="{W}"><w:body><w:p><w:pPr>{pstyle}</w:pPr>'
            f'<w:r><w:rPr>{rstyle}{properties(direct)}</w:rPr><w:t>IMPORT</w:t></w:r></w:p>'
            '<w:sectPr/></w:body></w:document>')
    parts = {
        "[Content_Types].xml": (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'),
        "_rels/.rels": f'<Relationships xmlns="{R}"><Relationship Id="d" Type="{O}officeDocument" Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels": f'<Relationships xmlns="{R}"><Relationship Id="s" Type="{O}styles" Target="styles.xml"/></Relationships>',
        "word/styles.xml": f'<w:styles xmlns:w="{W}">{styles}</w:styles>',
        "word/document.xml": body,
    }
    data = BytesIO()
    with ZipFile(data, "w") as archive:
        for name, value in parts.items():
            info = ZipInfo(name, (2020, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, value.encode())
    return data.getvalue()


def inputs():
    for values in itertools.product((None, False, True), repeat=6):
        default, pb, pc, cb, cc, direct = values
        name = "chain-" + "-".join("n" if value is None else str(int(value)) for value in values)
        styles = (f'<w:docDefaults><w:rPrDefault><w:rPr>{properties(default)}</w:rPr></w:rPrDefault></w:docDefaults>' +
                  style("PB", "paragraph", pb) + style("PC", "paragraph", pc, "PB") +
                  style("CB", "character", cb) + style("CC", "character", cc, "CB"))
        yield name, document(styles, '<w:pStyle w:val="PC"/>', '<w:rStyle w:val="CC"/>', direct)
    for values in itertools.product((False, True), repeat=4):
        default, paragraph, character, explicit = values
        name = "default-" + "-".join(str(int(value)) for value in values)
        styles = (f'<w:docDefaults><w:rPrDefault><w:rPr>{properties(default)}</w:rPr></w:rPrDefault></w:docDefaults>' +
                  style("P", "paragraph", paragraph, default=True) + style("C", "character", character, default=True))
        yield name, document(styles, "", '<w:rStyle w:val="C"/>' if explicit else "")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=["generate", "commercial", "current"])
    parser.add_argument("corpus", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.action == "generate":
        with ZipFile(args.corpus, "w") as archive:
            count = 0
            for name, data in inputs():
                archive.writestr(ZipInfo(name + ".docx", (2020, 1, 1, 0, 0, 0)), data)
                count += 1
        assert count == 745
        print(json.dumps({"inputs": count, "sha256": hashlib.sha256(args.corpus.read_bytes()).hexdigest()}))
        return
    if args.output is None:
        parser.error("--output is required for observations")
    commercial = args.action == "commercial"
    module = "aspose.words" if commercial else "aspose.words_foss"
    aw = importlib.import_module(module)
    raw = args.corpus.read_bytes()
    records = []
    with ZipFile(BytesIO(raw)) as archive:
        for name in sorted(archive.namelist()):
            data = archive.read(name)
            if commercial:
                doc = aw.Document(BytesIO(data))
                paragraphs = [node.as_paragraph() for node in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                              if node.get_text().strip() == "IMPORT"]
                assert len(paragraphs) == 1
                font = paragraphs[0].runs[0].as_run().font
            else:
                doc = aw.DocxDocument(BytesIO(data))
                assert doc.body.paragraphs[0].text == "IMPORT"
                font = doc.body.paragraphs[0].runs[0].effective_font
            records.append({"input": name, "sha256": hashlib.sha256(data).hexdigest(), "bold": font.bold, "italic": font.italic})
    assert len(records) == 745
    report = {"module": module, "version": version("aspose-words" if commercial else "aspose-words-foss-enhanced"),
              "python": platform.python_version(), "platform": platform.platform(),
              "corpus_sha256": hashlib.sha256(raw).hexdigest(), "records": records,
              "scope": "b/i getters on owned six-level and default-character-style inputs; no saved output or render acceptance",
              "licensed": False if commercial else None, "full_format_acceptance": False}
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"observations": len(records), "module": module}))


if __name__ == "__main__":
    main()
