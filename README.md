# DocKiln

[Project website](https://shenyankm.github.io/DocKiln/)

[English](README.md) | [简体中文](README.zh-CN.md)

**A Python toolkit for document conversion, Chinese/Unicode PDF generation, structured extraction, and DOCX editing — without Microsoft Word.**

## Features

- Read DOCX, legacy DOC, and Flat OPC; write DOCX, Flat OPC, or PDF.
- Generate PDF with bundled Chinese fonts; optionally use fallback fonts and text shaping.
- Extract body text, ordered paragraphs/tables, formatting, and source stories as structured data.
- Edit supported DOCX text, formatting, tables, images, and hyperlinks while retaining original package parts.
- Use a resource-limited conversion CLI or an optional LibreOffice PDF backend.

| Format | Built-in read | Built-in write |
|---|---|---|
| DOCX | Yes | Yes |
| DOC (Word 97–2003) | Yes | No |
| Flat OPC | Yes | Yes |
| PDF | No | Yes |
| RTF | OLE2/DOC-backed files only; not standard text RTF | No |

Input formats are detected from content, including Flat OPC XML files and files without an extension. Specify the save format for Flat OPC XML output. LibreOffice can separately convert original DOC/DOCX/RTF files to PDF, including standard text RTF.

## Installation

Requires **Python 3.10–3.14**. Install from the repository's `dev` branch in a clean virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install "git+https://github.com/shenyankm/DocKiln.git@dev"
```

On Windows PowerShell, activate with `.venv\Scripts\Activate.ps1`. For reproducible deployments, replace `@dev` with a tested full commit hash.

The distribution and import names are both `dockiln`, using an independent namespace. The former `aspose.words_foss` import is no longer provided; update old code to `import dockiln as dk`. Markdown/TXT reading, writing, and their load/save options have been removed; use `get_text()` to extract plain text.

Runtime dependencies install automatically. Optional text shaping:

```bash
python -m pip install "dockiln[shaping] @ git+https://github.com/shenyankm/DocKiln.git@dev"
```

## Quick Start

### Convert a document

```python
import dockiln as dk

doc = dk.Document("report.docx")
doc.save("copy.docx", dk.SaveFormat.DOCX)
doc.save("report.pdf", dk.SaveFormat.PDF)
```

`save(path)` can infer the format from the output extension. Conversion rebuilds supported content; it is not lossless DOCX round-tripping.

### In-memory input and output

```python
from io import BytesIO
from pathlib import Path
import dockiln as dk

doc = dk.Document(BytesIO(Path("report.docx").read_bytes()))
docx_bytes = doc.to_bytes(dk.SaveFormat.DOCX)
pdf_bytes = doc.to_bytes(dk.SaveFormat.PDF)
```

`to_bytes(format_or_options)` returns in-memory output without creating files.

### Edit DOCX without rebuilding it

```python
import dockiln as dk

editable = dk.DocxDocument("report.docx")
paragraph = editable.body.paragraphs[0]
paragraph.replace_text("客户公司", "新的客户公司")
paragraph.range(0, 2).apply_font(bold=True, size=12)
editable.save("edited.docx")
```

Use `DocxDocument` when unknown package parts must remain intact. Unmodified part payloads remain byte-identical, but edited XML and the ZIP are rewritten. Complex edits outside the supported range are rejected; signed packages and Strict OOXML main documents are not supported. See the [DOCX editing guide](docs/docx-dom.md).

### Structured content and diagnostics

```python
import json
from dataclasses import asdict
import dockiln as dk

doc = dk.Document("report.docx")
print(doc.get_text())
print(json.dumps(doc.to_dict(), ensure_ascii=False))
print([asdict(item) for item in doc.diagnostics])
```

`get_text()` extracts body text, including nested tables. `to_dict()` includes ordered blocks, formatting, source locations, and extracted notes/header/footer stories. Locations are model/XML paths, not PDF coordinates. Diagnostics accumulate across operations and are not a complete loss audit; structured output can contain sensitive paths and document content.

## Command Line and LibreOffice

For resource-limited conversion on **POSIX**:

```bash
python -m dockiln.convert report.docx report.pdf --strict --timeout 60 --memory-mb 1024
```

`--strict` rejects detected content loss and missing glyphs, not every fidelity issue. Ordinary library calls do not provide process timeout or memory isolation. Windows users can use the library API.

To use an independently installed LibreOffice:

```bash
python -m dockiln.convert report.docx native.pdf --backend libreoffice
```

This backend uses original input files and system fonts, not the bundled PDF fonts. It is PDF-only and cannot be combined with `--strict`, `--fallback-font`, or `--text-shaping`.

## Limitations and Safety

- No built-in PDF reader, DOC/RTF writer, Markdown/TXT reader or writer, OCR, or full Word layout engine.
- Complex fields, revisions, floating content, notes, and section-specific headers/footers have fidelity limits. Use representative documents to verify results.
- Bundled fonts cover common Chinese and Latin text, not all Unicode. Substitution can change layout; missing glyphs warn. PDF/A and PDF/UA compliance are not implemented.
- Input/archive/image limits and hardened XML parsing reduce risk, but are **not a sandbox**. Isolate untrusted documents and native parsers; retaining original DOCX parts also retains macros, embedded objects, and external relationships.
- Main output files are published atomically. Do not share a mutable document across concurrent operations.

See [conversion and safety](docs/enhanced-conversion.md) for save options, fonts, resource limits, and backend details.

## Development and Testing

```bash
git clone --branch dev https://github.com/shenyankm/DocKiln.git
cd DocKiln
python -m pip install -e ".[dev,shaping]" build docxtpl
python -m pytest tests -q
python -m pytest examples -q --rootdir=examples -c examples/pytest.ini
```

CI runs on Linux with Python 3.10–3.14. It builds an sdist and a wheel from that sdist once, checks distribution metadata, then tests the same installed wheel outside the checkout on each Python version. Historical baseline evidence is checked in a separate job; distributions and JUnit reports are retained for seven days.

Native LibreOffice and optional ecosystem tests skip when their dependencies are absent. Current results are available in [GitHub Actions](https://github.com/shenyankm/DocKiln/actions); passing tests do not guarantee Word-identical rendering or Windows/macOS compatibility.

## Documentation

- [DOCX editing guide](docs/docx-dom.md) — text ranges, formatting, tables, and package preservation.
- [Conversion and safety](docs/enhanced-conversion.md) — save options, PDF fonts, diagnostics, CLI, and LibreOffice.
- [Examples](examples/) — loading, conversion, images, and optional `docxtpl` templates.
- [Ecosystem comparison](docs/ecosystem-adoption.md) and [API alignment evidence](docs/commercial-alignment-plan.md).
- [Issues](https://github.com/shenyankm/DocKiln/issues) — bugs and feature requests.

## License and Credits

Code: [MIT](LICENSE); original notices: [third-party notices](THIRD_PARTY_NOTICES.md). Bundled fonts: [SIL OFL-1.1](src/dockiln/writers/pdf/fonts/OFL.txt); see [font provenance](src/dockiln/writers/pdf/fonts/README.md). Optional integrations have their own licenses.

DocKiln originated from [Aspose.Words FOSS for Python](https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python) and retains the original copyright notices. It is independently maintained, not an official Aspose release; API or rendering compatibility is not guaranteed.
