# DocWeave

[Project website](https://shenyankm.github.io/DocWeave/)

[English](README.md) | [简体中文](README.zh-CN.md)

**A Python toolkit for document conversion, Chinese/Unicode PDF generation, structured extraction, and DOCX editing — without Microsoft Word.**

## Features

- Convert DOCX, legacy DOC, Markdown, and TXT to supported output formats.
- Generate PDF with bundled Chinese fonts; optionally use fallback fonts and text shaping.
- Extract body text, ordered paragraphs/tables, formatting, and source stories as structured data.
- Edit supported DOCX text, formatting, tables, images, and hyperlinks while retaining original package parts.
- Use a resource-limited conversion CLI or an optional LibreOffice PDF backend.

| Format | Built-in read | Built-in write |
|---|---|---|
| DOCX | Yes | Yes |
| DOC (Word 97–2003) | Yes | No |
| Markdown | Yes | Yes |
| TXT | Yes | Yes |
| PDF | No | Yes |
| RTF | OLE2/DOC-backed files only; not standard text RTF | No |

Flat OPC is also supported; specify its load/save format explicitly. LibreOffice can separately convert original DOC/DOCX/RTF files to PDF, including standard text RTF.

## Installation

Requires **Python 3.10–3.14**. Install from the repository's `dev` branch in a clean virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install "git+https://github.com/shenyankm/DocWeave.git@dev"
```

On Windows PowerShell, activate with `.venv\Scripts\Activate.ps1`. For reproducible deployments, replace `@dev` with a tested full commit hash.

The distribution name is currently `aspose-words-foss-enhanced`, and the import is `aspose.words_foss`; there is no `docweave` import yet. Do not install alongside `aspose-words-foss`: it is a different distribution sharing the same modules, so installing or removing either can overwrite them.

Runtime dependencies install automatically. Optional text shaping:

```bash
python -m pip install "aspose-words-foss-enhanced[shaping] @ git+https://github.com/shenyankm/DocWeave.git@dev"
```

## Quick Start

### Convert a document

```python
import aspose.words_foss as aw

doc = aw.Document("report.docx")
doc.save("report.md", aw.SaveFormat.MARKDOWN)
doc.save("report.pdf", aw.SaveFormat.PDF)
doc.save("report.txt", aw.SaveFormat.TEXT)
```

`save(path)` can infer the format from the output extension. Conversion rebuilds supported content; it is not lossless DOCX round-tripping.

### Create from Markdown

```python
from io import BytesIO
import aspose.words_foss as aw

markdown = "# Report\n\n客户公司\n\n中文 **粗体** and English.\n"
doc = aw.Document(BytesIO(markdown.encode("utf-8")), aw.MarkdownLoadOptions())
doc.save("report.docx", aw.SaveFormat.DOCX)
pdf_bytes = doc.to_bytes(aw.SaveFormat.PDF)
```

Use an explicit load option for Markdown streams. `to_bytes(format_or_options)` returns in-memory output without creating files.

### Edit DOCX without rebuilding it

```python
import aspose.words_foss as aw

editable = aw.DocxDocument("report.docx")
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
import aspose.words_foss as aw

doc = aw.Document("report.docx")
print(doc.get_text())
print(json.dumps(doc.to_dict(), ensure_ascii=False))
print([asdict(item) for item in doc.diagnostics])
```

`get_text()` extracts body text, including nested tables. `to_dict()` includes ordered blocks, formatting, source locations, and extracted notes/header/footer stories. Locations are model/XML paths, not PDF coordinates. Diagnostics accumulate across operations and are not a complete loss audit; structured output can contain sensitive paths and document content.

## Command Line and LibreOffice

For resource-limited conversion on **POSIX**:

```bash
python -m aspose.words_foss.convert report.docx report.pdf --strict --timeout 60 --memory-mb 1024
```

`--strict` rejects detected content loss and missing glyphs, not every fidelity issue. Ordinary library calls do not provide process timeout or memory isolation. Windows users can use the library API.

To use an independently installed LibreOffice:

```bash
python -m aspose.words_foss.convert report.docx native.pdf --backend libreoffice
```

This backend uses original input files and system fonts, not the bundled PDF fonts. It is PDF-only and cannot be combined with `--strict`, `--fallback-font`, or `--text-shaping`.

## Limitations and Safety

- No built-in PDF reader, DOC/RTF writer, OCR, or full Word layout engine. Markdown implements selected CommonMark/GFM features.
- Complex fields, revisions, floating content, notes, and section-specific headers/footers have fidelity limits. Use representative documents to verify results.
- Bundled fonts cover common Chinese and Latin text, not all Unicode. Substitution can change layout; missing glyphs warn. PDF/A and PDF/UA compliance are not implemented.
- Input/archive/image limits and hardened XML parsing reduce risk, but are **not a sandbox**. Isolate untrusted documents and native parsers; retaining original DOCX parts also retains macros, embedded objects, and external relationships.
- Markdown does not download remote images. Local images are disabled by default; opt-in access is confined to the source document directory.
- Main output files are published atomically; external Markdown images are not a multi-file transaction. Do not share a mutable document across concurrent operations.

See [conversion and safety](docs/enhanced-conversion.md) for save options, fonts, resource limits, and backend details.

## Development and Testing

```bash
git clone --branch dev https://github.com/shenyankm/DocWeave.git
cd DocWeave
python -m pip install -e ".[dev,shaping]" build docxtpl
python -m pytest tests -q
python -m pytest ApiExamples -q --rootdir=ApiExamples -c ApiExamples/pytest.ini
```

Native LibreOffice tests skip when it is not installed. Current cross-platform results are available in [GitHub Actions](https://github.com/shenyankm/DocWeave/actions); passing tests do not guarantee Word-identical rendering.

## Documentation

- [DOCX editing guide](docs/docx-dom.md) — text ranges, formatting, tables, and package preservation.
- [Conversion and safety](docs/enhanced-conversion.md) — save options, PDF fonts, diagnostics, CLI, and LibreOffice.
- [Examples](ApiExamples/) — loading, conversion, images, and optional `docxtpl` templates.
- [Ecosystem comparison](docs/ecosystem-adoption.md) and [API alignment evidence](docs/commercial-alignment-plan.md).
- [Issues](https://github.com/shenyankm/DocWeave/issues) — bugs and feature requests.

## License and Credits

Code: [MIT](LICENSE). Bundled fonts: [SIL OFL-1.1](aspose/words_foss/pdf_writer/fonts/OFL.txt); see [font provenance](aspose/words_foss/pdf_writer/fonts/README.md). Optional integrations have their own licenses.

DocWeave originated from [Aspose.Words FOSS for Python](https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python) and retains the original copyright notices. It is independently maintained, not an official Aspose release; API or rendering compatibility is not guaranteed.
