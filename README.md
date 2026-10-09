# DocWeave — Unofficial Aspose.Words FOSS Fork

[English](README.md) | [简体中文](README.zh-CN.md)

[![Unofficial fork](https://img.shields.io/badge/status-unofficial_enhanced_fork-orange.svg)](#fork-status-and-upstream-differences) [![Code license: MIT](https://img.shields.io/badge/code_license-MIT-blue.svg)](LICENSE) [![Font license: OFL-1.1](https://img.shields.io/badge/font_license-OFL--1.1-blue.svg)](aspose/words_foss/pdf_writer/fonts/OFL.txt)

**Multi-format document conversion, Chinese/Unicode PDF, structured extraction, and bounded original-package DOCX editing — without Microsoft Word.**

DocWeave is an independently maintained enhanced fork of
[Aspose.Words FOSS for Python](https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python),
not an official Aspose release or a replacement for commercial Aspose.Words.
The repository has moved to [shenyankm/DocWeave](https://github.com/shenyankm/DocWeave).
The development branch is `dev`; the distribution and import names have **not** changed:

| Item | Current value |
|---|---|
| Distribution | `aspose-words-foss-enhanced` |
| Import | `aspose.words_foss` |
| Version | `26.7.0.post2` |
| Python | 3.10–3.14 (`>=3.10,<3.15`) |
| Code / bundled fonts | MIT / OFL-1.1 |

## Navigation

- [Fork status and upstream differences](#fork-status-and-upstream-differences)
- [At a glance](#at-a-glance)
- [Installation](#installation)
- [Quick start](#quick-start)
- [Original-package DOCX editing](#original-package-docx-editing)
- [Structured content and diagnostics](#structured-content-and-diagnostics)
- [PDF fonts and optional shaping](#pdf-fonts-and-optional-shaping)
- [Bounded CLI and LibreOffice](#bounded-cli-and-libreoffice)
- [Templates and additional examples](#templates-and-additional-examples)
- [API reference](#api-reference)
- [Scope and limitations](#scope-and-limitations)
- [Development and testing](#development-and-testing)
- [Documentation & resources](#documentation--resources)
- [License](#license)

## Fork Status and Upstream Differences

The fork started from upstream revision
[`2d2efee2787cb9e56d071d17f8d7b740dce8b784`](https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python/commit/2d2efee2787cb9e56d071d17f8d7b740dce8b784).
It adds Unicode PDF fonts/layout fixes, loss diagnostics, resource limits, original-package DOM editing,
structured source stories, and optional LibreOffice conversion. This is not a comparison with the latest upstream release.

The original authors and library are credited. Familiar `Document` / `SaveFormat` names do not guarantee
upstream or commercial feature, rendering, or behavior compatibility. This repository's implementation,
tests, and documentation describe this fork; upstream documentation is only background reference.

## At a Glance

### Formats

| Format | Built-in read | Built-in write | Boundary |
|---|---|---|---|
| DOCX | Yes | Yes | Supported Transitional OOXML content; conversion rebuilds the package, not lossless round-tripping |
| DOC | Yes | No | Legacy Word 97–2003 OLE2 input; `SaveFormat.DOC` is reserved, not a writer |
| RTF | Limited | No | OLE2/DOC-backed files only; standard text RTF is rejected by the built-in reader |
| Markdown | Yes | Yes | Selected CommonMark/GFM features, not full specification compliance |
| TXT | Yes | Yes | Body text only; output is UTF-8, with no images or original layout |
| PDF | No | Yes | Built-in renderer with bundled fonts; not Word-identical pagination |

The separate LibreOffice backend converts original DOC/DOCX/RTF files, including standard text RTF, **to PDF only**.

### Choose the right path

| Need | Entry point | What it retains |
|---|---|---|
| Convert formats or render PDF | `aw.Document` | Supported content in the Light Document Model (LDM), then reconstructed output |
| Extract JSON-safe content | `Document.to_dict()` | Ordered blocks, run formatting, model locations, note references, and extracted source stories |
| Edit DOCX while retaining unknown parts | `aw.DocxDocument` | Original OOXML package with supported, validated edits |
| Narrow literal replacement | `docx_edit.replace_text()` | Original parts; matches must fit inside individual `w:t` nodes |
| Native Office-style PDF conversion | `libreoffice.convert_to_pdf()` | Original input bytes passed to installed LibreOffice, without LDM reconstruction |

Supported inline images can be embedded in DOCX/Markdown and rendered in PDF; Markdown uses base64 unless
an external image folder is requested. TXT does not retain images. Complex positioning, fields, notes,
revisions, and section-specific headers/footers have fidelity limits even when their content can be extracted.

`PdfSaveOptions.export_document_structure` adds logical table, row, cell, and paragraph relationships,
including nested tables, merged-cell spans, and image alternate text. These structures can increase PDF size and do not establish PDF/UA compliance.
Repeated table-header copies are marked as `Artifact`, retaining the first header's semantics; plain text extractors may still return visible copies.
Headers and footers are pagination artifacts. Their explicit page breaks and keep-together settings do not advance body pages;
oversized page bands may still overflow, without automatic scaling or a promise of Word-identical layout.
Page-band bookmarks and outlines retain their first visible destination across repeated pages.
Cross-page body paragraphs, headings, quotes, and code blocks retain one logical node with page-local content references.
Body and ordinary table-cell lists use `L → LI → Lbl/LBody`, with nested lists under their parent item's body.
Body lists retain text indents on continuation lines, pages and columns; markers are positioned after any required region advance.
Ordinary cells draw markers once per item across page fragments and apply source left/right indents and hanging offsets, aligning continuation lines with the body.
Lists without a positive left indent use level-based defaults; rotated cell lists, complex numbering tabs and mixed content remain limited.

## Installation

Use a clean virtual environment and install this repository's `dev` branch:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install "git+https://github.com/shenyankm/DocWeave.git@dev"
```

On Windows PowerShell, use `.venv\Scripts\Activate.ps1` instead of `source`.
For reproducible deployments, replace `@dev` with a tested full commit hash.

**`pip install aspose-words-foss` installs upstream, not this fork.** Do not assume that a similarly named
PyPI distribution is this repository. Both projects share `aspose.words_foss`, so side-by-side installation
is unsafe: installation or removal can overwrite/delete shared modules, and pip does not detect that collision.
The repository rename does not introduce a `docweave` import or package distribution.

Verify the installed module and metadata:

```bash
python -c "import aspose.words_foss as aw; from importlib.metadata import version; print(aw.__version__, version('aspose-words-foss-enhanced'), aw.__file__)"
```

### Dependencies

Runtime requirements are installed automatically:

| Package | Minimum | Purpose |
|---|---|---|
| `olefile` | 0.46 | Legacy DOC/OLE2 input |
| `fpdf2` | 2.8.9 | PDF rendering and WOFF font support |
| `pydantic` | 2.0.0 | Typed document model |
| `defusedxml` | 0.7.1 | Hardened XML parsing |
| `Pillow` | 10.0.0 | Raster-image handling and dimension checks |

Optional dependencies are separate: `[shaping]` installs `uharfbuzz>=0.39.0`; `docxtpl` enables the template
example; LibreOffice is an externally installed application. The built-in PDF path needs none of these
and does not require Word, COM, or system-font installation.

## Quick Start

### Convert a file

```python
import aspose.words_foss as aw

doc = aw.Document("report.docx")
doc.save("report.md", aw.SaveFormat.MARKDOWN)
doc.save("report.pdf", aw.SaveFormat.PDF)
doc.save("report.txt", aw.SaveFormat.TEXT)
```

Path input uses its extension; `save(path)` can infer the output format from the destination extension.
Only DOCX, Markdown, PDF, and TXT output are implemented.

### Markdown and in-memory output

```python
from io import BytesIO
import aspose.words_foss as aw

markdown = "# Report\n\n客户公司\n\n中文 **粗体** and English.\n"
doc = aw.Document(BytesIO(markdown.encode("utf-8")), aw.MarkdownLoadOptions())
doc.save("report.docx", aw.SaveFormat.DOCX)
pdf_bytes = doc.to_bytes(aw.SaveFormat.PDF)
print(doc.get_text())
```

For streams, DOCX and OLE2/DOC can be detected from magic bytes. Markdown has no distinguishing magic;
pass `MarkdownLoadOptions` or an explicit `LoadOptions.load_format`. Standard RTF magic is recognized,
but the built-in reader still rejects that format. Pass bytes via `BytesIO`; `stream=` / `data=` are deprecated.

`to_bytes(format_or_options)` requires an explicit format and creates no output files. TXT is UTF-8;
Markdown honors its encoding option. External Markdown image files require `save()` instead.

### Save options

```python
import aspose.words_foss as aw

doc = aw.Document("report.docx")
md_options = aw.saving.MarkdownSaveOptions()
md_options.encoding = "utf-8-sig"
md_options.export_underline_formatting = True
doc.save("report.md", md_options)

docx_options = aw.saving.OoxmlSaveOptions()
docx_options.pretty_format = True
docx_options.compression_level = aw.saving.CompressionLevel.MAXIMUM
doc.save("report-rebuilt.docx", docx_options)
```

Options describe implemented behavior, not full commercial API compatibility. DOCX options affect the
**rebuilding conversion path**, not `DocxDocument` original-package saves.

## Original-Package DOCX Editing

Use `DocxDocument` instead of rebuilding a DOCX when unsupported parts must remain intact.
This self-contained example creates a simple source, edits its text, and adds a horizontally merged table:

```python
from io import BytesIO
import aspose.words_foss as aw

source = aw.Document(BytesIO("客户公司".encode("utf-8")))
editable = aw.DocxDocument(BytesIO(source.to_bytes("docx")))
paragraph = editable.body.paragraphs[0]
paragraph.replace_text("客户公司", "新的客户公司")
paragraph.range(0, 2).apply_font(bold=True, size=12)
paragraph.paragraph_format.alignment = "left"

table = editable.create_table(1, 2)
editable.body.append_child(table)
table.rows[0].cells[0].paragraphs[0].append_child(editable.create_run("left"))
table.rows[0].cells[1].paragraphs[0].append_child(editable.create_run("right"))
table.merge_cells(0, 0, 2)
editable.save("edited.docx")
```

Add an inline PNG and hyperlink to that document:

```python
from io import BytesIO
from PIL import Image
import aspose.words_foss as aw

editable = aw.DocxDocument("edited.docx")
paragraph = editable.body.paragraphs[0]
link = paragraph.add_hyperlink("Documentation", "https://example.com/docs")
link.target = "mailto:team@example.com"

image = BytesIO()
Image.new("RGB", (96, 48), "navy").save(image, format="PNG")
paragraph.add_picture(image.getvalue(), width=72, alternative_text="Company logo")
editable.save("with-resources.docx")
```

- Text ranges are paragraph-local Python Unicode code-point offsets, `[start, end)`, not grapheme clusters.
  Replacement/range editing only accepts supported plain inline structures; complex ranges are rejected.
- Direct formatting supports bold, italic, size, alignment, and existing paragraph/character style IDs.
  `effective_font` / `effective_paragraph_format` resolve a supported subset of defaults/style inheritance;
  numbering, conditional table styles, theme fonts/colors, and complex-script rules are not fully resolved.
- Picture insertion accepts PNG/JPEG bytes, streams, or paths. Dimensions are in points; one specified
  dimension preserves aspect ratio. Media are deduplicated, relationships/content types updated, and IDs allocated safely.
- New hyperlink targets are HTTP(S), mailto, or nonempty bookmark fragments. Relationship IDs are part-local;
  redirecting one shared link does not change others. No destination is fetched or validated online.
- Horizontal merging requires a supported complete grid and cannot split an existing cell. Vertical/legacy merges,
  omitted grid cells, row/column edits, splitting, and cross-document/part imports remain unsupported.
- Unmodified part payloads remain byte-identical; modified XML is reserialized and the ZIP is repacked.
  **The complete file is not byte-identical.** Signed packages and Strict OOXML main documents are rejected by this DOM.
- Body and header/footer stories can be edited within supported boundaries. A `to_light_document()` snapshot
  is independent: changing it does not update the DOM, and converting it can still lose unsupported content.

For the narrower replacement API, see `docx_edit.replace_text()` and the [DOM guide](docs/docx-dom.md).
Retaining original parts is **not sanitization**: macros, embedded objects, and external relationships are preserved.

## Structured Content and Diagnostics

```python
import json
from dataclasses import asdict
import aspose.words_foss as aw

doc = aw.Document("report.docx")
content = doc.to_dict()
print(json.dumps(content, ensure_ascii=False))
print([asdict(item) for item in doc.diagnostics])
```

For a complete supported LDM snapshot, use `doc.light_document_model.model_dump_json(by_alias=True)`
and restore it with `light_document_model.Document.model_validate_json(...)`. Image bytes in this JSON
use `{"encoding": "base64", "data": "..."}`; Python-mode `model_dump()` retains bytes, and legacy UTF-8
image strings remain readable. This snapshot differs from `to_dict()` and does not retain original OOXML parts.

`to_dict()` currently uses `schema_version: 1` with additive fields:

| Field | Meaning |
|---|---|
| `source` | Input path when loaded from a path; may contain sensitive information |
| `blocks` | Ordered body paragraphs/tables with formatting, links, image metadata, and model locations |
| Cell `blocks` | Ordered paragraphs and nested tables; legacy `paragraphs` / `tables` remain available |
| Paragraph `note_references` | Footnote/endnote kind and identifier; hidden references are marked |
| Paragraph/table `provenance` | Original DOCX XML `part_name` and zero-based element `child_path`; null for generated content |
| `source_stories` | Extracted note IDs/content and header/footer parts with section, variant, and inheritance references |
| `headers_footers` | Existing LDM header/footer representation, not a full section-aware rendering contract |
| `diagnostics` | Snapshot of accumulated `code`, `severity`, `location`, and `message` records |

Locations are model paths, **not PDF page coordinates**. `get_text()` / TXT include nested body-table text
in content order, exclude hidden runs/field instructions, and show real hyperlink labels without
misinterpreting ordinary literal link syntax. They do not append source stories, headers/footers, or images.

DOCX/PDF/TXT and default Markdown conversions report `*.notes_omitted`. Opt-in Markdown
`export_notes=True` emits visible anchored footnotes/endnotes and referenced definitions; hidden
references and unreferenced notes are excluded. Code-block references move after the block with a warning.
Markdown reports merged-cell geometry loss and nested-table flattening.

`MarkdownSaveOptions.style_map` maps exact source style names (including inherited styles) to
`Heading 1` through `Heading 6`, `Quote`, `Code`, or `Normal`, without changing the source model.
For example: `opts.style_map = {"Business Title": "Heading 2"}; opts.export_notes = True`.
Provenance describes the original parsed XML snapshot, not PDF coordinates or positions after model edits.
`diagnostics` accumulate across loading and conversions, including detected warnings before failure;
records are not cleared automatically and are not a complete loss audit. Treat known loss and missing glyphs as errors:

```python
import warnings
import aspose.words_foss as aw
from aspose.words_foss.pdf_writer import PdfMissingGlyphWarning

warnings.simplefilter("error", aw.ContentLossWarning)
warnings.simplefilter("error", PdfMissingGlyphWarning)
aw.Document("report.docx").save("checked.pdf")
```

Font-substitution warnings alone do not fail these filters. Filter source paths/text before exporting JSON
outside your trust boundary; this output is not a redacted document archive.

## PDF Fonts and Optional Shaping

The built-in renderer embeds subsets of bundled Document Sans SC fonts for common Simplified/Traditional
Chinese, Latin text, and punctuation. It uses a dedicated bold face and derived oblique faces; the four
WOFF resources total about **24.6 MiB**. Original font families, monospace metrics, and exact pagination
are not preserved. Missing glyphs warn; source-font substitution may change layout.

Source superscript/subscript flags are rendered in body, table, page-band, list, heading, quote and code
formatted runs, including shaped text. Glyph scaling/rise, decorations and links follow fpdf2 metrics;
automatic line advance reserves superscript ascent, while fixed spacing retains the source value.
Code paragraphs retain per-run size, color, bold/oblique styles, decorations and script positions;
backgrounds are painted per row after page/column breaks. Mixed-position rotated cells still report
a conversion loss. These metrics are not Word-identical; see the
[script validation](docs/benchmarks/vertical-positions.json) and
[code-run validation](docs/benchmarks/code-runs.json).
Justified rich-text and Code paragraphs stretch spaces on soft-wrapped rows, including links and
decorations. Final and explicit-break rows remain natural. Table-cell justification, CJK character
distribution and Word manual-break compatibility remain incomplete; see
[justification validation](docs/benchmarks/justification.json).

For Arabic and other complex scripts, install shaping support and deploy suitable trusted fallback fonts:

```bash
python -m pip install "aspose-words-foss-enhanced[shaping] @ git+https://github.com/shenyankm/DocWeave.git@dev"
```

Set `PdfSaveOptions.text_shaping = True` and `fallback_fonts` to trusted font paths as shown in
[conversion and safety](docs/enhanced-conversion.md). Fonts are not downloaded automatically;
fallback styling and emoji coverage are limited. Neither shaping nor structure export proves PDF/UA or PDF/A compliance.

## Bounded CLI and LibreOffice

### Built-in backend

```bash
python -m aspose.words_foss.convert report.docx report.pdf --strict --timeout 60 --memory-mb 1024
```

This single-job CLI requires **POSIX**; use the ordinary library API on Windows. It has a default 60-second
wall timeout, 1024 MiB memory threshold, and 256 MiB output-file limit. Linux also applies an address-space
limit; macOS uses an RSS watchdog that can briefly overshoot, not a hard memory quota. `--memory-mb 0`
disables memory checks. Combined subprocess stdout/stderr has a 64 MiB watchdog threshold;
checks occur at up to 100 ms intervals and can briefly overshoot. Excessive logs fail the job and
preserve existing output; returned messages remain limited to 4 KiB. This also applies to the
optional LibreOffice backend. `--strict` rejects detected loss/missing glyphs, not every possible fidelity problem.
The ordinary `Document.save()` API does not impose process timeout/memory isolation.

### Original-file LibreOffice backend

```python
from aspose.words_foss.libreoffice import convert_to_pdf

convert_to_pdf("report.docx", "native.pdf", timeout=60)
```

```bash
python -m aspose.words_foss.convert report.docx native.pdf --backend libreoffice
```

Install LibreOffice separately, with `soffice` / `libreoffice` on PATH or the standard macOS application path.
The backend passes original DOC/DOCX/RTF bytes to a private per-job profile and checks the PDF before publication.
It uses **system fonts**, not the bundled WOFF resources; fidelity still depends on LibreOffice and installed fonts.
It does not silently replace the built-in backend. `--strict`, `--fallback-font`, and `--text-shaping`
are unsupported with `--backend libreoffice`. No persistent conversion service or queue is provided.

**Resource checks and a private profile are not a sandbox.** Untrusted input needs restricted network/filesystem
access, native-parser isolation, non-privileged execution, and application-level concurrency limits.

## Templates and Additional Examples

Use the existing optional `docxtpl` integration, not a built-in template language:

```bash
python -m pip install docxtpl
python ApiExamples/template_report.py template.docx context.json report.docx report.pdf
```

Templates must be trusted. The example uses Jinja `StrictUndefined` and autoescaping; DOCX and PDF are
published separately, not as a two-file transaction. Template and output paths must be distinct.

From the repository root, a Python context factory can create images bound to the actual template.
Place `{{ logo }}` in the Word template:

```python
from docxtpl import InlineImage
from docx.shared import Mm
from ApiExamples.template_report import render_report

def context(template):
    return {"company": "Example", "logo": InlineImage(template, "logo.png", width=Mm(25))}

render_report("template.docx", context, "report.docx", "report.pdf")
```

The factory runs once per report and must return a dict. Plain dict and CLI JSON contexts still work.
Missing image files and context/render failures happen before publishing either output. A later DOCX
load or PDF conversion failure can leave the new DOCX alongside the previous PDF.

| Example in [ApiExamples/](ApiExamples/) | Purpose |
|---|---|
| `convert_document.py` | Supported format conversions using supplied fixtures |
| `loading_document.py`, `loading_markdown.py` | Paths, streams, explicit formats, and Markdown import |
| `working_with_markdown_save_options.py` | Markdown encoding, line breaks, and underline export |
| `working_with_ooxml_save_options.py` | DOCX XML formatting and compression |
| `working_with_pdf_save_options.py` | Built-in PDF options and conversion |
| `working_with_txt_save_options.py` | Body text extraction and TXT output |
| `working_with_images.py` | Supported image-containing documents; TXT remains text-only |
| `template_report.py` | Trusted DOCX templates followed by PDF conversion |

## API Reference

These are the main implemented entry points, not an exhaustive catalog of internal classes:

| API | Contract |
|---|---|
| `aw.Document(source, load_options=None)` | Path or binary stream; constructs the LDM immediately |
| `Document.save(path, format_or_options=None)` | Infers format from extension or accepts a format/save-options object; atomic main-file publication |
| `Document.to_bytes(format_or_options)` | Explicit DOCX/Markdown/PDF/TXT in-memory output |
| `Document.get_text()` / `to_dict()` | Body text / structured content extraction |
| `Document.light_document_model` | Mutable parsed model; not an OOXML-preserving DOM |
| `Document.page_count` | Estimated model count, not exact Word/PDF pagination |
| `aw.LoadOptions`, `aw.MarkdownLoadOptions` | Explicit load format/encoding; Markdown empty-line preservation and local-image opt-in |
| `aw.saving.MarkdownSaveOptions` | Encoding, paragraph breaks, tables/lists/links/HTML export modes, underline, and image handling |
| `aw.saving.OoxmlSaveOptions` | Transitional DOCX compression/XML formatting/ZIP64 settings |
| `aw.saving.PdfSaveOptions` | Implemented PDF/viewer/image/outline options, fallback fonts, and optional shaping; some compatibility fields are unused |
| `aw.DocxDocument(source)` | Retained-package DOM with `save()`, `to_bytes()`, and independent `to_light_document()` |
| `DocxDocument.body`, `story(part_name)`, `part_xml(name)` | Body, editable header/footer story, read-only part XML snapshot |
| `Paragraph.replace_text()`, `range()`, `add_picture()`, `add_hyperlink()` | Validated, bounded DOM operations described above |
| `DocxDocument.create_table()`, `Table.merge_cells()` | Simple table creation and bounded horizontal grid merge |
| `docx_edit.replace_text(source, destination, replacements)` | Literal per-`w:t` replacement; every key must occur; returns replacement count |
| `libreoffice.convert_to_pdf(source, output, timeout=60)` | Separate installed-LibreOffice original-file PDF conversion |

## Scope and Limitations

- No built-in PDF reading, DOC/RTF writing, OCR, or complete Word layout/field-computation engine.
- Conversion via LDM is not lossless. Notes are extracted separately; only opt-in Markdown exports visible anchored notes;
  comments, revisions, complex fields, content controls, math, floating content, and header/footer variants have limits.
  Supported inline images do not imply arbitrary image/shape placement or complete OOXML round-trip fidelity.
- Markdown merged grids/nested tables lose geometry; the parser implements a selected CommonMark/GFM subset.
  Tab indentation retains inherited flat `+4` behavior, not column-based tab stops.
- Strict OOXML output raises `NotImplementedError`. `ECMA376_2006` and `ISO29500_2008_TRANSITIONAL`
  use the same output; `Zip64Mode.ALWAYS` behaves like `IF_NECESSARY`, not forced ZIP64 records.
- PDF/A and PDF/UA conformance are not implemented. Explicit unused PDF fields warn:
  `embed_full_fonts`, `use_core_fonts`, `font_embedding_mode`,
  `color_mode`, `preserve_form_fields`, and `memory_optimization`.
- PDF `text_compression` controls page content streams (`NONE` / `FLATE`), independently of font/image
  compression. `page_mode` requests a viewer opening mode; the default is `USE_OUTLINES`.
  Readers may ignore this preference; it does not create outlines, layers, or attachments.
- PDF outline gaps are compacted by default; `create_missing_outline_levels=True` inserts empty
  intermediate entries. `expanded_outline_levels` accepts integers 0–9: 0 collapses all items, 1 expands
  the first PDF tree level, and so on. Readers may override the stored opening state.
- Markdown `image_resolution` is unused; `export_as_html=NON_COMPATIBLE_TABLES` behaves like `NONE`.
  Attribute presence is not evidence that an option affects output.
- Defaults: input/individual ZIP part/image data **64 MiB**, expanded DOCX **256 MiB**, **10,000** ZIP entries,
  raster images **25,000,000** pixels, and table grid/span width **1,024** columns. Unsafe/duplicate/encrypted
  ZIP entries and XML entity expansion are rejected; SVG external resources are restricted.
- Markdown does not fetch remote images. Local images are disabled by default; enabling
  `MarkdownLoadOptions.allow_local_images` requires path input and confines access to the document directory.
- Main outputs use same-directory staging and atomic replacement. New files default to `0600` permissions;
  existing permissions are preserved. External Markdown images are not a multi-file transaction, and atomic
  replacement is not a complete crash-recovery guarantee.
- Do not edit/use a single mutable document concurrently. These APIs do not sanitize macro-enabled input
  or provide a complete sandbox or all possible loss diagnostics.

Detailed restrictions: [conversion and safety](docs/enhanced-conversion.md) / [DOM guide](docs/docx-dom.md).

## Development and Testing

```bash
git clone --branch dev https://github.com/shenyankm/DocWeave.git
cd DocWeave
python -m pip install -e ".[dev,shaping]" build docxtpl
python -m pytest tests -q
python -m pytest ApiExamples -q --rootdir=ApiExamples -c ApiExamples/pytest.ini
```

The `[dev]` extra installs pytest, PyMuPDF, pypdf, python-docx, and docx2python for regression, independent
PDF checks, and DOCX content comparisons. Shaping and docxtpl are optional runtime integrations;
installing them above exercises their tests rather than skipping them. Native LibreOffice tests skip when it is absent.

CI uses **four combinations**: Linux Python **3.10/3.14**, Windows **3.14**, and macOS **3.14**. Each builds
and installs the wheel, checks imports/fonts/license/typing resources, then runs the complete regression
suite and examples once **outside the checkout**. Intermediate versions are not run on every push;
configured jobs are not proof of successful cross-platform execution.

Latest installed-wheel verification on macOS/Python 3.13.15: **1,301 regression tests and 31 API examples passed**;
installed-wheel tests, docxtpl, and available LibreOffice rendering were also exercised.
See the [current validation report](docs/ecosystem-adoption.md) for environment and boundaries.
These results do not establish Microsoft Word visual equivalence or all OS/Python combinations.

```bash
python -m build --wheel --outdir wheelhouse
python scripts/benchmark.py --repeat 3 > benchmark.json
```

For wheel verification, install it in a separate environment and run `scripts/check_wheel.py` and pytest
with `--import-mode=importlib` from outside this checkout; see [upgrade notes](docs/upgrade-notes.md).
Benchmarks record import/parse/layout/serialization/write time, RSS, PDF bytes, and pages in cold workers.
The latest small comparison does **not** demonstrate an overall speedup; some medians increased about 10%.

## Documentation & Resources

| Resource | Contents |
|---|---|
| [Chinese README](README.zh-CN.md) | Corresponding Chinese usage guide |
| [DOCX DOM guide](docs/docx-dom.md) | Ranges, formatting inheritance, resources, merges, and preservation rules |
| [Conversion and safety](docs/enhanced-conversion.md) | Fonts, diagnostics, limits, CLI, and LibreOffice |
| [Upgrade notes](docs/upgrade-notes.md) | Structured output, APIs, templates, and packaging |
| [Current validation](docs/ecosystem-adoption.md) | Latest phase acceptance, installed-wheel checks, benchmark data and remaining boundaries |
| [Earlier optimization report](docs/optimization-report.md) | Historical post2 checks and performance measurements |
| [Earlier verification report](docs/verification-report.md) | Historical post1 validation, not current coverage |
| [Issues](https://github.com/shenyankm/DocWeave/issues) | Fork-specific bugs and requests |

[Upstream source](https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python),
[upstream PyPI](https://pypi.org/project/aspose-words-foss/), and
[upstream API documentation](https://reference.aspose.org/words/python/) describe the original project,
not this fork's guarantees. [Commercial Aspose.Words](https://products.aspose.com/words/python-net/)
is a separate proprietary product, not a guaranteed drop-in replacement.

## License

Library code is under the [MIT License](LICENSE); retain the copyright and permission notices.
Bundled fonts are separately under [SIL OFL-1.1](aspose/words_foss/pdf_writer/fonts/OFL.txt), not MIT;
see [font provenance and modifications](aspose/words_foss/pdf_writer/fonts/README.md).
The software is provided without warranty. Optional integrations have their own licenses and deployment requirements.

### Reference DOCX styles

Set `OoxmlSaveOptions.reference_docx = "brand.docx"` to apply supported reference style definitions
when generating DOCX. Style display names match existing output styles; IDs and based-on links
are remapped so body references and hyperlink relationships remain valid. Existing direct
formatting takes precedence. The reference body, media, headers/footers and page setup are not imported.
Styles pass through the existing LDM reader/writer, so unsupported OOXML properties are not preserved.
This does not provide raw template copying or a lossless style import.
Run fonts now resolve character-style inheritance before direct formatting. DOCX output omits
matching inherited values so reference character styles can take effect, including links and PAGE fields.
Explicit direct values equal to the original inherited value are not distinguishable in the LDM
and may follow a replacement reference style. Complete Word toggle semantics are not guaranteed.
DOCX paragraphs without an explicit style use the XML-marked default paragraph style, including
custom defaults and their base chains. LDM `Style.is_default` preserves the marker; older models remain accepted.

Current ecosystem adoption, validation results, and remaining work: [acceptance ledger](docs/ecosystem-adoption.md).
