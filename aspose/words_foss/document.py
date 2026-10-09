"""
Document class.
supporting .doc, .docx, .rtf, .txt, and .md input formats.

Usage:
    import aspose.words_foss as aw

    doc = aw.Document("input.docx")
    doc.save("output.md", aw.SaveFormat.MARKDOWN)

    # From a stream (Aspose.Words-compatible):
    import io
    with io.FileIO("input.docx") as stream:
        doc = aw.Document(stream)

    # Or with options:
    opts = aw.saving.MarkdownSaveOptions()
    opts.table_content_alignment = aw.saving.TableContentAlignment.CENTER
    doc.save("output.md", opts)
"""

import warnings
from contextlib import ExitStack
from enum import IntEnum
from pathlib import Path
from typing import Optional, Union, BinaryIO

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._io import atomic_output, check_input_size, read_bounded
from aspose.words_foss._flat_opc import is_flat_opc
from aspose.words_foss.models import ConversionOptions
from aspose.words_foss.diagnostics import (
    ContentLossWarning, ConversionDiagnostic, ConversionWarning, collect_diagnostics, source_story_losses, warn,
)
from aspose.words_foss.reader_factory import create_reader
from aspose.words_foss.saving import (
    MarkdownSaveOptions,
    OoxmlSaveOptions,
    PdfSaveOptions,
)

_ZIP_MAGIC = b"PK\x03\x04"
_OLE2_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
_RTF_MAGIC = b"{\\rtf"


def _detect_format_from_bytes(header: bytes) -> str:
    """Return a file extension suffix based on magic bytes."""
    if header[:4] == _ZIP_MAGIC:
        return ".docx"
    if header[:8] == _OLE2_MAGIC:
        return ".doc"
    if header[:5] == _RTF_MAGIC:
        return ".rtf"
    return ".txt"


class LoadFormat(IntEnum):
    """Document load format constants, valued as in ``aspose.words``."""

    AUTO = 0
    DOC = 10
    DOCX = 20
    FLAT_OPC = 24
    FLAT_OPC_MACRO_ENABLED = 25
    FLAT_OPC_TEMPLATE = 26
    FLAT_OPC_TEMPLATE_MACRO_ENABLED = 27
    RTF = 30
    TEXT = 62
    MARKDOWN = 63


_LOAD_FORMAT_ALIASES = {
    "auto": LoadFormat.AUTO,
    "doc": LoadFormat.DOC,
    "docx": LoadFormat.DOCX,
    "flat_opc": LoadFormat.FLAT_OPC,
    "flat_opc_macro_enabled": LoadFormat.FLAT_OPC_MACRO_ENABLED,
    "flat_opc_template": LoadFormat.FLAT_OPC_TEMPLATE,
    "flat_opc_template_macro_enabled": LoadFormat.FLAT_OPC_TEMPLATE_MACRO_ENABLED,
    "rtf": LoadFormat.RTF,
    "text": LoadFormat.TEXT,
    "txt": LoadFormat.TEXT,
    "markdown": LoadFormat.MARKDOWN,
    "md": LoadFormat.MARKDOWN,
}


def _coerce_load_format(value: object) -> "Optional[LoadFormat]":
    """Normalise a load format to a :class:`LoadFormat`, or ``None``."""
    if isinstance(value, LoadFormat):
        return value
    if isinstance(value, str):
        return _LOAD_FORMAT_ALIASES.get(value.lower())
    if isinstance(value, int):
        try:
            return LoadFormat(value)
        except ValueError:
            return None
    return None


_LOAD_FORMAT_TO_SUFFIX = {
    LoadFormat.DOC: ".doc",
    LoadFormat.DOCX: ".docx",
    LoadFormat.FLAT_OPC: ".flatopc",
    LoadFormat.FLAT_OPC_MACRO_ENABLED: ".flatopc",
    LoadFormat.FLAT_OPC_TEMPLATE: ".flatopc",
    LoadFormat.FLAT_OPC_TEMPLATE_MACRO_ENABLED: ".flatopc",
    LoadFormat.RTF: ".rtf",
    LoadFormat.TEXT: ".txt",
    LoadFormat.MARKDOWN: ".md",
}


class LoadOptions:
    """Options for loading a document.

    Mirrors the ``Aspose.Words.Loading.LoadOptions`` surface.

    Usage::

        opts = aw.LoadOptions()
        opts.load_format = aw.LoadFormat.DOC
        doc = aw.Document(stream, opts)
    """

    def __init__(self) -> None:
        self.load_format: LoadFormat | int | str = LoadFormat.AUTO
        self.encoding: str = "utf-8"


class MarkdownLoadOptions(LoadOptions):
    """Load options specific to Markdown (.md) input.

    Mirrors ``aspose.words.loading.MarkdownLoadOptions``.

    Usage::

        opts = aw.MarkdownLoadOptions()
        opts.preserve_empty_lines = True
        doc = aw.Document("input.md", opts)
    """

    def __init__(self) -> None:
        super().__init__()
        self.load_format = LoadFormat.MARKDOWN
        self.preserve_empty_lines: bool = False
        self.allow_local_images: bool = False


class SaveFormat(IntEnum):
    """Document save format constants, valued as in ``aspose.words``."""

    DOC = 10
    DOCX = 20
    PDF = 40
    TEXT = 70
    MARKDOWN = 73


_SAVE_FORMAT_ALIASES = {
    "doc": SaveFormat.DOC,
    "docx": SaveFormat.DOCX,
    "pdf": SaveFormat.PDF,
    "text": SaveFormat.TEXT,
    "txt": SaveFormat.TEXT,
    "markdown": SaveFormat.MARKDOWN,
    "md": SaveFormat.MARKDOWN,
}


def _coerce_save_format(value: object) -> "Optional[SaveFormat]":
    """Normalise a save format to a :class:`SaveFormat`, or ``None``."""
    if isinstance(value, SaveFormat):
        return value
    if isinstance(value, str):
        return _SAVE_FORMAT_ALIASES.get(value.lower())
    if isinstance(value, int):
        try:
            return SaveFormat(value)
        except ValueError:
            return None
    return None


def _is_stream(obj: object) -> bool:
    """Return True if *obj* looks like a binary stream."""
    return hasattr(obj, "read") and hasattr(obj, "seek")


class Document:
    """
    Represents a Word document.

    Loads the file at construction time
    and populates the internal Light Document Model immediately.

    Usage::

        # From a file path:
        doc = Document("input.docx")

        # From a stream (Aspose.Words-compatible):
        doc = Document(stream)

        # With load options:
        opts = LoadOptions()
        opts.load_format = LoadFormat.DOC
        doc = Document(stream, opts)

        doc.save("output.md", SaveFormat.MARKDOWN)
    """

    def __init__(
        self,
        source: Optional[Union[str, Path, BinaryIO]] = None,
        load_options: Optional[LoadOptions] = None,
        *,
        stream: Optional[BinaryIO] = None,
        data: Optional[bytes] = None,
    ):
        """Initialize Document, loading from a file path, stream, or bytes.

        The first positional argument is a file path (``str`` / ``Path``) or a
        binary stream. ``stream=`` / ``data=`` are deprecated.
        """
        self._document: Optional[ldm.Document] = None
        self.diagnostics: list[ConversionDiagnostic] = []
        self._source_name = str(source) if isinstance(source, (str, Path)) else None

        if stream is not None or data is not None:
            name = "stream" if stream is not None else "data"
            hint = (
                "pass the stream as the first positional argument"
                if stream is not None
                else "wrap the bytes in io.BytesIO() and pass that positionally"
            )
            warnings.warn(
                f"Document({name}=...) is deprecated and has no counterpart in "
                f"aspose.words; {hint}.",
                DeprecationWarning,
                stacklevel=2,
            )

        with collect_diagnostics(self.diagnostics):
            if source is not None:
                if _is_stream(source):
                    self._load_from_stream(source, load_options)
                else:
                    self._load_from_file(Path(source), load_options)
            elif stream is not None:
                self._load_from_stream(stream, load_options)
            elif data is not None:
                self._load_from_bytes(data, load_options)

    @staticmethod
    def _apply_load_options(reader: object, load_options: Optional[LoadOptions]) -> None:
        if load_options and hasattr(reader, "allow_local_images"):
            reader.allow_local_images = getattr(load_options, "allow_local_images", False)
        if load_options and hasattr(reader, "encoding"):
            reader.encoding = load_options.encoding  # type: ignore[union-attr]
        if load_options and hasattr(load_options, "preserve_empty_lines") and hasattr(
            reader, "preserve_empty_lines"
        ):
            reader.preserve_empty_lines = load_options.preserve_empty_lines  # type: ignore[union-attr]

    def _load_from_file(
        self,
        filepath: Path,
        load_options: Optional[LoadOptions],
    ) -> None:
        """Load from a file path."""
        check_input_size(filepath.stat().st_size)
        fmt = _coerce_load_format(load_options.load_format) if load_options else None
        if fmt is not None and fmt != LoadFormat.AUTO:
            suffix = _LOAD_FORMAT_TO_SUFFIX.get(fmt, filepath.suffix.lower())
        else:
            suffix = filepath.suffix.lower()
        reader = create_reader(suffix)
        self._apply_load_options(reader, load_options)
        with filepath.open("rb") as stream:
            reader.load_bytes(read_bounded(stream))
        if hasattr(reader, "_base_dir"):
            reader._base_dir = filepath.resolve().parent
        self._document = reader.to_light_document()

    def _load_from_stream(
        self,
        stream: BinaryIO,
        load_options: Optional[LoadOptions],
    ) -> None:
        """Load from a binary stream, auto-detecting format."""
        raw = read_bounded(stream)
        suffix = self._resolve_suffix(raw, load_options)
        reader = create_reader(suffix)
        self._apply_load_options(reader, load_options)
        reader.load_bytes(raw)
        self._document = reader.to_light_document()

    def _load_from_bytes(
        self,
        data: bytes,
        load_options: Optional[LoadOptions],
    ) -> None:
        """Load from raw bytes, auto-detecting format."""
        check_input_size(len(data))
        suffix = self._resolve_suffix(data, load_options)
        reader = create_reader(suffix)
        self._apply_load_options(reader, load_options)
        reader.load_bytes(data)
        self._document = reader.to_light_document()

    @staticmethod
    def _resolve_suffix(
        data: bytes,
        load_options: Optional[LoadOptions],
    ) -> str:
        fmt = _coerce_load_format(load_options.load_format) if load_options else None
        if fmt is not None and fmt != LoadFormat.AUTO:
            return _LOAD_FORMAT_TO_SUFFIX.get(fmt, ".docx")
        if is_flat_opc(data):
            return ".flatopc"
        return _detect_format_from_bytes(data[:8])

    @property
    def light_document_model(self) -> ldm.Document:
        """Access the internal Light Document Model.

        Returns:
            The LDM Document populated during construction.

        Raises:
            ValueError: If no document was loaded.
        """
        if self._document is None:
            raise ValueError("No document loaded. Provide a filepath to Document().")
        return self._document

    @property
    def sections(self) -> "list[ldm.Section]":
        """All document sections."""
        return self.light_document_model.sections

    @property
    def first_section(self) -> "Optional[ldm.Section]":
        """The first section of the document.

        Returns ``None`` if the document has no sections.
        """
        secs = self.light_document_model.sections
        return secs[0] if secs else None

    @property
    def last_section(self) -> "Optional[ldm.Section]":
        """The last section of the document.

        Returns ``None`` if the document has no sections.
        """
        secs = self.light_document_model.sections
        return secs[-1] if secs else None

    @property
    def styles(self) -> "list[ldm.Style]":
        """All document styles."""
        return self.light_document_model.styles

    @property
    def lists(self) -> "list[ldm.DocList]":
        """All document list definitions."""
        return self.light_document_model.lists

    @property
    def page_count(self) -> int:
        """Estimated page count."""
        return self.light_document_model.page_count

    def get_child_nodes(self, node_type: int = ldm.NodeType.ANY, deep: bool = False) -> list:
        """Child nodes, optionally filtered by :class:`NodeType`."""
        return self.light_document_model.get_child_nodes(node_type, deep)

    def get_text(self) -> str:
        """Extract plain text from the loaded document.

        Returns the raw text content of the document with paragraphs
        separated by newlines. Tables are included with cell text.

        Returns:
            The plain text content of the document.
        """
        return self.light_document_model.text

    def save(
        self,
        output_path: Union[str, Path],
        save_format_or_options: Union[
            "SaveFormat", int, str, MarkdownSaveOptions, PdfSaveOptions,
            OoxmlSaveOptions, None
        ] = None,
    ) -> None:
        """Save the document to the specified format.

        Args:
            output_path: Path to save the output file.
            save_format_or_options: A :class:`SaveFormat` member (or one of the
                legacy string spellings such as ``"pdf"``), or a
                MarkdownSaveOptions, PdfSaveOptions, or OoxmlSaveOptions
                instance. If ``None``, the format is inferred from
                ``output_path``'s extension.
        """
        with collect_diagnostics(self.diagnostics):
            self._save(output_path, save_format_or_options)

    def _save(self, output_path, save_format_or_options) -> None:
        doc = self.light_document_model  # validates that a document is loaded
        output_path = Path(output_path)
        suffix = output_path.suffix.lower()

        if isinstance(save_format_or_options, MarkdownSaveOptions):
            self._save_as_markdown(output_path, doc, save_format_or_options)
            return
        if isinstance(save_format_or_options, PdfSaveOptions):
            self._save_as_pdf(output_path, doc, save_format_or_options)
            return
        if isinstance(save_format_or_options, OoxmlSaveOptions):
            self._save_as_docx(output_path, doc, save_format_or_options)
            return

        fmt = _coerce_save_format(save_format_or_options)
        if fmt is None and save_format_or_options is not None:
            raise ValueError(
                f"Unsupported save format: {save_format_or_options!r}. "
                f"Supported formats: Markdown, Text, PDF, DOCX."
            )

        if fmt == SaveFormat.MARKDOWN or (fmt is None and suffix == ".md"):
            self._save_as_markdown(output_path, doc)
        elif fmt == SaveFormat.TEXT or (fmt is None and suffix == ".txt"):
            self._save_as_text(output_path, doc)
        elif fmt == SaveFormat.PDF or (fmt is None and suffix == ".pdf"):
            self._save_as_pdf(output_path, doc)
        elif fmt == SaveFormat.DOCX or (fmt is None and suffix == ".docx"):
            self._save_as_docx(output_path, doc)
        else:
            raise ValueError(
                f"Unsupported save format: {save_format_or_options!r}. "
                f"Supported formats: Markdown, Text, PDF, DOCX."
            )

    def to_bytes(
        self,
        save_format_or_options: SaveFormat | int | str | MarkdownSaveOptions | PdfSaveOptions | OoxmlSaveOptions,
    ) -> bytes:
        """Render in memory. An explicit format is required; no files are created.

        External Markdown image files require save(path, options) instead.
        Diagnostics accumulate alongside load/save diagnostics, including on failure.
        """
        doc = self.light_document_model
        with collect_diagnostics(self.diagnostics):
            fmt = _coerce_save_format(save_format_or_options)
            if isinstance(save_format_or_options, MarkdownSaveOptions) or fmt == SaveFormat.MARKDOWN:
                md_options = save_format_or_options if isinstance(save_format_or_options, MarkdownSaveOptions) else None
                if md_options and md_options.images_folder and not md_options.export_images_as_base64:
                    raise ValueError("External Markdown images require save(path, options)")
                markdown, encoding = self._render_markdown(doc, md_options)
                return markdown.encode(encoding)
            if isinstance(save_format_or_options, PdfSaveOptions) or fmt == SaveFormat.PDF:
                from aspose.words_foss.pdf_writer import LdmPdfWriter
                pdf_options = save_format_or_options if isinstance(save_format_or_options, PdfSaveOptions) else None
                return LdmPdfWriter(pdf_options).write_to_bytes(doc)
            if isinstance(save_format_or_options, OoxmlSaveOptions) or fmt == SaveFormat.DOCX:
                from aspose.words_foss.docx_writer import LdmDocxWriter
                docx_options = save_format_or_options if isinstance(save_format_or_options, OoxmlSaveOptions) else None
                return LdmDocxWriter(docx_options).write_to_bytes(doc)
            if fmt == SaveFormat.TEXT:
                return self._render_text(doc).encode("utf-8")
            raise ValueError("Unsupported save format; supported formats: Markdown, Text, PDF, DOCX")

    def to_dict(self) -> dict:
        """Export ordered content with model locations, not inferred PDF page positions."""
        from dataclasses import asdict
        from aspose.words_foss.structure import export_document

        result = export_document(self.light_document_model, self._source_name)
        result["diagnostics"] = [asdict(item) for item in self.diagnostics]
        return result

    def _render_markdown(
        self,
        doc: ldm.Document,
        options: Optional[MarkdownSaveOptions] = None,
        output_path: Optional[Path] = None,
        *,
        _image_transactions: Optional[ExitStack] = None,
    ) -> tuple[str, str]:
        """Shared Markdown rendering for path and memory output."""
        from aspose.words_foss.md_writer import LdmMarkdownWriter

        if options is not None and _coerce_save_format(options.save_format) != SaveFormat.MARKDOWN:
            raise ValueError("MarkdownSaveOptions.save_format must be MARKDOWN")
        if options is not None and options.image_resolution != 96:
            warn("Markdown image_resolution is not applied; source image bytes are preserved",
                 ConversionWarning, code="markdown.unsupported_option")
        conversion_opts = ConversionOptions()
        encoding = "utf-8"
        if options is not None:
            if options.export_underline_formatting:
                conversion_opts.export_underline = True
            conversion_opts.table_content_alignment = options.table_content_alignment
            conversion_opts.list_export_mode = options.list_export_mode
            conversion_opts.link_export_mode = options.link_export_mode
            conversion_opts.export_as_html = options.export_as_html
            conversion_opts.empty_paragraph_export_mode = options.empty_paragraph_export_mode
            conversion_opts.export_images_as_base64 = options.export_images_as_base64
            conversion_opts.images_folder = options.images_folder
            conversion_opts.images_folder_alias = options.images_folder_alias
            conversion_opts.paragraph_break = options.paragraph_break
            conversion_opts.style_map = options.style_map
            conversion_opts.export_notes = options.export_notes
            encoding = options.encoding

        writer = LdmMarkdownWriter(conversion_opts)
        if _image_transactions is not None:
            _image_transactions.enter_context(writer._image_transaction())
        return writer.write(doc, output_path=output_path), encoding

    def _save_as_markdown(
        self, output_path: Path, doc: ldm.Document, options: Optional[MarkdownSaveOptions] = None,
    ) -> None:
        # ``newline=""`` disables Python's universal-newline translation
        # so a caller's choice of ``paragraph_break`` ("\r\n", "\n",
        # ...) round-trips verbatim instead of being doubled into
        # "\r\r\n" on Windows (where text mode would otherwise rewrite
        # every ``\n`` to ``\r\n``).
        with ExitStack() as transactions:
            markdown, encoding = self._render_markdown(doc, options, output_path, _image_transactions=transactions)
            with atomic_output(output_path) as temporary:
                temporary.write_text(markdown, encoding=encoding, newline="")

    def _render_text(self, doc):
        for code, message in source_story_losses(doc):
            warn(message, ContentLossWarning, code="text." + code)
        return doc.text

    def _save_as_text(self, output_path: Path, doc: ldm.Document) -> None:
        """Extract body text, including tables; publish only after loss checks."""
        text = self._render_text(doc)
        with atomic_output(output_path) as temporary:
            temporary.write_text(text, encoding="utf-8", newline="")

    def _save_as_pdf(
        self,
        output_path: Path,
        doc: ldm.Document,
        options: Optional[PdfSaveOptions] = None,
    ) -> None:
        """Convert the loaded LDM to PDF and save."""
        from aspose.words_foss.pdf_writer import LdmPdfWriter

        writer = LdmPdfWriter(options)
        writer.write(doc, output_path)

    def _save_as_docx(
        self,
        output_path: Path,
        doc: ldm.Document,
        options: Optional[OoxmlSaveOptions] = None,
    ) -> None:
        """Convert the loaded LDM to DOCX and save."""
        from aspose.words_foss.docx_writer import LdmDocxWriter

        writer = LdmDocxWriter(options)
        writer.write(doc, output_path)
