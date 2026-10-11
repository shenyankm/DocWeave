"""Word document loading, structured access, DOCX saving, and PDF rendering."""

import warnings
from io import BufferedIOBase, RawIOBase
from enum import IntEnum
from pathlib import Path
from typing import Optional, Union, BinaryIO

from dockiln import light_document_model as ldm
from dockiln._io import atomic_output, check_input_size, read_bounded
from dockiln._flat_opc import MAIN_TYPES, is_flat_opc
from dockiln.diagnostics import (
    ConversionDiagnostic, collect_diagnostics,
)
from dockiln.readers.factory import create_reader
from dockiln.saving import (
    OoxmlSaveOptions,
    PdfSaveOptions,
    SaveOutputParameters,
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
    raise ValueError("Unsupported input format; expected DOC, DOCX, RTF, or Flat OPC")


class LoadFormat(IntEnum):
    """Word document load format identifiers."""

    AUTO = 0
    DOC = 10
    DOCX = 20
    FLAT_OPC = 24
    FLAT_OPC_MACRO_ENABLED = 25
    FLAT_OPC_TEMPLATE = 26
    FLAT_OPC_TEMPLATE_MACRO_ENABLED = 27
    RTF = 30


_LOAD_FORMAT_ALIASES = {
    "auto": LoadFormat.AUTO,
    "doc": LoadFormat.DOC,
    "docx": LoadFormat.DOCX,
    "flat_opc": LoadFormat.FLAT_OPC,
    "flat_opc_macro_enabled": LoadFormat.FLAT_OPC_MACRO_ENABLED,
    "flat_opc_template": LoadFormat.FLAT_OPC_TEMPLATE,
    "flat_opc_template_macro_enabled": LoadFormat.FLAT_OPC_TEMPLATE_MACRO_ENABLED,
    "rtf": LoadFormat.RTF,
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
}


class LoadOptions:
    """Options for loading a document.


    Usage::

        opts = dk.LoadOptions()
        opts.load_format = dk.LoadFormat.DOC
        doc = dk.Document(stream, opts)
    """

    def __init__(self) -> None:
        self.load_format: LoadFormat | int | str = LoadFormat.AUTO




class SaveFormat(IntEnum):
    """Word document save format identifiers."""

    DOC = 10
    DOCX = 20
    FLAT_OPC = 24
    FLAT_OPC_MACRO_ENABLED = 25
    FLAT_OPC_TEMPLATE = 26
    FLAT_OPC_TEMPLATE_MACRO_ENABLED = 27
    PDF = 40


_SAVE_FORMAT_ALIASES = {
    "doc": SaveFormat.DOC,
    "docx": SaveFormat.DOCX,
    "flat_opc": SaveFormat.FLAT_OPC,
    "flat_opc_macro_enabled": SaveFormat.FLAT_OPC_MACRO_ENABLED,
    "flat_opc_template": SaveFormat.FLAT_OPC_TEMPLATE,
    "flat_opc_template_macro_enabled": SaveFormat.FLAT_OPC_TEMPLATE_MACRO_ENABLED,
    "pdf": SaveFormat.PDF,
}


_SAVE_CONTENT_TYPES = {
    SaveFormat.DOCX: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    SaveFormat.FLAT_OPC: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    SaveFormat.FLAT_OPC_MACRO_ENABLED: "application/vnd.ms-word.document.macroEnabled.12",
    SaveFormat.FLAT_OPC_TEMPLATE: "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
    SaveFormat.FLAT_OPC_TEMPLATE_MACRO_ENABLED: "application/vnd.ms-word.template.macroEnabled.12",
    SaveFormat.PDF: "application/pdf",
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

        # From a stream:
        doc = Document(stream)

        # With load options:
        opts = LoadOptions()
        opts.load_format = LoadFormat.DOC
        doc = Document(stream, opts)

        doc.save("output.pdf", SaveFormat.PDF)
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
        if load_options is not None and _coerce_load_format(load_options.load_format) is None:
            raise ValueError(f"Unsupported load format: {load_options.load_format!r}")
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
                f"Document({name}=...) is deprecated; {hint}.",
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

        if self._document is None:
            self._document = ldm.Document(sections=[ldm.Section(
                body=ldm.Body(children=[ldm.Paragraph()]))])
            from xml.etree import ElementTree as ET
            from dockiln.settings import _source_compat
            self._document.compatibility_mode = 12
            self._document.source_settings = ldm.SourceSettings(
                data=ET.tostring(_source_compat(self._document)), compatibility_mode=12)

        from dockiln.notes import bind_note_options
        bind_note_options(self._document)


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
        with filepath.open("rb") as stream:
            reader.load_bytes(read_bounded(stream))
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
    def default_tab_stop(self) -> float:
        """Default tab interval in points, stored at twip precision."""
        return self.light_document_model.default_tab_stop

    @property
    def footnote_options(self):
        return self.light_document_model.footnote_options

    @property
    def endnote_options(self):
        return self.light_document_model.endnote_options

    @property
    def compatibility_options(self):
        """Live compatibility preferences for this document."""
        from dockiln.settings import CompatibilityOptions
        if not hasattr(self, '_compatibility_options'):
            self._compatibility_options = CompatibilityOptions._for_document(self)
        self._compatibility_options._state()
        return self._compatibility_options

    @default_tab_stop.setter
    def default_tab_stop(self, value: float) -> None:
        from dockiln.utils.tab_stops import coerce_default_tab_stop
        self.light_document_model.default_tab_stop = coerce_default_tab_stop(value)

    @property
    def font_infos(self):
        """Live document-owned font registrations and embedding declarations."""
        from dockiln.fonts import FontInfoCollection
        model = self.light_document_model
        if not hasattr(self, '_font_infos') or self._font_infos._document is not model:
            self._font_infos = FontInfoCollection(model)
        return self._font_infos

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
        output_path: Union[str, Path, BinaryIO],
        save_format_or_options: Union[
            "SaveFormat", int, str, PdfSaveOptions,
            OoxmlSaveOptions, None
        ] = None,
    ) -> SaveOutputParameters:
        """Save to a path or caller-owned binary stream and return content metadata.

        Paths infer their format from the extension. Streams require an explicit
        format. DOCX starts at position zero; other formats use the current
        position. Streams remain open and are never truncated.
        """
        with collect_diagnostics(self.diagnostics):
            if isinstance(output_path, (BufferedIOBase, RawIOBase)):
                if save_format_or_options is None:
                    raise TypeError("Saving to a stream requires an explicit format")
                if output_path.closed or not output_path.writable():
                    raise RuntimeError("Output stream must be open and writable")
                fmt = (SaveFormat.PDF if isinstance(save_format_or_options, PdfSaveOptions) else
                       _coerce_save_format(getattr(save_format_or_options, "save_format", save_format_or_options)))
                data = self.to_bytes(save_format_or_options)
                commits = ()
                if fmt == SaveFormat.DOCX or fmt in MAIN_TYPES:
                    from dockiln.fonts import prepare_font_table_commit
                    from dockiln.utils.table_row import prepare_row_height_commit
                    from dockiln._font_sizes import prepare_saved_compatibility_commit
                    commits = (prepare_saved_compatibility_commit(self.light_document_model, data),
                               prepare_font_table_commit(self.light_document_model, data),
                               prepare_row_height_commit(self.light_document_model))
                try:
                    if fmt == SaveFormat.DOCX:
                        output_path.seek(0)
                    written = output_path.write(data)
                    if written != len(data):
                        raise OSError("Output stream accepted an incomplete write")
                except (OSError, ValueError) as error:
                    raise RuntimeError("Output stream could not be written") from error
                for commit in commits:
                    commit()
            else:
                fmt = self._save(output_path, save_format_or_options)
            return SaveOutputParameters._from_content_type(_SAVE_CONTENT_TYPES[fmt])

    def _save(self, output_path, save_format_or_options) -> SaveFormat:
        doc = self.light_document_model
        output_path = Path(output_path)
        if isinstance(save_format_or_options, PdfSaveOptions):
            self._save_as_pdf(output_path, doc, save_format_or_options)
            return SaveFormat.PDF
        if isinstance(save_format_or_options, OoxmlSaveOptions):
            self._save_as_docx(output_path, doc, save_format_or_options)
            return _coerce_save_format(save_format_or_options.save_format)

        fmt = _coerce_save_format(save_format_or_options)
        if fmt is None and save_format_or_options is None:
            fmt = {".pdf": SaveFormat.PDF, ".docx": SaveFormat.DOCX}.get(output_path.suffix.lower())
        if fmt in MAIN_TYPES:
            self._save_as_docx(output_path, doc, OoxmlSaveOptions(fmt))
        elif fmt == SaveFormat.PDF:
            self._save_as_pdf(output_path, doc)
        elif fmt == SaveFormat.DOCX:
            self._save_as_docx(output_path, doc)
        else:
            raise ValueError(
                f"Unsupported save format: {save_format_or_options!r}. "
                "Supported formats: PDF, DOCX, Flat OPC."
            )
        return fmt

    def to_bytes(
        self,
        save_format_or_options: SaveFormat | int | str | PdfSaveOptions | OoxmlSaveOptions,
    ) -> bytes:
        """Render in memory. An explicit format is required; no files are created.

        Diagnostics accumulate alongside load/save diagnostics, including on failure.
        """
        doc = self.light_document_model
        with collect_diagnostics(self.diagnostics):
            fmt = _coerce_save_format(save_format_or_options)
            if isinstance(save_format_or_options, PdfSaveOptions) or fmt == SaveFormat.PDF:
                from dockiln.writers.pdf import LdmPdfWriter
                pdf_options = save_format_or_options if isinstance(save_format_or_options, PdfSaveOptions) else None
                return LdmPdfWriter(pdf_options).write_to_bytes(doc)
            if isinstance(save_format_or_options, OoxmlSaveOptions) or fmt == SaveFormat.DOCX or fmt in MAIN_TYPES:
                from dockiln.writers.docx import LdmDocxWriter
                docx_options = (save_format_or_options if isinstance(save_format_or_options, OoxmlSaveOptions)
                                else OoxmlSaveOptions(fmt) if fmt in MAIN_TYPES else None)
                return LdmDocxWriter(docx_options).write_to_bytes(doc)
            raise ValueError("Unsupported save format; supported formats: PDF, DOCX, Flat OPC")

    def to_dict(self) -> dict:
        """Export ordered content with model locations, not inferred PDF page positions."""
        from dataclasses import asdict
        from dockiln.structure import export_document

        result = export_document(self.light_document_model, self._source_name)
        result["diagnostics"] = [asdict(item) for item in self.diagnostics]
        return result





    def _save_as_pdf(
        self,
        output_path: Path,
        doc: ldm.Document,
        options: Optional[PdfSaveOptions] = None,
    ) -> None:
        """Convert the loaded LDM to PDF and save."""
        from dockiln.writers.pdf import LdmPdfWriter

        writer = LdmPdfWriter(options)
        writer.write(doc, output_path)

    def _save_as_docx(
        self,
        output_path: Path,
        doc: ldm.Document,
        options: Optional[OoxmlSaveOptions] = None,
    ) -> None:
        """Convert the loaded LDM to DOCX and save."""
        from dockiln.writers.docx import LdmDocxWriter

        from dockiln.fonts import prepare_font_table_commit

        writer = LdmDocxWriter(options)
        data = writer.write_to_bytes(doc)
        commit_fonts = prepare_font_table_commit(doc, data)
        from dockiln.utils.table_row import prepare_row_height_commit
        commit_rows = prepare_row_height_commit(doc)
        from dockiln._font_sizes import prepare_saved_compatibility_commit
        commit_settings = prepare_saved_compatibility_commit(doc, data)
        with atomic_output(output_path) as temporary:
            temporary.write_bytes(data)
        commit_fonts()
        commit_rows()
        commit_settings()
