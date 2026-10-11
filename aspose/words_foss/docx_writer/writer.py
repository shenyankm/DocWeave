"""LdmDocxWriter — orchestrates DOCX generation from an LDM Document.

Public surface:

* ``write(doc, output_path)`` — render ``doc`` and zip it to disk.
* ``write_to_bytes(doc)``     — same but returns the raw ``bytes`` of
  the resulting ``.docx`` for in-memory use (tests, services).
"""


import io
import zipfile
from pathlib import Path
from typing import Optional, Union

from dataclasses import dataclass

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._flat_opc import MAIN_TYPES, encode
from aspose.words_foss._io import atomic_output
from aspose.words_foss.diagnostics import ContentLossWarning, ConversionWarning, document_nodes, header_footer_losses, source_story_losses, warn
from aspose.words_foss.saving import (
    CompressionLevel,
    OoxmlCompliance,
    Zip64Mode,
    coerce_enum,
)
from aspose.words_foss.docx_writer.bookmarks import BookmarkState
from aspose.words_foss.docx_writer.document_part import _build_style_pf_map, render_document_xml
from aspose.words_foss.docx_writer.drawing import ImageEntry, ImageRenderState
from aspose.words_foss.docx_writer.headers_footers import render_section_stories
from aspose.words_foss.docx_writer.numbering_part import render_numbering_xml
from aspose.words_foss.docx_writer.package import write_docx_package
from aspose.words_foss.docx_writer.settings_part import needs_settings, render_settings_xml
from aspose.words_foss.docx_writer.styles_part import (
    build_style_font_map,
    build_style_id_map,
    render_styles_xml,
    apply_reference_styles,
)


class DocxWriterLossyWarning(ContentLossWarning):
    """Known LDM content cannot be faithfully written as DOCX."""

    code = "docx.content_loss"


def _warn_about_unsupported_constructs(doc: ldm.Document) -> None:
    if any(isinstance(node, ldm.UnknownNode) for node in document_nodes(doc)):
        warn("Unknown document nodes are omitted from DOCX", DocxWriterLossyWarning,
             stacklevel=3, code="docx.unknown_node")
    for code, message in (*header_footer_losses(doc), *source_story_losses(doc)):
        if code in {"comments_omitted", "header_footer_flattened", "header_footer_variant"}:
            continue  # DOCX/Flat OPC emit classic comment stories and ordered anchors.
        warn(message, DocxWriterLossyWarning, stacklevel=3, code="docx." + code)


_COMPRESSION_MAP: dict[int, tuple[int, int]] = {
    CompressionLevel.NORMAL: (zipfile.ZIP_DEFLATED, 6),
    CompressionLevel.MAXIMUM: (zipfile.ZIP_DEFLATED, 9),
    CompressionLevel.FAST: (zipfile.ZIP_DEFLATED, 3),
    CompressionLevel.SUPER_FAST: (zipfile.ZIP_DEFLATED, 1),
}

_ZIP64_MAP: dict[int, bool] = {
    Zip64Mode.NEVER: False,
    Zip64Mode.IF_NECESSARY: True,
    Zip64Mode.ALWAYS: True,
}


class LdmDocxWriter:
    """Convert an :class:`ldm.Document` into DOCX or Flat OPC XML.

    The writer is intentionally stateless across calls — every
    invocation produces a fresh hyperlink relationship map.  This keeps
    parallel/sequential exports independent and avoids the bug-prone
    ``reset()``-style APIs the reader had to evolve.
    """

    def __init__(self, options: Optional[object] = None):
        # ``options`` accepts an ``OoxmlSaveOptions`` instance only —
        # No separate ``DocxSaveOptions`` is exposed —
        # ``DocxSaveOptions``.  Weakly typed so callers don't have to
        # import the save-options module to use the writer.
        self.options = options
        self._enforce_compliance()

    # ------------------------------------------------------------------
    def _enforce_compliance(self) -> None:
        """Reject compliance levels we cannot satisfy.

        ECMA-376 and ISO 29500 Transitional share the same on-disk
        layout for the constructs we emit, so both are accepted.
        Strict ISO 29500 mandates a separate namespace and forbids
        legacy elements (e.g. ``w:cantSplit``) — silently emitting
        transitional XML when the caller asked for Strict would be
        worse than failing fast.
        """
        compliance = getattr(self.options, "compliance", None)
        if compliance == OoxmlCompliance.ISO29500_2008_STRICT:
            raise NotImplementedError(
                "OoxmlCompliance.ISO29500_2008_STRICT is not supported by "
                "the FOSS writer; use ECMA376_2006 or "
                "ISO29500_2008_TRANSITIONAL."
            )

    def _compression(self) -> tuple[int, int]:
        """Return ``(compress_type, compresslevel)`` for :mod:`zipfile`."""
        level = coerce_enum(
            CompressionLevel, getattr(self.options, "compression_level", CompressionLevel.NORMAL)
        )
        return _COMPRESSION_MAP.get(level, _COMPRESSION_MAP[CompressionLevel.NORMAL])

    def _allow_zip64(self) -> bool:
        mode = getattr(self.options, "zip_64_mode", Zip64Mode.NEVER)
        return _ZIP64_MAP.get(coerce_enum(Zip64Mode, mode), False)

    def _pretty_format(self) -> bool:
        return bool(getattr(self.options, "pretty_format", False))

    def _render_parts(self, doc: ldm.Document) -> "_RenderedParts":
        """Render every XML part the package needs and return them bundled."""
        from aspose.words_foss.document import SaveFormat, _coerce_save_format

        if _coerce_save_format(getattr(self.options, "save_format", SaveFormat.DOCX)) not in {SaveFormat.DOCX, *MAIN_TYPES}:
            raise ValueError("OoxmlSaveOptions.save_format must be DOCX or a Flat OPC variant")
        if coerce_enum(Zip64Mode, getattr(self.options, "zip_64_mode", Zip64Mode.NEVER)) == Zip64Mode.ALWAYS:
            warn("DOCX zip_64_mode=ALWAYS does not force ZIP64 records; behaves like IF_NECESSARY",
                 ConversionWarning, code="docx.unsupported_option")
        _warn_about_unsupported_constructs(doc)
        # The hyperlink table and image state are mutated by :mod:`runs`
        # / :mod:`paragraphs` while rendering the document body; gather
        # them first, then build dependents.
        hyperlinks: dict[str, str] = {}
        image_state = ImageRenderState()
        bookmark_state = BookmarkState()
        styles_xml = render_styles_xml(doc)
        reference = getattr(self.options, "reference_docx", None)
        if reference is not None:
            if not isinstance(reference, (str, Path)):
                raise ValueError("reference_docx must be a DOCX path")
            from aspose.words_foss.docx_reader import DocumentReader
            reader = DocumentReader()
            reader.load_file(reference)
            styles_xml = apply_reference_styles(styles_xml, reader.to_light_document())
        numbering_xml = render_numbering_xml(doc) or None
        settings_xml = render_settings_xml(doc) if needs_settings(doc) else None

        style_pf_map = _build_style_pf_map(doc)
        style_id_map = build_style_id_map(doc)
        style_font_map = build_style_font_map(doc)
        header_footer_parts, section_refs = render_section_stories(
            doc, style_pf_map, style_id_map, style_font_map)
        document_xml = render_document_xml(doc, hyperlinks, image_state=image_state,
            bookmark_state=bookmark_state, section_header_footer_refs=section_refs)

        from aspose.words_foss._comments import render_comments
        comments_xml, comment_hyperlinks, comment_images = render_comments(doc, style_pf_map, style_id_map, style_font_map)
        comments_extended_xml = None
        if comments_xml is not None:
            from aspose.words_foss.dom.notes import annotation_id
            values = {str(annotation_id(story.identifier)): story.comment_metadata.model_dump()
                      for story in doc.source_stories if story.kind == "comment" and story.comment_metadata is not None}
            if any(value.get("ancestor_identifier") is not None or value.get("done") is not None for value in values.values()):
                from aspose.words_foss._comment_threads import extension_xml
                comments_xml, comments_extended_xml = extension_xml(comments_xml, values)
        return _RenderedParts(
            comments_extended_xml=comments_extended_xml,
            comments_xml=comments_xml, comment_hyperlinks=comment_hyperlinks, comment_images=comment_images,
            theme_xml=doc.source_theme.data if doc.source_theme is not None else None,
            source_theme=doc.source_theme,
            source_font_table=doc.source_font_table,
            embed_fonts=bool(doc.font_embedding and doc.font_embedding.embed_true_type_fonts),
            document_xml=document_xml,
            styles_xml=styles_xml,
            numbering_xml=numbering_xml,
            settings_xml=settings_xml,
            hyperlinks=hyperlinks,
            images=image_state.images,
            header_xml=None, header_hyperlinks={}, header_images=[],
            footer_xml=None, footer_hyperlinks={}, footer_images=[],
            header_footer_parts=header_footer_parts,
        )

    # ------------------------------------------------------------------
    def write(self, doc: ldm.Document, output_path: Union[str, Path]) -> None:
        """Render ``doc`` and atomically write DOCX or explicit Flat OPC XML."""
        from aspose.words_foss.document import _coerce_save_format

        if _coerce_save_format(getattr(self.options, "save_format", None)) in MAIN_TYPES:
            with atomic_output(output_path) as temporary:
                temporary.write_bytes(self.write_to_bytes(doc))
            return
        parts = self._render_parts(doc)
        compress_type, compresslevel = self._compression()
        with atomic_output(output_path) as temporary:
            write_docx_package(
                temporary,
                **parts.as_kwargs(),
                compression=compress_type,
                compresslevel=compresslevel,
                allow_zip64=self._allow_zip64(),
                pretty_format=self._pretty_format(),
            )

    def write_to_bytes(self, doc: ldm.Document) -> bytes:
        """Render ``doc`` and return DOCX or explicit Flat OPC XML bytes."""
        parts = self._render_parts(doc)
        buffer = io.BytesIO()
        compress_type, compresslevel = self._compression()
        write_docx_package(
            buffer,
            **parts.as_kwargs(),
            compression=compress_type,
            compresslevel=compresslevel,
            allow_zip64=self._allow_zip64(),
            pretty_format=self._pretty_format(),
        )
        data = buffer.getvalue()
        from aspose.words_foss import _io
        from zipfile import ZipFile

        _io.check_input_size(len(data))
        with ZipFile(io.BytesIO(data)) as package:
            _io.validate_docx_archive(package)
        from aspose.words_foss.document import _coerce_save_format

        fmt = _coerce_save_format(getattr(self.options, "save_format", None))
        return encode(data, MAIN_TYPES[fmt]) if fmt in MAIN_TYPES else data


@dataclass
class _RenderedParts:
    """Bundle of fully-rendered DOCX part payloads for the package writer."""

    theme_xml: bytes | None
    source_theme: ldm.SourceTheme | None
    source_font_table: ldm.SourceFontTable | None
    embed_fonts: bool
    comments_xml: str | None
    comments_extended_xml: str | None
    comment_hyperlinks: dict[str, str]
    comment_images: list[ImageEntry]
    document_xml: str
    styles_xml: str
    numbering_xml: Optional[str]
    settings_xml: Optional[str]
    hyperlinks: dict[str, str]
    images: list[ImageEntry]
    header_footer_parts: list
    header_xml: Optional[str]
    header_hyperlinks: dict[str, str]
    header_images: list[ImageEntry]
    footer_xml: Optional[str]
    footer_hyperlinks: dict[str, str]
    footer_images: list[ImageEntry]

    def as_kwargs(self) -> dict[str, object]:
        """Return the kwargs ``write_docx_package`` accepts."""
        return {
            "comments_xml": self.comments_xml,
            "comments_extended_xml": self.comments_extended_xml,
            "comment_hyperlinks": self.comment_hyperlinks,
            "comment_images": self.comment_images,
            "theme_xml": self.theme_xml,
            "source_theme": self.source_theme,
            "source_font_table": self.source_font_table,
            "embed_fonts": self.embed_fonts,
            "document_xml": self.document_xml,
            "styles_xml": self.styles_xml,
            "numbering_xml": self.numbering_xml,
            "settings_xml": self.settings_xml,
            "hyperlinks": self.hyperlinks,
            "images": self.images,
            "header_footer_parts": self.header_footer_parts,
            "header_xml": self.header_xml,
            "header_hyperlinks": self.header_hyperlinks,
            "header_images": self.header_images,
            "footer_xml": self.footer_xml,
            "footer_hyperlinks": self.footer_hyperlinks,
            "footer_images": self.footer_images,
        }
