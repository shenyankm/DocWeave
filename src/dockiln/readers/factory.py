"""Reader factory and protocol for document format readers.

Provides a common Protocol defining the reader interface and a factory
function that returns the appropriate reader for a given file extension.
"""

from pathlib import Path
from typing import BinaryIO, Protocol, Union, runtime_checkable

from dockiln import light_document_model as ldm


@runtime_checkable
class DocumentFormatReader(Protocol):
    """Protocol defining the interface all document readers must implement."""

    def load_file(self, filepath: Union[str, Path]) -> None: ...

    def load_stream(self, stream: BinaryIO) -> None: ...

    def load_bytes(self, data: bytes) -> None: ...

    def to_light_document(self) -> ldm.Document: ...


def create_reader(suffix: str) -> DocumentFormatReader:
    """Create the appropriate reader based on file extension.

    Args:
        suffix: Lowercase file extension including dot (e.g. ".docx").

    Returns:
        A reader instance implementing the DocumentFormatReader protocol.
    """
    if suffix == ".doc":
        from dockiln.readers.doc import DocFileReader

        return DocFileReader()
    elif suffix == ".rtf":
        from dockiln.readers.rtf import RtfFileReader

        return RtfFileReader()
    elif suffix in {".docx", ".flatopc"}:
        from dockiln.readers.docx import DocumentReader

        return DocumentReader()

    raise ValueError(f"Unsupported input format: {suffix!r}; expected DOC, DOCX, RTF, or Flat OPC")
