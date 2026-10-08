"""Read OLE2/DOC files carrying an .rtf suffix, not standard text RTF."""

from pathlib import Path
from typing import Optional, Union, BinaryIO, Iterator

from aspose.words_foss._io import check_input_size, read_bounded
from aspose.words_foss.doc_reader import DocFileReader
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader import (
    ParagraphData,
    TableData,
    NumberingInfo,
)



class RtfFileReader:
    """
    Reads RTF files (OLE2-format) and produces the same data structures
    as DocumentReader (for .docx) and DocFileReader (for .doc).

    Standard text RTF is rejected with a native-conversion hint.
    An OLE2 file renamed to .rtf is handled by DocFileReader.
    """

    def __init__(self):
        self._delegate = DocFileReader()

    def load_file(self, filepath: Union[str, Path]) -> None:
        """Load .rtf from file path."""
        check_input_size(Path(filepath).stat().st_size)
        with Path(filepath).open("rb") as stream:
            self.load_stream(stream)

    def load_stream(self, stream: BinaryIO) -> None:
        """Load .rtf from stream."""
        self.load_bytes(read_bounded(stream))

    def load_bytes(self, data: bytes) -> None:
        """Load .rtf from bytes."""
        check_input_size(len(data))
        if not data.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
            raise ValueError(
                "Standard text RTF is not supported by the builtin reader. "
                "Use aspose.words_foss.libreoffice.convert_to_pdf(original_path, output_path) "
                "or the CLI --backend libreoffice for PDF conversion. "
                "Only OLE2/DOC-backed .rtf files can be loaded into the light document model."
            )
        self._delegate.load_bytes(data)

    def _iterate_body_elements(self) -> Iterator[Union[ParagraphData, TableData]]:
        """Iterate over document body elements in order."""
        return self._delegate._iterate_body_elements()

    def _get_list_format(self, num_id: int, level: int) -> tuple[str, int]:
        """Get list format and start number for a given list id and level."""
        return self._delegate._get_list_format(num_id, level)

    def _get_numbering_info(self, num_id: int) -> Optional[NumberingInfo]:
        """Get numbering info (for compatibility with DocumentReader)."""
        return self._delegate._get_numbering_info(num_id)

    def to_light_document(self) -> ldm.Document:
        """Build a light_document_model.Document from the loaded RTF file."""
        return self._delegate.to_light_document()
