"""Structural protocol for the PDF sub-renderers' ``writer`` parameter.

Lets the sub-renderer modules type their ``writer`` without importing
:class:`LdmPdfWriter` (which would form a runtime circular import).
``LdmPdfWriter`` satisfies it structurally.
"""

from typing import Any, ContextManager, Optional, Protocol, Union

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.saving import PdfSaveOptions


class PDFWriterContext(Protocol):
    options: PdfSaveOptions

    _page_width: float
    _page_height: float
    _page_margin_left: float
    _page_full_margin_left: float
    _page_margin_right: float
    _page_margin_bottom: float
    _page_number_offset: int
    _paragraph_insets: tuple[float, float]

    _anchor_links: dict[str, int]
    _outline_levels: list[int]
    _pre_rendered_shapes: set[int]

    _paragraph_renderer: Any
    _shape_renderer: Any
    _table_renderer: Any

    def _link_target_for(self, pdf: FPDF, url: Optional[str]) -> Union[int, str]: ...

    def _structure(self, pdf: FPDF, struct_type: str, source: Any,
                   position: Any = None) -> ContextManager[None]: ...

    def _artifact(self, pdf: FPDF, subtype: Optional[str] = None) -> ContextManager[None]: ...

    def _list_structure(self, pdf: FPDF, list_format: ldm.ListFormat,
                        source: Any = None) -> ContextManager[None]: ...

    def _end_list(self, pdf: FPDF) -> None: ...

    def _cell_structure(self, pdf: FPDF, table: ldm.Table, row_index: int,
                        column: int) -> ContextManager[None]: ...

    def _estimate_paragraph_height(self, para: ldm.Paragraph, col_w_mm: float) -> float: ...

    def _estimate_child_height(
        self, child: Union[ldm.Paragraph, ldm.Table], col_w_mm: float
    ) -> float: ...
