"""Structured conversion diagnostics without changing Python warning filters."""

from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
from typing import Iterator
import warnings


@dataclass(frozen=True)
class ConversionDiagnostic:
    """One known conversion limitation; locations are best-effort, not page numbers."""

    code: str
    severity: str
    location: str | None
    message: str


class ConversionWarning(UserWarning):
    """Base for known content loss, substitutions and unsupported options."""

    code = "conversion.loss"


class ContentLossWarning(ConversionWarning):
    """Known content is omitted or its layout cannot be retained."""


_collector: ContextVar[list[ConversionDiagnostic] | None] = ContextVar(
    "conversion_diagnostics", default=None
)


@contextmanager
def collect_diagnostics(target: list[ConversionDiagnostic]) -> Iterator[None]:
    token = _collector.set(target)
    try:
        yield
    finally:
        _collector.reset(token)


def warn(message: str, category: type[ConversionWarning], *, stacklevel: int = 2,
         code: str | None = None, location: str | None = None) -> None:
    """Record even ignored/error-filtered warnings; do not intercept unrelated warnings."""
    collector = _collector.get()
    if collector is not None:
        collector.append(ConversionDiagnostic(code or category.code, "warning", location, message))
    warnings.warn(message, category, stacklevel=stacklevel + 1)


def header_footer_losses(doc):
    for header_type in (0, 1):
        parts = [hf for sec in doc.sections for hf in sec.headers_footers
                 if hf.header_footer_type == header_type]
        if parts and any(hf.children != parts[0].children for hf in parts[1:]):
            yield "header_footer_flattened", ("Different per-section headers/footers cannot be preserved; "
                                               "only the first populated part is written")
            break
    if any(hf.header_footer_type not in (0, 1) for sec in doc.sections for hf in sec.headers_footers):
        yield "header_footer_variant", "First-page/even-page header/footer variants are omitted"


def document_nodes(doc):
    """Walk model nodes, including unknown body nodes and text-box paragraphs."""
    from aspose.words_foss import light_document_model as ldm

    stack = [doc]
    while stack:
        node = stack.pop()
        yield node
        children = list(ldm._walk_children(node))
        if isinstance(node, ldm.Shape) and node.text_box:
            children.extend(node.text_box.get("paragraphs", []) or [])
        children.extend(c for c in getattr(node, "children", [])
                        if isinstance(c, ldm.UnknownNode))
        stack.extend(reversed(children))
