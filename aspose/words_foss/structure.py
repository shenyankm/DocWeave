"""Small, JSON-safe content export for search/knowledge-base ingestion."""

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._visible_runs import visible_runs
from aspose.words_foss.pdf_writer.text import apply_caps, extract_link_segments


def _paragraph(node: ldm.Paragraph, location: str) -> dict:
    pf = node.paragraph_format
    runs = [{"text": apply_caps(text, run.font), "link": link, "bold": run.font.bold,
             "italic": run.font.italic, "size": run.font.size}
            for run in visible_runs(node) if not run.font.hidden
            for text, link in extract_link_segments(run.text)]
    result = {
        "type": "paragraph", "location": location,
        "text": "".join(run["text"] for run in runs), "style": pf.style_name,
        "heading_level": pf.outline_level + 1 if pf.is_heading else None,
        "list": node.list_format.model_dump() if node.list_format else None,
        "runs": runs,
        "images": [{"alternative_text": item.alternative_text,
                    "width_pt": item.width, "height_pt": item.height}
                   for item in node._children if isinstance(item, ldm.Shape) and item.has_image],
    }
    return result


def _block(node, location: str) -> dict:
    if isinstance(node, ldm.Paragraph):
        return _paragraph(node, location)
    if isinstance(node, ldm.Table):
        rows = []
        for r, row in enumerate(node.rows):
            cells = []
            for c, cell in enumerate(row.cells):
                path = f"{location}.rows[{r}].cells[{c}]"
                cells.append({
                    "location": path,
                    "grid_span": cell.cell_format.grid_span,
                    "horizontal_merge": cell.cell_format.horizontal_merge,
                    "vertical_merge": cell.cell_format.vertical_merge,
                    "paragraphs": [_paragraph(p, f"{path}.paragraphs[{i}]")
                                   for i, p in enumerate(cell.paragraphs)],
                    "tables": [_block(t, f"{path}.tables[{i}]")
                               for i, t in enumerate(cell.tables)],
                })
            rows.append({"heading": row.row_format.heading_format, "cells": cells})
        return {"type": "table", "location": location, "title": node.title, "rows": rows}
    return {"type": "unsupported", "location": location, "node_type": node.type}


def export_document(doc: ldm.Document, source: str | None) -> dict:
    blocks, headers_footers = [], []
    for s, section in enumerate(doc.sections):
        path = f"sections[{s}]"
        blocks.extend(_block(node, f"{path}.body.children[{i}]")
                      for i, node in enumerate(section.body.children))
        for i, hf in enumerate(section.headers_footers):
            headers_footers.append({
                "location": f"{path}.headers_footers[{i}]", "type": hf.header_footer_type,
                "blocks": [_block(node, f"{path}.headers_footers[{i}].children[{j}]")
                           for j, node in enumerate(hf.children)],
            })
    # ponytail: the LDM separates cell paragraphs/tables; source interleaving and page coordinates are unavailable.
    return {"schema_version": 1, "source": source, "blocks": blocks,
            "headers_footers": headers_footers}
