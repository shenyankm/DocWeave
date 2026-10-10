"""Bounded, inert HTML table import for raw HTML blocks in Markdown."""

import re
from dataclasses import dataclass, field
from html.parser import HTMLParser

from aspose.words_foss import _io
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.diagnostics import ContentLossWarning, warn

MAX_HTML_NODES = 100_000
MAX_HTML_DEPTH = 64
VOID = {"br", "hr", "img", "input", "meta", "link", "col", "source", "wbr"}
BLOCKED = {"script", "style", "iframe", "object", "embed", "head"}
INLINE = {
    "b": "bold",
    "strong": "bold",
    "i": "italic",
    "em": "italic",
    "s": "strike_through",
    "strike": "strike_through",
    "del": "strike_through",
    "sup": "superscript",
    "sub": "subscript",
}


class TableBlockBoundary(HTMLParser):
    """Track real table tags, ignoring lookalikes in comments and script text."""

    def __init__(self):
        super().__init__()
        self.depth = 0
        self.seen = False

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.seen = True
            self.depth += 1
            if self.depth >= MAX_HTML_DEPTH:
                raise ValueError("HTML table input exceeds the nesting safety limit")

    def handle_endtag(self, tag):
        if tag == "table":
            self.depth = max(0, self.depth - 1)

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        self.handle_endtag(tag)


@dataclass
class Element:
    tag: str
    attrs: dict = field(default_factory=dict)
    children: list = field(default_factory=list)


class TableHTMLParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Element("root")
        self.stack = [self.root]
        self.count = 0
        self.reported = set()
        self.generated = 0

    def report(self, code, message):
        if code not in self.reported:
            self.reported.add(code)
            warn(message, ContentLossWarning, code=code)

    def append(self, node):
        self.count += 1
        if self.count > MAX_HTML_NODES:
            raise ValueError("HTML table input exceeds the node safety limit")
        self.stack[-1].children.append(node)

    def handle_starttag(self, tag, attrs):
        # Repair omitted cell/row closing tags within the current table only.
        boundary = next(
            (
                i
                for i in range(len(self.stack) - 1, 0, -1)
                if self.stack[i].tag == "table"
            ),
            0,
        )
        closable = (
            {"td", "th"}
            if tag in {"td", "th"}
            else {"td", "th", "tr"}
            if tag == "tr"
            else set()
        )
        for i in range(len(self.stack) - 1, boundary, -1):
            if self.stack[i].tag in closable:
                del self.stack[i:]
                self.report(
                    "markdown.html_repaired",
                    "Unclosed HTML table cells or rows were repaired",
                )
        if len(self.stack) >= MAX_HTML_DEPTH:
            raise ValueError("HTML table input exceeds the nesting safety limit")
        node = Element(tag, dict(attrs))
        if tag in BLOCKED:
            self.report(
                "markdown.html_active_content",
                "Active HTML content was omitted without execution",
            )
        if node.attrs.get("style") or any(key.startswith("on") for key in node.attrs):
            self.report(
                "markdown.html_css", "HTML CSS and event formatting are not retained"
            )
        supported_attrs = (
            {"dir"}
            if tag == "table"
            else {"colspan", "rowspan"}
            if tag in {"td", "th"}
            else {"src", "alt"}
            if tag == "img"
            else {"href"}
            if tag == "a"
            else set()
        )
        if any(
            key not in supported_attrs and key != "style" and not key.startswith("on")
            for key in node.attrs
        ):
            self.report(
                "markdown.html_semantics",
                "Unsupported HTML attributes are not retained",
            )
        if tag in {"col", "colgroup", "caption"}:
            self.report(
                "markdown.html_semantics",
                "HTML column and caption layout is not retained",
            )
        self.append(node)
        if tag not in VOID:
            self.stack.append(node)

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in VOID:
            self.handle_endtag(tag)

    def handle_endtag(self, tag):
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                if i != len(self.stack) - 1:
                    self.report(
                        "markdown.html_repaired",
                        "Unclosed HTML table cells or rows were repaired",
                    )
                del self.stack[i:]
                return
        if tag not in VOID:
            self.report(
                "markdown.html_repaired", "Unmatched HTML closing tags were ignored"
            )

    def handle_data(self, data):
        self.append(data)

    def handle_decl(self, decl):
        raise ValueError("HTML declarations are not supported in table imports")

    def span(self, node, key, limit):
        try:
            value = int(node.attrs.get(key, "1"))
        except (TypeError, ValueError):
            value = 1
            self.report(
                "markdown.html_span", "Invalid HTML table spans were treated as one"
            )
        if value <= 0:
            value = 1  # Fixed 26.9 treats rowspan=0 and negative spans as one.
            self.report(
                "markdown.html_span",
                "Non-positive HTML table spans were treated as one",
            )
        if value > limit:
            raise ValueError("HTML table span exceeds the safety limit")
        return value

    def new_cell(self, **kwargs):
        self.generated += 1
        if self.generated > MAX_HTML_NODES:
            raise ValueError("Expanded HTML table exceeds the cell safety limit")
        return ldm.Cell(**kwargs)

    def table(self, node, font):
        def rows(children, heading=False):
            for child in children:
                if isinstance(child, Element):
                    if child.tag == "tr":
                        yield child, heading
                    elif child.tag in {"thead", "tbody", "tfoot"}:
                        yield from rows(child.children, child.tag == "thead")
                    else:
                        self.report(
                            "markdown.html_table_content",
                            "HTML content outside rows is not retained",
                        )
                elif child.strip():
                    self.report(
                        "markdown.html_table_content",
                        "Text outside HTML table cells is not retained",
                    )

        source_rows = list(rows(node.children))
        table = ldm.Table(bidi=node.attrs.get("dir", "").lower() == "rtl")
        active, widths = {}, []
        for index, (source, heading) in enumerate(source_rows):
            active = {
                start: value for start, value in active.items() if value[1] > index
            }
            row = ldm.Row(row_format=ldm.RowFormat(heading_format=heading))
            column = 0

            def continuation(row=row, active=active):
                nonlocal column
                span, _end = active[column]
                row.cells.append(
                    self.new_cell(
                        cell_format=ldm.CellFormat(grid_span=span, vertical_merge=2),
                        paragraphs=[ldm.Paragraph()],
                    )
                )
                column += span

            for cell in source.children:
                if not isinstance(cell, Element) or cell.tag not in {"td", "th"}:
                    if isinstance(cell, Element) or cell.strip():
                        self.report(
                            "markdown.html_table_content",
                            "Text outside HTML table cells is not retained",
                        )
                    continue
                while column in active and active[column][1] > index:
                    continuation()
                span = self.span(cell, "colspan", _io.MAX_TABLE_COLUMNS)
                rowspan = min(
                    self.span(cell, "rowspan", MAX_HTML_NODES), len(source_rows) - index
                )
                if column + span > _io.MAX_TABLE_COLUMNS:
                    raise ValueError("HTML table grid exceeds the column safety limit")
                if any(
                    column < start < column + span and end > index
                    for start, (_, end) in active.items()
                ):
                    raise ValueError("HTML table cells overlap a row span")
                cell_font = font.model_copy(deep=True)
                if cell.tag == "th":
                    cell_font.bold = True
                children = self.blocks(cell.children, cell_font)
                # Fixed 26.9 omits contentless cell paragraphs, but a br-created
                # paragraph survives even when its final break is consumed.
                children = [
                    child
                    for child in children
                    if not isinstance(child, ldm.Paragraph)
                    or any(run.text for run in child._children)
                ]
                for child in children:
                    if isinstance(child, ldm.Paragraph):
                        last = next(
                            (run for run in reversed(child._children) if run.text), None
                        )
                        if last is not None and last.text.endswith("\n"):
                            last.text = last.text[:-1]
                if children and isinstance(children[-1], ldm.Table):
                    children.append(ldm.Paragraph())
                if cell.tag == "th":
                    for child in children:
                        if isinstance(child, ldm.Paragraph):
                            child.paragraph_format.alignment = 1
                row.cells.append(
                    self.new_cell(
                        cell_format=ldm.CellFormat(
                            grid_span=span, vertical_merge=1 if rowspan > 1 else 0
                        ),
                        children=children or [ldm.Paragraph()],
                    )
                )
                if rowspan > 1:
                    active[column] = (span, index + rowspan)
                column += span
            pending = [
                (start, span)
                for start, (span, end) in active.items()
                if end > index and start >= column
            ]
            for start, _ in sorted(pending):
                while column < start:
                    row.cells.append(self.new_cell(paragraphs=[ldm.Paragraph()]))
                    column += 1
                continuation()
            widths.append(column)
            table.rows.append(row)
        # Fixed 26.9 collapses columns whose boundary never occurs in any row.
        # A lone colspan=2 becomes one; an unmerged reference row retains two.
        boundaries = {0}
        for row in table.rows:
            column = 0
            for cell in row.cells:
                column += cell.cell_format.grid_span
                boundaries.add(column)
        offsets = {value: index for index, value in enumerate(sorted(boundaries))}
        columns = len(offsets) - 1
        for row, width in zip(table.rows, widths):
            column = 0
            for cell in row.cells:
                end = column + cell.cell_format.grid_span
                cell.cell_format.grid_span = offsets[end] - offsets[column]
                column = end
            # Native fills short rows with real empty cells, including rowspan=0.
            for _ in range(columns - offsets[width]):
                row.cells.append(self.new_cell(paragraphs=[ldm.Paragraph()]))
        return table

    def blocks(self, children, font):
        result, current = [], None

        def paragraph():
            nonlocal current
            if current is None:
                current = ldm.Paragraph()
                result.append(current)
            return current

        def visit(nodes, style):
            nonlocal current
            for node in nodes:
                if isinstance(node, str):
                    text = re.sub(r"[ \t\r\n\f]+", " ", node)
                    if text.strip() or current is not None:
                        paragraph()._children.append(
                            ldm.Run(text=text, font=style.model_copy(deep=True))
                        )
                    continue
                if node.tag in BLOCKED:
                    self.report(
                        "markdown.html_active_content",
                        "Active HTML content was omitted without execution",
                    )
                    continue
                if node.tag == "table":
                    current = None
                    result.append(self.table(node, style))
                    continue
                if node.tag in {
                    "p",
                    "div",
                    "body",
                    "html",
                    "caption",
                    "h1",
                    "h2",
                    "h3",
                    "h4",
                    "h5",
                    "h6",
                }:
                    current = None
                    if node.tag == "p":
                        paragraph()
                    elif node.tag.startswith("h") and node.tag[1:].isdigit():
                        self.report(
                            "markdown.html_semantics",
                            "HTML heading styling is not retained",
                        )
                    visit(node.children, style)
                    current = None
                    continue
                if node.tag == "br":
                    paragraph()._children.append(
                        ldm.Run(text="\n", font=style.model_copy(deep=True))
                    )
                    continue
                changed = style.model_copy(deep=True)
                if node.tag in INLINE:
                    setattr(changed, INLINE[node.tag], True)
                elif node.tag == "u":
                    changed.underline = 1
                elif node.tag in {"img", "a", "video", "audio"}:
                    self.report(
                        "markdown.html_resources",
                        "HTML resources and link actions were not loaded",
                    )
                    if node.tag == "img" and node.attrs.get("alt"):
                        paragraph()._children.append(
                            ldm.Run(text=node.attrs["alt"], font=changed)
                        )
                elif node.tag not in {"span", "em"}:
                    self.report(
                        "markdown.html_semantics",
                        "Unsupported HTML semantics were reduced to visible text",
                    )
                if node.attrs.get("style"):
                    self.report(
                        "markdown.html_css",
                        "HTML CSS formatting is not fully supported",
                    )
                visit(node.children, changed)

        visit(children, font)
        for block in result:
            if isinstance(block, ldm.Paragraph) and block._children:
                if isinstance(block._children[0], ldm.Run):
                    block._children[0].text = block._children[0].text.lstrip(" ")
                if isinstance(block._children[-1], ldm.Run):
                    block._children[-1].text = block._children[-1].text.rstrip(" ")
        return result


def parse_html_tables(text, font):
    _io.check_input_size(len(text.encode("utf-8")))
    parser = TableHTMLParser()
    parser.feed(text)
    parser.close()
    if len(parser.stack) > 1:
        parser.report(
            "markdown.html_repaired",
            "Unclosed HTML elements were repaired at end of input",
        )
    return parser.blocks(parser.root.children, font)
