"""Paragraph-local, half-open text ranges; stale ranges fail instead of shifting silently."""

from aspose.words_foss.dom.nodes import (
    _replace_span,
    _safe_structure,
    _size_value,
    _text,
    _text_spans,
    _validate_text,
    _validate_toggle,
)

_UNSET = object()


class TextRange:
    def __init__(self, paragraph, start=0, end=None):
        nodes = paragraph._plain_text_nodes()
        length = sum(len(_text(node)) for node in nodes)
        if end is None:
            end = length
        if type(start) is not int or type(end) is not int or not 0 <= start <= end <= length:
            raise ValueError("Text range requires integer bounds 0 <= start <= end <= paragraph length")
        self._paragraph = paragraph
        self._start = start
        self._end = end
        self._snapshot = paragraph.xml

    @property
    def paragraph(self):
        return self._paragraph

    @property
    def start(self):
        return self._start

    @property
    def end(self):
        return self._end

    def _check(self):
        if self.paragraph.xml != self._snapshot:
            raise ValueError("Text range is stale; select a new range after other paragraph edits")
        return self.paragraph._plain_text_nodes()

    @property
    def text(self):
        self._check()
        return self.paragraph.text[self.start:self.end]

    def replace(self, text):
        """Replace selection; collapsed insertion uses the right-hand run or last run at end."""
        _validate_text(text)
        nodes = self._check()
        if text == self.paragraph.text[self.start:self.end]:
            return self
        if nodes:
            _replace_span(_text_spans(nodes), self.start, self.end, text)
            self.paragraph._changed()
        elif self.paragraph.runs:
            self.paragraph.runs[-1].text = text
        else:
            document = self.paragraph.owner_document
            self.paragraph.append_child(document.create_run(text, part_name=self.paragraph.part_name))
        self._end = self.start + len(text)
        self._snapshot = self.paragraph.xml
        return self

    def _selected_runs(self):
        selected, offset = [], 0
        for run in self.paragraph.runs:
            end = offset + len(run.text)
            if offset < self.end and end > self.start:
                selected.append(run)
            offset = end
        return selected

    def _split_at(self, position):
        offset = 0
        for run in self.paragraph.runs:
            end = offset + len(run.text)
            if offset < position < end:
                run.split(position - offset)
                return
            offset = end

    def apply_font(self, *, bold=_UNSET, italic=_UNSET, size=_UNSET):
        """Apply direct formatting to the selection, splitting boundary runs as needed."""
        self._check()
        values = {name: value for name, value in (("bold", bold), ("italic", italic), ("size", size))
                  if value is not _UNSET}
        for name, value in values.items():
            if name == "size":
                _size_value(value)
            else:
                _validate_toggle(value)
        if not values:
            return self
        if self.start == self.end:
            raise ValueError("Font formatting requires a nonempty range")
        self.paragraph._structural_editable()
        if any(not _safe_structure(run._element) for run in self._selected_runs()):
            raise NotImplementedError("Range formatting with unsupported XML metadata is unsafe")
        # Validate the entire selection before splitting anything.
        self._split_at(self.end)
        self._split_at(self.start)
        for run in self._selected_runs():
            for name, value in values.items():
                setattr(run.font, name, value)
        self._snapshot = self.paragraph.xml
        return self
