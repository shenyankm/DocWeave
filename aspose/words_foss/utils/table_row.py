"""Shared row-height measures for canonical DOM and conversion snapshots."""

import re
from math import isfinite

MAX_SAVED_ROW_HEIGHT = 1584.0


def parse_row_split(value):
    """Return allow-break for the strict OOXML cantSplit boolean."""
    if value in (None, "", "1", "true", "on"):
        return False
    if value in ("0", "false", "off"):
        return True
    raise RuntimeError("Invalid row-splitting boolean")


def prepare_row_height_commit(document):
    from aspose.words_foss.light_document_model import NodeType

    plan = [row.row_format for row in document.get_child_nodes(NodeType.ROW, True)
            if row.row_format.height_explicit is not False
            and row.row_format.height > MAX_SAVED_ROW_HEIGHT]

    def commit():
        for properties in plan:
            properties.__dict__["height"] = MAX_SAVED_ROW_HEIGHT

    return commit


def coerce_row_height(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TypeError("height requires a number")
    scaled = float(value) * 20
    if not isfinite(scaled):
        return -2147483648 / 20
    return min(2147483647, max(-2147483648, round(scaled))) / 20


def parse_row_height(value, rule=""):
    """Read native signed-32 row twips; a source zero always resolves Auto."""
    from aspose.words_foss.docx_reader.ldm_builder._helpers import parse_universal_measure

    value = value or ""
    last = next((i for i in range(len(value) - 1, -1, -1) if '0' <= value[i] <= '9'), -1)
    unit = value[last + 1:] if last >= 0 else ""
    if unit and unit.isspace():
        return 0.0, 2
    if unit and unit not in {"pt", "in", "cm", "mm", "pc", "pi"}:
        raise RuntimeError("Invalid row-height measurement unit")
    if re.fullmatch(r"[+-]?[0-9]+", value):
        digits = value.lstrip('+-').lstrip('0') or '0'
        number = (-9223372036854775808 if value.startswith('-') else 9223372036854775807) if len(digits) > 19 else int(digits) * (-1 if value.startswith('-') else 1)
    else:
        number = (parse_universal_measure(value, 'pt') * 20 if unit else
                  parse_universal_measure(value, 'twip')) if last >= 0 and '_' not in value else 0.0
    number = min(9223372036854775807, max(-9223372036854775808, number))
    twips = round(number) if unit else int(number)
    twips = (twips + 2147483648) % 4294967296 - 2147483648
    return twips / 20, 2 if twips == 0 else 1 if rule == 'exact' else 0
