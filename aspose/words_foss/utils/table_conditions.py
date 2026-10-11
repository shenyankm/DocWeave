"""Shared table-region precedence for ordinary font-size getters."""


def _regions(row, rows, column, span, columns, look, bands, *, getter):
    first_row = row == 0 and look.first_row
    last_row = row == rows - 1 and look.last_row
    first_column = column == 0 and look.first_column
    last_column = column + span == columns and look.last_column
    kinds = []
    if look.column_banding and not (first_column or last_column):
        band = (
            ((column + 1 - int(look.first_column)) // bands["tblStyleColBandSize"] - 1)
            if getter
            else ((column - int(look.first_column)) // bands["tblStyleColBandSize"])
        ) % 2
        kinds.append("band2Vert" if band else "band1Vert")
    if look.row_banding and not (first_row or last_row):
        band = (
            ((row + 1 - int(look.first_row)) // bands["tblStyleRowBandSize"] - 1)
            if getter
            else ((row - int(look.first_row)) // bands["tblStyleRowBandSize"])
        ) % 2
        kinds.append("band2Horz" if band else "band1Horz")
    kinds.extend(
        name
        for name, enabled in (
            ("firstCol", first_column),
            ("lastCol", last_column),
            ("firstRow", first_row),
            ("lastRow", last_row),
            ("nwCell", first_row and first_column),
            ("neCell", first_row and last_column),
            ("swCell", last_row and first_column),
            ("seCell", last_row and last_column),
        )
        if enabled
    )
    return kinds


def table_font_size(layers, conditions):
    """Resolve size per base-to-derived style; the first region declaration owns its slot."""
    size = None
    for whole, regions in layers:
        if whole is not None:
            size = whole
        first = {}
        for kind, value in regions:
            first.setdefault(kind, value)
        for kind in conditions:
            value = first.get(kind)
            if value is not None:
                size = value
    return size


def table_size_conditions(positions, look, bands, declared):
    """Ordinary size getter regions, independent of name/Boolean/paint rules.

    Empty first-row/column groups shift band numbering. Last-row look excludes
    row bands; column bands include boundary columns until a local size wins.
    """
    result = {}
    for row, cells in enumerate(positions):
        for column, (cell, _, _) in enumerate(cells):
            kinds = []
            if look.column_banding:
                offset = int(look.first_column and "firstCol" in declared)
                band = ((column + 1 - offset) // bands["tblStyleColBandSize"] - 1) % 2
                kinds.append("band2Vert" if band else "band1Vert")
            if look.row_banding and not (look.last_row and row == len(positions) - 1):
                offset = int(look.first_row and "firstRow" in declared)
                band = ((row + 1 - offset) // bands["tblStyleRowBandSize"] - 1) % 2
                kinds.append("band2Horz" if band else "band1Horz")
            boundary = _regions(row, len(positions), column, 1, len(cells), look, bands, getter=True)
            kinds.extend(kind for kind in boundary if not kind.startswith("band"))
            result[id(cell)] = tuple(kinds)
    return result


def table_run_size(default, whole, conditional, paragraph, character, direct,
                   *, table_style, paint=False, override_normal=False):
    """Ordinary Run size; native Normal's 12pt sentinel differs from paragraph marks."""
    size = conditional if conditional is not None else default
    if (table_style and not paint and not override_normal and paragraph
            and paragraph[-1] == (True, 12.0)):
        size = whole if whole is not None else default
        if size == 10.0:
            size = 12.0
    else:
        for normal, value in paragraph:
            if value is not None and not (paint and table_style and normal and value == 12.0
                                          and not override_normal):
                size = value
    for value in (character, direct):
        if value is not None:
            size = value
    return size


def table_size_override(mode, settings, *, paint=False):
    """Only literal zero disables this native compatSetting; first declaration wins."""
    namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    if settings is not None:
        for setting in settings.iter(namespace + 'compatSetting'):
            if (setting.get(namespace + 'name') == 'overrideTableStyleFontSizeAndJustification'
                    and setting.get(namespace + 'uri') == 'http://schemas.microsoft.com/office/word'):
                return setting.get(namespace + 'val', '') != '0'
    return paint and mode in (11, 15)

def table_conditions(positions, columns, look, bands, column_band_declared=False):
    getters, rendering = {}, {}
    for row_index, cells in enumerate(positions):
        for cell_index, (cell, column, span) in enumerate(cells):
            # Native name getters count physical cells; painting uses grid positions.
            getters[id(cell)] = tuple(
                _regions(
                    row_index,
                    len(positions),
                    cell_index,
                    1,
                    len(cells),
                    look,
                    bands,
                    getter=True,
                )
            )
            kinds = _regions(
                row_index,
                len(positions),
                column,
                span,
                columns,
                look,
                bands,
                getter=False,
            )
            if (
                column_band_declared
                and len(kinds) > 1
                and kinds[0].endswith("Vert")
                and kinds[1].endswith("Horz")
            ):
                kinds[:2] = reversed(kinds[:2])
            rendering[id(cell)] = tuple(kinds)
    return getters, rendering



def table_size_paint_conditions(positions, columns, look, bands, declared_bands):
    """Paint band sizes only when their interval is declared; getters default to one."""
    prefix = uniform_omitted_prefix([cells[0][1] if cells else None for cells in positions])
    if prefix:
        positions = [[(cell, column - prefix, span) for cell, column, span in cells]
                     for cells in positions]
        columns -= prefix
    _, rendering = table_conditions(positions, columns, look, bands,
                                    "tblStyleColBandSize" in declared_bands)
    return {key: tuple(kind for kind in kinds
                       if not kind.startswith("band") or
                       ("tblStyleRowBandSize" if kind.endswith("Horz") else
                        "tblStyleColBandSize") in declared_bands)
            for key, kinds in rendering.items()}


def uniform_omitted_prefix(offsets):
    """A uniformly absent source prefix has no physical column in the table."""
    if not offsets or offsets[0] is None:
        return 0
    prefix = offsets[0]
    return prefix if all(value == prefix for value in offsets) else 0


def table_font_paint_size(layers, conditions):
    """Painting combines whole styles first, then inherited condition declarations."""
    size = None
    for whole, _ in layers:
        if whole is not None:
            size = whole
    for kind in conditions:
        for _, regions in layers:
            first = {}
            for region, value in regions:
                first.setdefault(region, value)
            if first.get(kind) is not None:
                size = first[kind]
    return size


