"""Observed default-tab value coercion shared by the conversion and package DOMs."""
from math import isfinite


def coerce_default_tab_stop(value: float) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TypeError('default_tab_stop requires int or float')
    if isinstance(value, int) and not -2**63 <= value < 2**64:
        raise OverflowError('default_tab_stop integer is out of range')
    if value < 0:
        raise RuntimeError('default_tab_stop must be nonnegative')
    # Native nonfinite inputs use the signed minimum; finite values clamp.
    twips = -2147483648 if not isfinite(value) else round(min(value, 2147483647 / 20) * 20)
    return twips / 20


def parse_default_tab_stop(value: str) -> float:
    from dockiln.utils.table_row import parse_row_height
    last = next((i for i in range(len(value) - 1, -1, -1) if '0' <= value[i] <= '9'), -1)
    suffix = value[last + 1:] if last >= 0 else ''
    if suffix and suffix not in {'pt', 'in', 'cm', 'mm', 'pc', 'pi', '\n'}:
        raise RuntimeError('Invalid default-tab measurement unit')
    return parse_row_height(value)[0]
