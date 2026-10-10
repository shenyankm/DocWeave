"""Comment date mapping, including the native distinction between warm and XML values."""

import os
import re
import time
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_DATE = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}(?:[Tt ][0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]+)?(?:[Zz]|[+-][0-9]{2}:?[0-9]{2})?)?\Z")


def normalize_text(value):
    if value is None:
        raise RuntimeError("Comment text metadata cannot be None")
    if not isinstance(value, str):
        raise TypeError("Comment text metadata requires a string")
    return value


def normalize_date(value, *, constructor=False):
    if isinstance(value, datetime):
        try:
            value = value.astimezone(timezone.utc) if value.utcoffset() is not None else value.replace(tzinfo=None)
        except (OverflowError, ValueError) as error:
            if isinstance(error, OverflowError) and not constructor:
                raise
            raise TypeError("Comment date_time cannot be converted to datetime") from error
        return datetime(value.year, value.month, value.day, value.hour, value.minute,
                        value.second, value.microsecond, tzinfo=value.tzinfo)
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    raise TypeError("Comment date_time requires datetime or date")


def _local_wall(value):
    """Keep the XML wall clock; native selects standard time at DST transitions."""
    zone = None
    try:
        if os.environ.get("TZ"):
            zone = ZoneInfo(os.environ["TZ"])
        else:
            with open("/etc/localtime", "rb") as source:
                zone = ZoneInfo.from_file(source)
    except (OSError, ValueError, ZoneInfoNotFoundError):
        pass
    base = timedelta(seconds=-time.timezone)
    if zone is not None:
        first = value.replace(tzinfo=zone, fold=0)
        second = value.replace(tzinfo=zone, fold=1)
        selected = second if second.dst() == timedelta(0) else first
        offset = selected.utcoffset()
    else:
        try:
            offset = value.astimezone().utcoffset()
        except (OverflowError, ValueError, OSError):
            offset = base
    # TimeZoneInfo adjustment deltas have minute precision, including historic LMT.
    offset = base + timedelta(minutes=int((offset - base).total_seconds() / 60))
    return value.replace(tzinfo=timezone(offset))


def _parse_date(raw):
    lexical = raw.strip()
    if not _DATE.fullmatch(lexical) or (len(lexical) > 10 and int(lexical[11:13]) >= 24):
        return datetime.min, 0
    fraction = re.search(r"\.([0-9]+)", lexical)
    ticks = 0
    if fraction:
        digits = fraction[1]
        ticks = int(digits[:7].ljust(7, '0')) + (len(digits) > 7 and digits[7] >= '5')
        lexical = lexical[:fraction.start()] + lexical[fraction.end():]
    # Normalize forms that Python 3.10's narrower ISO parser cannot consume.
    lexical = lexical.replace("t", "T")
    lexical = re.sub(r"([+-][0-9]{2})([0-9]{2})$", r"\1:\2", lexical)
    if lexical.endswith(("Z", "z")):
        lexical = lexical[:-1] + "+00:00"
    try:
        value = datetime.fromisoformat(lexical)
        offset = value.utcoffset()
        if offset is not None and abs(offset) > timedelta(hours=14):
            return datetime.min, 0
        value = value.replace(tzinfo=None) + timedelta(microseconds=ticks // 10)
    except (ValueError, OverflowError):
        return datetime.min, 0
    residual = ticks % 10
    return (datetime.min if value == datetime.min and not residual else _local_wall(value)), residual


def parse_date(raw):
    return _parse_date(raw)[0]


def get_date(element):
    raw = element.getAttributeNS(W, "date")
    state = getattr(element, "_comment_date_warm", None)
    if state is not None and state[0] is element and state[1] == raw:
        return state[2]
    value, residual = _parse_date(raw)
    element._comment_date_warm = (element, raw, value, residual)
    return value


def set_date(element, value):
    from aspose.words_foss.dom.importing import _set_word_attribute
    value = normalize_date(value)
    raw = value.isoformat()
    _set_word_attribute(element, "date", raw)
    element._comment_date_warm = (element, raw, value, 0)


def copy_date_state(source, target):
    state = getattr(source, "_comment_date_warm", None)
    if (state is not None and state[0] is source
            and source.getAttributeNS(W, "date") == target.getAttributeNS(W, "date") == state[1]):
        target._comment_date_warm = (target, *state[1:])


def project_date(source, target):
    """Emit seconds and a Z suffix without committing cold precision to held nodes."""
    from aspose.words_foss.dom.importing import _set_word_attribute
    value = get_date(source).replace(tzinfo=None)
    residual = source._comment_date_warm[3]
    raw = "" if value == datetime.min and not residual else value.isoformat(timespec="seconds") + "Z"
    if raw == target.getAttributeNS(W, "date") and (raw or not target.hasAttributeNS(W, "date")):
        return False
    if raw:
        _set_word_attribute(target, "date", raw)
    elif target.hasAttributeNS(W, "date"):
        target.removeAttributeNS(W, "date")
    return True


def projected_comment_tree(source, projected=None):
    root = source.documentElement
    if root.namespaceURI != W or root.localName != "comments":
        return source if projected is None else projected
    target = source if projected is None else projected
    for index, old in enumerate(root.childNodes):
        if (old.nodeType != old.ELEMENT_NODE or old.namespaceURI != W
                or old.localName != "comment" or old.getAttributeNS(W, "type") not in {"", "normal"}):
            continue
        probe = old.cloneNode(deep=False)
        if not project_date(old, probe):
            continue
        if target is source:
            target = source.cloneNode(deep=True)
        # The root child order is preserved by the row projection and XML clone.
        project_date(old, target.documentElement.childNodes[index])
    return target
