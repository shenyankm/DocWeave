"""Native short-offset comment date recovery with real save/cold paths."""
from datetime import datetime, timedelta, timezone
from io import BytesIO

import pytest

from aspose.words_foss.dom import DocxDocument
from .test_dom_comment_metadata import comment, source_date
from .test_dom_comment_utc import modern


@pytest.mark.parametrize('suffix, minutes', [
    ('+0', 0), ('+00', 0), ('-00', 0), ('+8', 480), ('+08', 480),
    ('-08', -480), ('+008', 8), ('+000', 0), ('+830', 510), ('+140', 100),
    ('-140', -100), ('+1400', 840), ('+14', 840), ('+8:3', 483),
    ('+8:30', 510), ('+14:0', 840), ('+0:8', 8), ('+0:08', 8),
    ('+08:0', 480), ('+08:00', 480),
    ('+15', None), ('+080', None), ('+1401', None), ('+14:01', None),
    ('+08:000', None), ('+00000', None), ('+00:', None), ('+8:', None),
    ('+8:60', None), ('+15:00', None),
])
@pytest.mark.parametrize('modern_value', [False, True])
@pytest.mark.parametrize('flat', [False, True])
def test_short_zone_source_warm_and_two_cold_saves(tmp_path, suffix, minutes, modern_value, flat):
    raw = '2030-01-02T03:04:05' + suffix
    doc = modern(tmp_path, raw) if modern_value else source_date(tmp_path, raw)
    held = comment(doc);wall = datetime(2030, 1, 2, 3, 4, 5)
    if modern_value:
        expected = (wall - timedelta(minutes=minutes)).replace(tzinfo=timezone.utc) if minutes is not None else held.date_time_utc
        assert held.date_time_utc == expected
        assert held.date_time_utc.year == (2030 if minutes is not None else 2026)
    else:
        expected = wall if minutes is not None else datetime.min
        assert held.date_time.replace(tzinfo=None) == expected
    warm = (held.date_time, held.date_time_utc)
    originals = {name: doc._package.payload(name) for name in doc._package.part_names}
    data = doc.to_flat_opc() if flat else doc.to_bytes()
    cold = comment(DocxDocument(BytesIO(data)))
    if modern_value:assert cold.date_time_utc == expected
    else:assert cold.date_time.replace(tzinfo=None) == expected
    assert {name: doc._package.payload(name) for name in originals} == originals
    for index in range(2):
        out = tmp_path / str(index)
        (doc.save_flat_opc if flat else doc.save)(out)
        assert comment(doc) is held and (held.date_time, held.date_time_utc) == warm
        cold = comment(DocxDocument(out))
        if modern_value:assert cold.date_time_utc == expected
        else:assert cold.date_time.replace(tzinfo=None) == expected
