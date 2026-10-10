"""Owned comment scalar metadata, warm dates, source mapping and atomic saves."""
from datetime import date, datetime, timedelta, timezone
from io import BytesIO
from zipfile import ZipFile
from xml.etree.ElementTree import fromstring

import pytest

from aspose.words_foss import NodeType
from aspose.words_foss.dom import DocxDocument
from aspose.words_foss.dom.importing import ImportFormatMode
from .test_dom_note_lifecycle import document

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def source_date(tmp_path, raw):
    doc = document(tmp_path)
    comment = doc.create_comment('Owned author', 'OA')
    comment.ensure_minimum()
    doc.body.paragraphs[0].append_child(comment)
    with ZipFile(BytesIO(doc.to_bytes())) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    if raw is not None:
        parts['word/comments.xml'] = parts['word/comments.xml'].replace(
            b'<w:comment ', ('<w:comment w:date="' + raw + '" ').encode(), 1)
    path = tmp_path / 'dated.docx'
    with ZipFile(path, 'w') as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)
    return DocxDocument(path)


def comment(doc):
    return doc.get_child_nodes(NodeType.COMMENT, True)[0]


@pytest.mark.parametrize('raw, expected', [
    (None, datetime.min), ('', datetime.min), ('Owned invalid date', datetime.min),
    ('0000-01-01T00:00:00Z', datetime.min), ('0001-01-01T00:00:00Z', datetime.min),
    ('2026-13-01T00:00:00Z', datetime.min), ('2026-10-11T24:00:00Z', datetime.min),
    ('2026-10-11T12:34:60Z', datetime.min), ('20261011', datetime.min),
    ('2026-10-11T12:34:56+15:00', datetime.min),
    ('2026-10-11', datetime(2026, 10, 11)),
    ('2026-10-11T12:34:56Z', datetime(2026, 10, 11, 12, 34, 56)),
    ('2026-10-11T12:34:56-05:00', datetime(2026, 10, 11, 12, 34, 56)),
    ('2026-10-11T12:34:56+0800', datetime(2026, 10, 11, 12, 34, 56)),
    (' 2026-10-11t12:34:56z ', datetime(2026, 10, 11, 12, 34, 56)),
    ('2026-10-11T12:34:56.1234567Z', datetime(2026, 10, 11, 12, 34, 56, 123456)),
    ('9999-12-31T23:59:59.9999999Z', datetime.max),
])
@pytest.mark.parametrize('flat', [False, True])
def test_loaded_date_keeps_wall_clock_and_warm_precision(tmp_path, raw, expected, flat):
    doc = source_date(tmp_path, raw)
    node = comment(doc)
    value = node.date_time
    assert value.replace(tzinfo=None) == expected
    assert (value.tzinfo is None) == (expected == datetime.min)
    output = tmp_path / 'result'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(output)
        assert node.date_time is value
        cold = comment(DocxDocument(output))
        assert cold.date_time.replace(tzinfo=None) == expected.replace(microsecond=0)
        assert cold.author == 'Owned author' and cold.initial == 'OA'
        xml = fromstring(DocxDocument(output).part_xml('word/comments.xml'))
        lexical = xml.find(W + 'comment').get(W + 'date')
        assert lexical == (None if expected == datetime.min else expected.isoformat(timespec='seconds') + 'Z')


@pytest.mark.parametrize('flat', [False, True])
def test_unvisited_comment_dates_are_projected_only_in_typed_parts(tmp_path, flat):
    doc = source_date(tmp_path, '2026-10-11T12:34:56.123456Z')
    assert doc._note_graph is None
    opaque = b'<x:root xmlns:x="urn:owned" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:date="2026-10-11T12:34:56.123456Z"/></x:root>'
    parts = dict(doc._package._payloads)
    parts['customXml/owned.xml'] = opaque
    parts['[Content_Types].xml'] = parts['[Content_Types].xml'].replace(b'</Types>', b'<Override PartName="/customXml/owned.xml" ContentType="application/xml"/></Types>')
    raw = BytesIO()
    with ZipFile(raw, 'w') as archive:
        for name, payload in parts.items():archive.writestr(name, payload)
    doc = DocxDocument(BytesIO(raw.getvalue()))
    output = tmp_path / 'result'
    (doc.save_flat_opc if flat else doc.save)(output)
    cold = DocxDocument(output)
    assert comment(cold).date_time.microsecond == 0
    saved_opaque = cold._package.payload('customXml/owned.xml')
    if not flat:assert saved_opaque == opaque
    shape = lambda raw: [(node.tag, node.attrib, node.text, node.tail) for node in fromstring(raw).iter()]
    assert shape(saved_opaque) == shape(opaque)
    assert doc._note_graph is None


@pytest.mark.parametrize('value', [datetime(2026, 10, 11, 12, 34, 56, 123456),
    datetime(2026, 10, 11, 12, 34, 56, 123456, tzinfo=timezone(timedelta(hours=8))),
    date(2026, 10, 11), datetime.min, datetime.max])
@pytest.mark.parametrize('flat', [False, True])
def test_create_and_edit_date_warm_state_survives_saves(tmp_path, value, flat):
    doc = document(tmp_path)
    node = doc.create_comment('Owned', 'O', value)
    doc.body.paragraphs[0].append_child(node)
    expected = value.astimezone(timezone.utc) if isinstance(value, datetime) and value.tzinfo else datetime(value.year, value.month, value.day) if not isinstance(value, datetime) else value
    assert node.date_time == expected and node.date_time.tzinfo == expected.tzinfo
    node.date_time = value
    output = tmp_path / 'result'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(output)
        assert node.date_time == expected and node.date_time.tzinfo == expected.tzinfo
        assert comment(DocxDocument(output)).date_time.replace(tzinfo=None) == expected.replace(tzinfo=None, microsecond=0)


@pytest.mark.parametrize('value', [None, 0, False, 'Owned invalid date', 1.5,
    datetime.min.replace(tzinfo=timezone(timedelta(hours=8))),
    datetime.max.replace(tzinfo=timezone(timedelta(hours=-5)))])
def test_invalid_date_setter_and_constructor_preserve_model(tmp_path, value):
    doc = source_date(tmp_path, '2026-10-11T12:34:56Z')
    node = comment(doc);before = doc.to_bytes();held = node.date_time
    error = OverflowError if isinstance(value, datetime) else TypeError
    with pytest.raises(error):node.date_time = value
    assert node.date_time is held and doc.to_bytes() == before
    with pytest.raises(TypeError):doc.create_comment('Owned', 'O', value)
    assert doc.to_bytes() == before


@pytest.mark.parametrize('operation', ['clone', 'shallow', 'import', 'paragraph-clone'])
@pytest.mark.parametrize('flat', [False, True])
def test_date_warm_state_is_rebound_on_copy_and_import(tmp_path, operation, flat):
    doc = source_date(tmp_path, None)
    original = comment(doc)
    value = datetime(2026, 10, 11, 12, 34, 56, 123456)
    original.date_time = value
    target = doc
    if operation == 'import':
        target = document(tmp_path)
        copied = target.import_node(original, True, ImportFormatMode.USE_DESTINATION_STYLES)
        target.body.paragraphs[0].append_child(copied)
    elif operation == 'paragraph-clone':
        paragraph = original.parent_node.clone(True)
        target.body.append_child(paragraph)
        copied = paragraph.get_child_nodes(NodeType.COMMENT, True)[0]
    else:
        copied = original.clone(operation != 'shallow')
        target.body.paragraphs[1].append_child(copied)
    assert copied.date_time == value and copied.date_time.tzinfo is None
    output = tmp_path / 'result'
    (target.save_flat_opc if flat else target.save)(output)
    assert copied.date_time == value and original.date_time == value
    assert all(node.date_time.microsecond == 0 for node in DocxDocument(output).get_child_nodes(NodeType.COMMENT, True))


def test_date_cache_does_not_mask_direct_xml_changes_or_wrong_identity(tmp_path):
    node = comment(source_date(tmp_path, None))
    node.date_time = datetime(2026, 10, 11, 12, 34, 56, 123456)
    copied = node.clone(True)
    copied._element._comment_date_warm = node._element._comment_date_warm
    assert copied.date_time.tzinfo is not None
    node._element.setAttributeNS(W[1:-1], 'w:date', '2026-10-12T00:00:00Z')
    assert node.date_time.replace(tzinfo=None) == datetime(2026, 10, 12)


@pytest.mark.parametrize('flat', [False, True])
def test_failed_date_save_preserves_warm_metadata_and_original_file(tmp_path, monkeypatch, flat):
    from aspose.words_foss import _io
    doc = source_date(tmp_path, '2026-10-11T12:34:56.123456Z');node = comment(doc)
    node.date_time = datetime(2026, 10, 12, 12, 34, 56, 654321)
    held = node.date_time;before = node._element.toxml();output = tmp_path / 'result';output.write_bytes(b'Owned existing output')
    def fail(*args):raise OSError('Owned failed publish')
    monkeypatch.setattr(_io.os, 'replace', fail)
    with pytest.raises(OSError):(doc.save_flat_opc if flat else doc.save)(output)
    assert node.date_time is held and node._element.toxml() == before
    assert output.read_bytes() == b'Owned existing output'


@pytest.mark.parametrize('field', ['author', 'initial'])
@pytest.mark.parametrize('value, error', [(None, RuntimeError), (0, TypeError), (False, TypeError)])
def test_text_metadata_validation_is_shared_and_atomic(tmp_path, field, value, error):
    doc = source_date(tmp_path, None);node = comment(doc);before = doc.to_bytes()
    with pytest.raises(error):setattr(node, field, value)
    with pytest.raises(error):doc.create_comment(**{field: value})
    assert doc.to_bytes() == before


@pytest.mark.parametrize('field', ['author', 'initial'])
@pytest.mark.parametrize('flat', [False, True])
def test_unicode_text_metadata_edits_survive_cold_save(tmp_path, field, flat):
    doc = source_date(tmp_path, None);node = comment(doc)
    value = 'Owned 中<&"\t\n'
    setattr(node, field, value);output = tmp_path / 'result'
    (doc.save_flat_opc if flat else doc.save)(output)
    assert getattr(node, field) == getattr(comment(DocxDocument(output)), field) == value


@pytest.mark.parametrize('raw, nonzero', [
    ('0001-01-01T00:00:00.0000001Z', True),
    ('0001-01-01T00:00:00.00000005Z', True),
    ('0001-01-01T00:00:00.00000001Z', False),
])
@pytest.mark.parametrize('flat', [False, True])
def test_submicrosecond_minimum_retains_source_kind_and_date_presence(tmp_path, raw, nonzero, flat):
    doc = source_date(tmp_path, raw);node = comment(doc);held = node.date_time
    assert held.replace(tzinfo=None) == datetime.min
    assert (held.tzinfo is not None) == nonzero
    copied = node.clone(True);doc.body.paragraphs[1].append_child(copied)
    assert copied.date_time == held and copied.date_time.tzinfo == held.tzinfo
    output = tmp_path / 'result'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(output)
        assert node.date_time is held
        cold = DocxDocument(output)
        entries = fromstring(cold.part_xml('word/comments.xml')).findall(W + 'comment')
        assert [entry.get(W + 'date') for entry in entries] == ['0001-01-01T00:00:00Z' if nonzero else None] * 2
        assert all(entry.date_time == datetime.min for entry in cold.get_child_nodes(NodeType.COMMENT, True))


@pytest.mark.parametrize('raw, expected', [
    ('2026-10-11T12:34:56.99999995Z', datetime(2026, 10, 11, 12, 34, 57)),
    ('9999-12-31T23:59:59.99999995Z', datetime.min),
])
def test_source_tick_rounding_carries_or_recovers_overflow(tmp_path, raw, expected):
    assert comment(source_date(tmp_path, raw)).date_time.replace(tzinfo=None) == expected


def test_datetime_subclass_and_fold_are_normalized(tmp_path):
    class OwnedDateTime(datetime):pass
    node = comment(source_date(tmp_path, None))
    for value in [OwnedDateTime(2026, 10, 11), datetime(2026, 11, 1, 1, 30, fold=1),
                  datetime(2026, 11, 1, 1, 30, fold=1, tzinfo=timezone.utc)]:
        node.date_time = value
        assert type(node.date_time) is datetime and node.date_time.fold == 0


@pytest.mark.parametrize('zone, offsets', [
    ('UTC', [0, 0, 0, 0]),
    ('America/New_York', [-240, -300, -300, -297]),
    ('Asia/Singapore', [480, 480, 480, 416]),
])
def test_source_dates_use_local_standard_time_and_historical_minute_precision(tmp_path, monkeypatch, zone, offsets):
    import time
    if not hasattr(time, 'tzset'):pytest.skip('Explicit POSIX timezone control is unavailable')
    monkeypatch.setenv('TZ', zone);time.tzset()
    try:
        for raw, minutes in zip(['2026-10-11T12:34:56Z', '2026-03-08T02:30:00Z',
                                 '2026-11-01T01:30:00Z', '0001-01-01T00:00:01Z'], offsets):
            value = comment(source_date(tmp_path, raw)).date_time
            assert value.utcoffset() == timedelta(minutes=minutes)
    finally:
        monkeypatch.undo();time.tzset()


@pytest.mark.parametrize('variant', ['same', 'conflict', 'alias', 'collision'])
@pytest.mark.parametrize('flat', [False, True])
def test_duplicate_content_date_keeps_first_winner_and_warm_state(tmp_path, variant, flat):
    from .test_dom_duplicate_note_ids import duplicate_document
    doc, part, _ = duplicate_document(tmp_path, 'comment', variant)
    storage = doc.story(part).child_nodes
    storage[0]._element.setAttributeNS(W[1:-1], 'w:date', '2026-10-11T12:34:56.1234567+08:00')
    storage[1]._element.setAttributeNS(W[1:-1], 'w:date', '2030-01-02T03:04:05Z')
    node = comment(doc)
    warm = node.date_time
    assert warm.replace(tzinfo=None) == datetime(2026, 10, 11, 12, 34, 56, 123456)
    original = doc.part_xml(part)
    pure = doc.to_flat_opc() if flat else doc.to_bytes()
    assert doc.part_xml(part) == original
    assert node.date_time is warm
    assert comment(DocxDocument(BytesIO(pure))).date_time.replace(tzinfo=None) == warm.replace(tzinfo=None, microsecond=0)
    output = tmp_path / 'dated-normalized'
    (doc.save_flat_opc if flat else doc.save)(output)
    assert node.date_time is warm
    assert storage[1].parent_node is None
    assert comment(DocxDocument(output)).date_time.replace(tzinfo=None) == warm.replace(tzinfo=None, microsecond=0)
