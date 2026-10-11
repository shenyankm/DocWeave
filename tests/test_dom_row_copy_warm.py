"""A copied live row keeps setter state without sharing source XML/cache identity."""
from io import BytesIO

import pytest

import dockiln as dk
from dockiln.dom import DocxDocument
from .test_dom_row_format import row_document
from .test_dom_story_row_height import W, story_package


@pytest.mark.parametrize('story', ['body', 'header', 'footer'])
@pytest.mark.parametrize('operation', ['clone', 'import'])
@pytest.mark.parametrize('flat', [False, True])
def test_zero_exact_row_copy_warm_and_cold(tmp_path, story, operation, flat):
    if story == 'body':
        path, _ = row_document(tmp_path)
        document = DocxDocument(path)
        parent = document.body
    else:
        path = tmp_path / 'source.docx'
        name, _ = story_package(path, story)
        document = DocxDocument(path)
        parent = document.story(name)
    table = parent.get_child_nodes(dk.NodeType.TABLE, True)[0]
    source = table.get_child_nodes(dk.NodeType.ROW, True)[0].row_format
    source.height = 0
    source.height_rule = dk.HeightRule.EXACTLY
    if operation == 'clone':
        copied = table.clone(True)
        parent.append_child(copied)
        destination = document
        rows = parent.get_child_nodes(dk.NodeType.ROW, True)
        copied_index = 1
    else:
        target = tmp_path / 'target.docx'
        story_package(target, 'header')
        destination = DocxDocument(target)
        copied = destination.import_node(table, True)
        destination.body.append_child(copied)
        rows = destination.body.get_child_nodes(dk.NodeType.ROW, True)
        copied_index = 0
    held = rows[copied_index].row_format
    assert held.height == 0
    assert held.height_rule is dk.HeightRule.EXACTLY
    pure = destination.to_flat_opc() if flat else destination.to_bytes()
    assert held.height_rule is dk.HeightRule.EXACTLY
    output = tmp_path / 'output'
    (destination.save_flat_opc if flat else destination.save)(output)
    assert held.height_rule is dk.HeightRule.EXACTLY
    for payload in [pure, output.read_bytes()]:
        cold = DocxDocument(BytesIO(payload))
        cold_parent = cold.body if operation == 'import' or story == 'body' else cold.story(name)
        format = cold_parent.get_child_nodes(dk.NodeType.ROW, True)[copied_index].row_format
        assert format.height == 0
        assert format.height_rule is dk.HeightRule.AUTO
    held.height = 3
    assert held.height == 3 and held.height_rule is dk.HeightRule.EXACTLY
    assert source.height == 0 and source.height_rule is dk.HeightRule.EXACTLY


def test_shallow_clone_and_stale_cache_rebind(tmp_path):
    path, _ = row_document(tmp_path)
    document = DocxDocument(path)
    row = document.body.get_child_nodes(dk.NodeType.ROW, True)[0]
    row.row_format.height = 0
    row.row_format.height_rule = dk.HeightRule.EXACTLY
    cloned = row.clone(False)
    assert cloned.row_format.height_rule is dk.HeightRule.EXACTLY
    assert cloned.get_child_nodes(dk.NodeType.CELL, True) == []
    # Editing the copied declaration invalidates only its own warm projection.
    height = cloned._element.getElementsByTagNameNS(W, 'trHeight')[0]
    height.setAttributeNS(W, 'w:hRule', 'auto')
    assert cloned.row_format.height_rule is dk.HeightRule.AUTO
    assert row.row_format.height_rule is dk.HeightRule.EXACTLY
    # A stale source tuple must not become an authoritative state in a new copy.
    original = row._element.getElementsByTagNameNS(W, 'trHeight')[0]
    original.setAttributeNS(W, 'w:hRule', 'atLeast')
    assert row.clone(True).row_format.height_rule is dk.HeightRule.AUTO
