"""Text selection, boundary splitting, stale detection and failure-before-mutation."""

from io import BytesIO

import pytest
from defusedxml.ElementTree import fromstring
from test_docx_dom import package, payloads

import aspose.words_foss as aw
from aspose.words_foss.dom import TextRange
from aspose.words_foss.dom.nodes import W


def paragraph_doc(tmp_path, runs):
    source = tmp_path / "range.docx"
    package(source, '<w:p>' + runs + '</w:p><w:sectPr/>')
    return aw.DocxDocument(source)


@pytest.mark.parametrize("start,end,expected", [(0, 4, "甲😀乙丙"), (1, 2, "😀"), (2, 2, ""), (4, None, "")])
def test_range_uses_python_code_points(tmp_path, start, end, expected):
    doc = paragraph_doc(tmp_path, '<w:r><w:t>甲😀乙丙</w:t></w:r>')
    selected = doc.body.paragraphs[0].range(start, end)
    assert isinstance(selected, TextRange)
    assert selected.text == expected
    assert selected.paragraph is doc.body.paragraphs[0]


@pytest.mark.parametrize("start,end", [(-1, 0), (0, 5), (2, 1), (True, 2), (0, 1.5), ("0", 1)])
def test_invalid_bounds_do_not_mutate(tmp_path, start, end):
    doc = paragraph_doc(tmp_path, '<w:r><w:t>abcd</w:t></w:r>')
    original = payloads(doc.to_bytes())
    with pytest.raises(ValueError):
        doc.body.paragraphs[0].range(start, end)
    assert payloads(doc.to_bytes()) == original


def test_range_replace_refreshes_own_bounds_and_invalidates_other_ranges(tmp_path):
    doc = paragraph_doc(tmp_path, '<w:r><w:rPr><w:b/></w:rPr><w:t>客户</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>公司</w:t></w:r>')
    p = doc.body.paragraphs[0]
    selected = p.range(1, 3)
    other = p.range()
    assert selected.replace("新的").text == "新的"
    assert p.text == "客新的司"
    assert p.runs[0].text == "客新的" and p.runs[1].text == "司"
    assert p.runs[0].font.bold is True and p.runs[1].font.italic is True
    assert (selected.start, selected.end) == (1, 3)
    for action in (lambda: other.text, lambda: other.replace("x"), lambda: other.apply_font(bold=True)):
        with pytest.raises(ValueError, match="stale"):
            action()
    selected.replace("")
    assert selected.text == "" and selected.start == selected.end == 1
    selected.replace("插入")
    assert p.text == "客插入司"
    assert aw.DocxDocument(BytesIO(doc.to_bytes())).body.paragraphs[0].text == p.text


@pytest.mark.parametrize("position,expected_runs", [(0, ["Xab", "cd"]), (2, ["ab", "Xcd"]), (4, ["ab", "cdX"])])
def test_collapsed_insert_uses_right_run_or_last_run_at_end(tmp_path, position, expected_runs):
    doc = paragraph_doc(tmp_path, '<w:r><w:rPr><w:b/></w:rPr><w:t>ab</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>cd</w:t></w:r>')
    p = doc.body.paragraphs[0]
    selected = p.range(position, position)
    selected.replace("X")
    assert [run.text for run in p.runs] == expected_runs
    assert selected.text == "X"


@pytest.mark.parametrize("xml", ["", '<w:r><w:rPr><w:b/></w:rPr></w:r>', '<w:r><w:t/></w:r>'])
def test_empty_paragraph_or_run_accepts_insertion(tmp_path, xml):
    doc = paragraph_doc(tmp_path, xml)
    p = doc.body.paragraphs[0]
    assert p.range().replace("new").text == "new"
    assert p.text == "new"


def test_range_formatting_splits_only_boundaries_and_retains_other_properties(tmp_path):
    doc = paragraph_doc(tmp_path, '<w:r><w:rPr><w:i/><w:color w:val="FF0000"/><w:sz w:val="24"/></w:rPr><w:t>abcdef</w:t></w:r>')
    p = doc.body.paragraphs[0]
    selected = p.range(1, 5)
    original_run = p.runs[0]
    assert selected.apply_font(bold=True, italic=False, size=13.5) is selected
    assert selected.text == "bcde"
    assert [run.text for run in p.runs] == ["a", "bcde", "f"]
    assert p.runs[0] is original_run
    assert [run.font.bold for run in p.runs] == [None, True, None]
    assert [run.font.italic for run in p.runs] == [True, False, True]
    assert [run.font.size for run in p.runs] == [12, 13.5, 12]
    for run in p.runs:
        assert 'FF0000' in run.xml
    selected.apply_font(bold=None, size=None)
    assert selected.text == "bcde" and len(p.runs) == 3
    assert p.runs[1].font.bold is None and p.runs[1].font.size is None
    assert p.runs[1].effective_font.bold is False
    assert aw.DocxDocument(BytesIO(doc.to_bytes())).body.paragraphs[0].text == "abcdef"


def test_multirun_formatting_and_unrelated_edit_does_not_stale_range(tmp_path):
    doc = paragraph_doc(tmp_path, '<w:r><w:t>ab</w:t></w:r><w:r><w:t>cd</w:t></w:r><w:r><w:t>ef</w:t></w:r>')
    p = doc.body.paragraphs[0]
    selected = p.range(1, 5)
    doc.body.append_child(doc.create_paragraph("unrelated"))
    selected.apply_font(bold=True)
    assert [run.text for run in p.runs] == ["a", "b", "cd", "e", "f"]
    assert [run.font.bold for run in p.runs] == [None, True, True, True, None]
    p.runs[0].font.italic = True
    with pytest.raises(ValueError, match="stale"):
        selected.apply_font(bold=False)


@pytest.mark.parametrize("offset,left,right", [(1, "a", "bcd"), (2, "ab", "cd"), (3, "abc", "d")])
def test_split_multiple_text_nodes_preserves_comments_once(tmp_path, offset, left, right):
    doc = paragraph_doc(tmp_path, '<w:r><w:rPr><w:i/></w:rPr><w:t>a<!--inside-->b</w:t><!--between--><?keep instruction?><w:t>cd</w:t></w:r>')
    p = doc.body.paragraphs[0]
    original = p.runs[0]
    suffix = original.split(offset)
    assert original.text == left and suffix.text == right
    assert suffix.parent_node is p and suffix.font.italic is True
    assert p.xml.count('<!--inside-->') == p.xml.count('<!--between-->') == p.xml.count('<?keep instruction?>') == 1
    assert aw.DocxDocument(BytesIO(doc.to_bytes())).body.paragraphs[0].text == "abcd"


def test_split_edges_noops_and_space_preservation(tmp_path):
    doc = paragraph_doc(tmp_path, '<w:r><w:t>ab cd</w:t></w:r>')
    run = doc.body.paragraphs[0].runs[0]
    original = payloads(doc.to_bytes())
    assert run.split(0) is run
    assert run.split(5) is None
    assert payloads(doc.to_bytes()) == original
    right = run.split(2)
    assert right.text == " cd"
    root = fromstring(payloads(doc.to_bytes())["word/document.xml"])
    assert list(root.iter(f'{{{W}}}t'))[1].get('{http://www.w3.org/XML/1998/namespace}space') == "preserve"


@pytest.mark.parametrize("value", [0, -1, 12.25, float("nan"), float("inf"), True, "12"])
def test_invalid_size_rejected_before_range_split(tmp_path, value):
    doc = paragraph_doc(tmp_path, '<w:r><w:t>abcdef</w:t></w:r>')
    p = doc.body.paragraphs[0]
    original = payloads(doc.to_bytes())
    with pytest.raises((ValueError, TypeError)):
        p.range(1, 4).apply_font(bold=True, size=value)
    with pytest.raises((ValueError, TypeError)):
        p.runs[0].font.size = value
    assert payloads(doc.to_bytes()) == original
    assert len(p.runs) == 1


def test_unknown_metadata_rejected_before_any_boundary_is_split(tmp_path):
    doc = paragraph_doc(tmp_path, '<w:r><w:t>ab</w:t></w:r><w:r><w:t ext:flag="retain">cd</w:t></w:r>')
    p = doc.body.paragraphs[0]
    original = payloads(doc.to_bytes())
    with pytest.raises(NotImplementedError):
        p.range(1, 3).apply_font(bold=True)
    assert payloads(doc.to_bytes()) == original
    assert len(p.runs) == 2
    # Text replacement does not copy metadata and remains supported.
    p.range(1, 3).replace("X")
    assert p.text == "aXd"
    assert 'ext:flag="retain"' in p.xml


@pytest.mark.parametrize("inline", [
    '<w:hyperlink><w:r><w:t>link</w:t></w:r></w:hyperlink>',
    '<w:r><w:tab/></w:r>', '<w:r><w:br/></w:r>',
    '<w:bookmarkStart w:id="1" w:name="b"/>',
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
])
def test_complex_inline_content_is_not_flattened_into_editable_ranges(tmp_path, inline):
    doc = paragraph_doc(tmp_path, '<w:r><w:t>plain</w:t></w:r>' + inline)
    original = payloads(doc.to_bytes())
    with pytest.raises(NotImplementedError):
        doc.body.paragraphs[0].range(0, 1)
    assert payloads(doc.to_bytes()) == original


def test_range_noop_preserves_payload_and_collapsed_formatting_is_rejected(tmp_path):
    doc = paragraph_doc(tmp_path, '<w:r><w:t>abc</w:t></w:r>')
    original = payloads(doc.to_bytes())
    selected = doc.body.paragraphs[0].range(0, 3)
    selected.replace("abc").apply_font()
    assert payloads(doc.to_bytes()) == original
    with pytest.raises(ValueError):
        doc.body.paragraphs[0].range(1, 1).apply_font(bold=True)
    assert payloads(doc.to_bytes()) == original


def test_xml_ids_are_not_duplicated_by_range_splits(tmp_path):
    doc = paragraph_doc(tmp_path, '<w:r><w:t xml:id="unique">abcdef</w:t></w:r>')
    original = payloads(doc.to_bytes())
    with pytest.raises(NotImplementedError):
        doc.body.paragraphs[0].range(1, 3).apply_font(bold=True)
    assert payloads(doc.to_bytes()) == original


def test_namespace_context_survives_cloning_moving_and_new_node_insertion(tmp_path):
    source = tmp_path / "namespaces.docx"
    body = (f'<w:p><w:r><w:t>original</w:t></w:r></w:p>'
            f'<x:tbl xmlns:x="{W}" xmlns:w="urn:shadow"><x:tr><x:tc><x:p/></x:tc></x:tr></x:tbl>')
    package(source, body)
    doc = aw.DocxDocument(source)
    original = doc.body.paragraphs[0]
    cell = doc.body.tables[0].rows[0].cells[0]
    cell.append_child(original.clone())
    cell.append_child(original)
    cell.append_child(doc.create_paragraph("new"))
    reopened = aw.DocxDocument(BytesIO(doc.to_bytes()))
    paragraphs = reopened.body.tables[0].rows[0].cells[0].paragraphs
    assert [p.text for p in paragraphs] == ["", "original", "original", "new"]
    assert all(run.node_type == aw.NodeType.RUN for p in paragraphs for run in p.runs)


def test_invalid_split_offset_or_detached_parent_rejected(tmp_path):
    doc = paragraph_doc(tmp_path, '<w:r><w:t>abc</w:t></w:r>')
    run = doc.body.paragraphs[0].runs[0]
    original = payloads(doc.to_bytes())
    for offset in (-1, 4, True, 1.5):
        with pytest.raises(ValueError):
            run.split(offset)
    with pytest.raises(ValueError):
        doc.create_run("detached").split(1)
    assert payloads(doc.to_bytes()) == original
