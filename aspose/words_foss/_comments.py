"""Classic comment stories and ordered anchors in conversion models."""
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.diagnostics import document_nodes


def _comment_story_nodes(children):
    # LDM lists can be edited without DOM ownership checks. Reject a back edge,
    # but allow the same node to be used by independent, non-cyclic branches.
    active, complete = set(), set()
    stack = [(node, False) for node in reversed(children)]
    while stack:
        node, leaving = stack.pop()
        key = id(node)
        if leaving:
            active.remove(key)
            complete.add(key)
            continue
        if key in active:
            raise ValueError("Cyclic comment story")
        if key in complete:
            continue
        active.add(key)
        stack.append((node, True))
        if isinstance(node, ldm.Paragraph):
            nested = node._children
        elif isinstance(node, ldm.Shape):
            nested = (node.text_box or {}).get("paragraphs") or []
        elif isinstance(node, ldm.Table):
            nested = node.rows
        elif isinstance(node, ldm.Row):
            nested = node.cells
        elif isinstance(node, ldm.Cell):
            nested = node.children
        else:
            nested = []
        stack.extend((child, False) for child in reversed(nested))
        yield node


def validate_comments(document):
    from aspose.words_foss.dom.notes import annotation_id
    entries = {}
    for story in document.source_stories:
        if story.kind != "comment":
            continue
        identifier = str(annotation_id(story.identifier or ""))
        if identifier in entries:
            raise ValueError("Duplicate comment identifier")
        entries[identifier] = story
        for paragraph in _comment_story_nodes(story.children):
            if isinstance(paragraph, ldm.Paragraph) and any(isinstance(n, ldm.CommentMarker) and n.kind != "annotation"
                                                           for n in paragraph._children):
                raise ValueError("Comment story contains an external anchor")
        for child in story.children:
            if not isinstance(child, (ldm.Paragraph, ldm.Table)):
                raise NotImplementedError("Unsupported comment story node")
    from aspose.words_foss._comment_threads import validate_threads
    values = {identifier: story.comment_metadata.model_dump() if story.comment_metadata else {}
              for identifier, story in entries.items()}
    validate_threads(values)
    for value in values.values():
        if value.get("done") is not None and not isinstance(value["done"], bool):
            raise ValueError("Invalid comment completion state")
        if value.get("ancestor_identifier") is not None and value["ancestor_identifier"] not in entries:
            raise ValueError("Comment reply has no matching ancestor")
    markers = {}
    seen = set()
    for paragraph in document_nodes(document):
        if not isinstance(paragraph, ldm.Paragraph) or id(paragraph) in seen:
            continue
        seen.add(id(paragraph))
        for node in paragraph._children:
            if not isinstance(node, ldm.CommentMarker):
                continue
            if node.kind == "annotation":
                raise ValueError("Comment annotation mark outside a comment story")
            identifier = str(annotation_id(node.identifier))
            if identifier not in entries:
                raise ValueError("Comment anchor has no matching story")
            kinds = markers.setdefault(identifier, [])
            if node.kind in kinds:
                raise ValueError("Duplicate comment anchor")
            kinds.append(node.kind)
    for kinds in markers.values():
        if "reference" not in kinds or ("start" in kinds) != ("end" in kinds):
            raise ValueError("Incomplete comment anchor range")
        if "start" in kinds and kinds.index("start") > kinds.index("end"):
            raise ValueError("Reversed comment anchor range")
    return entries


def render_comments(document, style_pf_map, style_id_map, style_font_map):
    from aspose.words_foss.docx_writer.bookmarks import BookmarkState
    from aspose.words_foss.docx_writer.drawing import ImageRenderState
    from aspose.words_foss.docx_writer.paragraphs import render_paragraph
    from aspose.words_foss.docx_writer.tables import render_table
    from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el, _is_xml_char
    from aspose.words_foss.docx_writer.constants import W_URI, R_URI, MC_URI, V_URI, O_URI
    from aspose.words_foss.docx_writer.drawing import WP_URI, A_URI, PIC_URI, WPS_URI
    entries = validate_comments(document)
    if not entries:
        return None, {}, []
    state = ImageRenderState(media_name_prefix="comment_image", rel_id_prefix="rIdCommentImg", doc_pr_offset=300000)
    hyperlinks = {}
    bookmarks = BookmarkState()
    comments = []
    for identifier, story in entries.items():
        attributes = {"w:id": identifier}
        if story.comment_metadata is not None:
            for field, tag in [("author", "author"), ("initial", "initials"), ("date", "date")]:
                value = getattr(story.comment_metadata, field)
                if value is not None:
                    if not isinstance(value, str) or not all(_is_xml_char(c) for c in value):
                        raise ValueError("Invalid comment metadata")
                    attributes["w:" + tag] = value
        children = []
        for child in story.children:
            render = render_paragraph if isinstance(child, ldm.Paragraph) else render_table
            children.append(render(child, hyperlinks, image_state=state, bookmark_state=bookmarks,
                                   style_pf_map=style_pf_map, style_id_map=style_id_map, style_font_map=style_font_map))
        if not story.children or not isinstance(story.children[-1], ldm.Paragraph):
            children.append(el("w:p"))
        comments.append(el("w:comment", attributes, children))
    result = XML_DECL + el("w:comments", {"xmlns:w": W_URI, "xmlns:r": R_URI, "xmlns:wp": WP_URI,
                                         "xmlns:a": A_URI, "xmlns:pic": PIC_URI, "xmlns:wps": WPS_URI,
                                         "xmlns:mc": MC_URI, "xmlns:v": V_URI, "xmlns:o": O_URI, "mc:Ignorable": "wps"}, comments)
    from defusedxml.ElementTree import fromstring
    from xml.etree.ElementTree import ParseError
    try:
        fromstring(result, forbid_dtd=True)
    except ParseError:
        raise ValueError("Invalid comment metadata or content") from None
    return result, hyperlinks, state.images


def validate_source_anchors(reader, document):
    """Fail clearly if a recognized source anchor was skipped by an inline caller."""
    from collections import Counter
    from aspose.words_foss.dom.notes import annotation_id
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    kinds = {namespace + tag: kind for tag, kind in [("commentRangeStart", "start"), ("commentRangeEnd", "end"),
                                                    ("commentReference", "reference"), ("annotationRef", "annotation")]}
    expected = Counter()
    for root in [reader._document_xml] + [entry[3] for entry in reader._source_story_data]:
        for node in root.iter():
            if node.tag in kinds:
                kind = kinds[node.tag]
                expected[kind, "" if kind == "annotation" else str(annotation_id(node.get(namespace + "id", "")))] += 1
    observed = Counter()
    seen = set()
    stack = [document] + [child for story in document.source_stories for child in story.children]
    while stack:
        node = stack.pop()
        if id(node) in seen:
            continue
        seen.add(id(node))
        if isinstance(node, ldm.Paragraph):
            for marker in node._children:
                if isinstance(marker, ldm.CommentMarker):
                    observed[marker.kind, marker.identifier] += 1
        stack.extend(ldm._walk_children(node))
        if isinstance(node, ldm.Shape) and node.text_box:
            stack.extend(node.text_box.get("paragraphs", []) or [])
    if expected != observed:
        raise NotImplementedError("Unsupported comment anchor location")
