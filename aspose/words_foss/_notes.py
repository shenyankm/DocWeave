"""Ordered note stories, inline marks, and WordprocessingML note options."""
import math
import re

from aspose.words_foss import light_document_model as ldm


def note_id(value):
    """Normalize the numeric ID recovery used by the calibrated native loader."""
    if not isinstance(value, str):
        raise ValueError("Invalid note identifier")
    if not re.fullmatch(r'[ \t\r\n]*[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?[ \t\r\n]*', value):
        return '0'
    number = float(value)
    if not math.isfinite(number):
        return '0'
    integer = min(max(int(number), -(1 << 63)), (1 << 63) - 1)
    return str((integer + (1 << 31)) % (1 << 32) - (1 << 31))


def read_properties(root):
    from aspose.words_foss.docx_reader.constants import W_NS
    result = []
    if root is None:
        return result
    for node in root:
        if node.tag not in {W_NS + 'footnotePr', W_NS + 'endnotePr'}:
            continue
        kind = node.tag[len(W_NS):-2]
        if node.attrib:
            raise NotImplementedError("Unsupported note property attributes")
        children = []
        for child in node:
            name = child.tag[len(W_NS):] if child.tag.startswith(W_NS) else ''
            if name not in {'pos', 'numFmt', 'numStart', 'numRestart', kind} or len(child):
                raise NotImplementedError("Unsupported note property")
            attribute = 'id' if name == kind else 'val'
            allowed = {W_NS + attribute} | ({W_NS + 'format'} if name == 'numFmt' else set())
            if set(child.attrib) - allowed:
                raise NotImplementedError("Unsupported note property attributes")
            value = child.get(W_NS + attribute)
            if value is None:
                raise ValueError("Missing note property value")
            if name == kind:
                value = note_id(value)
            children.append(ldm.NoteProperty(kind=name, value=value, format=child.get(W_NS + 'format')))
        result.append(ldm.NoteProperties(kind=kind, children=children))
    return result


def render_properties(properties):
    from aspose.words_foss.docx_writer.xml_utils import el, _is_xml_char
    rendered, seen = [], set()
    for prop in properties:
        if prop.kind in seen:
            raise ValueError("Duplicate note properties")
        seen.add(prop.kind)
        children = []
        for child in prop.children:
            if not all(_is_xml_char(c) for c in child.value) or (
                    child.format is not None and not all(_is_xml_char(c) for c in child.format)):
                raise ValueError("Invalid note property text")
            if child.kind in {'footnote', 'endnote'} and child.kind != prop.kind:
                raise ValueError("Note property role disagrees with its container")
            attributes = {'w:id' if child.kind == prop.kind else 'w:val':
                          note_id(child.value) if child.kind == prop.kind else child.value}
            if child.format is not None:
                if child.kind != 'numFmt':
                    raise ValueError("Unexpected note number format")
                attributes['w:format'] = child.format
            children.append(el('w:' + child.kind, attributes))
        rendered.append(el('w:' + prop.kind + 'Pr', None, children))
    return rendered


def validate_notes(document):
    from aspose.words_foss._comments import _comment_story_nodes
    from aspose.words_foss.diagnostics import document_nodes
    entries = {}
    for story in document.source_stories:
        if story.kind not in {'footnote', 'endnote'}:
            continue
        if story.identifier is None or story.identifier == '':
            raise ValueError("Missing note story identifier")
        key = story.kind, note_id(story.identifier)
        if key in entries:
            raise ValueError("Duplicate note identifier")
        entries[key] = story
        if any(not isinstance(child, (ldm.Paragraph, ldm.Table)) for child in story.children):
            raise NotImplementedError("Unsupported note story node")
        for node in _comment_story_nodes(story.children):
            if isinstance(node, ldm.UnknownNode):
                raise NotImplementedError("Unsupported note story node")
            if isinstance(node, ldm.NoteReference):
                raise ValueError("Note story contains an external note anchor")
    for paragraph in document_nodes(document):
        if not isinstance(paragraph, ldm.Paragraph):
            continue
        anchors = [node for node in paragraph._children if isinstance(node, ldm.NoteReference)]
        if len(paragraph.note_references) != len(anchors) or any(
                (anchor.kind, note_id(anchor.identifier)) != (old.kind, note_id(old.identifier))
                for old, anchor in zip(paragraph.note_references, anchors)):
            raise ValueError("Note reference has no ordered inline anchor")
        if any(old.model_dump(exclude={"identifier"}) != anchor.model_dump(exclude={"identifier"})
               for old, anchor in zip(paragraph.note_references, anchors)):
            raise ValueError("Note reference metadata disagrees with ordered anchor")
        for node in anchors:
            if node.reference_mark is not None:
                from aspose.words_foss.docx_writer.xml_utils import _is_xml_char
                if not node.custom_mark_follows or not all(_is_xml_char(c) for c in node.reference_mark):
                    raise ValueError("Invalid custom note reference mark")
            story = entries.get((node.kind, note_id(node.identifier)))
            if story is None or story.note_type != 'normal':
                raise ValueError("Note anchor has no matching normal story")
    for properties in [document.note_properties, *[s.page_setup.note_properties for s in document.sections]]:
        render_properties(properties)
        for prop in properties:
            for child in prop.children:
                if child.kind == prop.kind:
                    story = entries.get((prop.kind, note_id(child.value)))
                    if story is None or story.note_type == 'normal':
                        raise ValueError("Note separator property has no matching special story")
    return entries


def render_notes(document, style_pf_map, style_id_map, style_font_map):
    from aspose.words_foss.docx_writer.bookmarks import BookmarkState
    from aspose.words_foss.docx_writer.drawing import ImageRenderState, WP_URI, A_URI, PIC_URI, WPS_URI
    from aspose.words_foss.docx_writer.constants import W_URI, R_URI, MC_URI, V_URI, O_URI
    from aspose.words_foss.docx_writer.numbering_part import build_num_id_map
    from aspose.words_foss.docx_writer.paragraphs import render_paragraph
    from aspose.words_foss.docx_writer.tables import render_table
    from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el
    entries = validate_notes(document)
    parts = []
    for index, kind in enumerate(('footnote', 'endnote')):
        stories = [(identifier, story) for (role, identifier), story in entries.items() if role == kind]
        if not stories:
            continue
        state = ImageRenderState(media_name_prefix=kind + '_image', rel_id_prefix='rId' + kind + 'Img',
                                 doc_pr_offset=400000 + index * 100000)
        links, children = {}, []
        bookmarks = BookmarkState()
        for identifier, story in stories:
            attributes = {'w:id': identifier}
            if story.note_type != 'normal':
                attributes['w:type'] = story.note_type
            blocks = []
            for child in story.children:
                render = render_paragraph if isinstance(child, ldm.Paragraph) else render_table
                blocks.append(render(child, links, num_id_map=build_num_id_map(document.lists),
                                     image_state=state, bookmark_state=bookmarks,
                                     style_pf_map=style_pf_map, style_id_map=style_id_map, style_font_map=style_font_map))
            for _, bookmark_id in bookmarks.drain_open():
                blocks.append(el('w:bookmarkEnd', {'w:id': bookmark_id}))
            if not story.children or not isinstance(story.children[-1], ldm.Paragraph):
                blocks.append(el('w:p'))
            children.append(el('w:' + kind, attributes, blocks))
        xml = XML_DECL + el('w:' + kind + 's', {'xmlns:w': W_URI, 'xmlns:r': R_URI, 'xmlns:wp': WP_URI,
                'xmlns:a': A_URI, 'xmlns:pic': PIC_URI, 'xmlns:wps': WPS_URI, 'xmlns:mc': MC_URI,
                'xmlns:v': V_URI, 'xmlns:o': O_URI, 'mc:Ignorable': 'wps'}, children)
        from defusedxml.ElementTree import fromstring
        from xml.etree.ElementTree import ParseError
        try:
            fromstring(xml, forbid_dtd=True)
        except ParseError:
            raise ValueError("Invalid note story content") from None
        parts.append((kind + 's.xml', 'rId' + kind + 's', kind + 's', xml, links, state.images))
    return parts


def validate_source_notes(reader, document):
    """Reject recognized note marks when an inline reader skipped their location."""
    from collections import Counter
    from aspose.words_foss._comments import _comment_story_nodes
    from aspose.words_foss.diagnostics import document_nodes
    from aspose.words_foss.docx_reader.constants import W_NS
    tags = {'footnoteReference', 'endnoteReference', 'footnoteRef', 'endnoteRef', 'separator', 'continuationSeparator'}
    expected = Counter()
    for root in [reader._document_xml] + [entry[3] for entry in reader._source_story_data]:
        for item in root.iter():
            name = item.tag[len(W_NS):] if item.tag.startswith(W_NS) else ''
            if name in tags:
                expected[name, note_id(item.get(W_NS + 'id', '')) if name.endswith('Reference') else ''] += 1
    observed, seen = Counter(), set()
    nodes = list(document_nodes(document))
    for story in document.source_stories:
        nodes.extend(_comment_story_nodes(story.children))
    for paragraph in nodes:
        if not isinstance(paragraph, ldm.Paragraph) or id(paragraph) in seen:
            continue
        seen.add(id(paragraph))
        for marker in paragraph._children:
            if isinstance(marker, ldm.NoteReference):
                observed[marker.kind + 'Reference', note_id(marker.identifier)] += 1
            elif isinstance(marker, ldm.NoteMarker):
                observed[marker.kind, ''] += 1
    if expected != observed:
        raise NotImplementedError("Unsupported note mark location")
