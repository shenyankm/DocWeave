"""Typed comment reply links and completion state (MS-DOCX commentEx)."""
import re

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
RELATION = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended'
CONTENT_TYPE = 'application/vnd.ms-word.commentsExtended+xml'


def _paragraph_id(value):
    if not re.fullmatch('[0-9a-fA-F]{8}', value or ''):
        raise ValueError('Invalid comment paragraph identifier')
    return value.upper()


def read_threads(comments, extended):
    """Return comment-ID keyed values; do not retain an opaque metadata part."""
    if extended is None:
        return {}
    if extended.tag != '{' + W15 + '}commentsEx':
        raise ValueError('Expected comment extension metadata')
    by_paragraph = {}
    for entry in comments:
        if entry.tag != '{' + W + '}comment':
            continue
        paragraphs = entry.findall('{' + W + '}p')
        if not paragraphs:
            continue
        raw = paragraphs[-1].get('{' + W14 + '}paraId')
        if raw is None:
            continue
        identifier = _paragraph_id(raw)
        if identifier in by_paragraph:
            raise ValueError('Duplicate comment paragraph identifier')
        from aspose.words_foss.dom.notes import annotation_id
        by_paragraph[identifier] = str(annotation_id(entry.get('{' + W + '}id', '')))
    result = {}
    for record in extended:
        if len(record) or (record.text or '').strip():
            raise NotImplementedError('Unsupported comment extension content')
        if record.tag != '{' + W15 + '}commentEx':
            raise NotImplementedError('Unsupported comment extension metadata')
        if any(key not in {'{' + W15 + '}' + name for name in ('paraId', 'paraIdParent', 'done')} for key in record.attrib):
            raise NotImplementedError('Unsupported comment extension attribute')
        identifier = by_paragraph.get(_paragraph_id(record.get('{' + W15 + '}paraId')))
        if identifier is None:
            raise ValueError('Comment extension has no matching story')
        if identifier in result:
            raise ValueError('Duplicate comment extension')
        parent_raw = record.get('{' + W15 + '}paraIdParent')
        parent = by_paragraph.get(_paragraph_id(parent_raw)) if parent_raw is not None else None
        if parent_raw is not None and parent is None:
            raise ValueError('Comment reply has no matching ancestor')
        raw_done = record.get('{' + W15 + '}done', '0')
        # Fixed 26.9 uses exact "0" as false; absent defaults to zero.
        # Even the schema-valid lexical values "false" and "off" load as true.
        result[identifier] = {'ancestor_identifier': parent, 'done': raw_done != '0'}
    validate_threads(result)
    return result


def validate_threads(values):
    for identifier, value in values.items():
        parent = value.get('ancestor_identifier')
        if parent is not None and (parent == identifier or values.get(parent, {}).get('ancestor_identifier') is not None):
            raise ValueError('Nested or cyclic comment reply')


def extension_xml(comments_xml, values):
    """Regenerate identifiers without changing namespace-bearing story content."""
    from xml.dom.minidom import Document

    from defusedxml.minidom import parseString

    from aspose.words_foss.dom.comment_utc import elements, set_attr
    from aspose.words_foss.utils.xml_helpers import serialize_xml
    comments = parseString(comments_xml, forbid_dtd=True)
    entries = {entry.getAttributeNS(W, 'id'): entry for entry in elements(comments.documentElement, W, 'comment')}
    validate_threads(values)
    paragraphs = {}
    for index, (identifier, entry) in enumerate(entries.items(), 1):
        children = [node for node in entry.childNodes if node.nodeType == node.ELEMENT_NODE]
        if not children or children[-1].namespaceURI != W or children[-1].localName != 'p':
            entry.appendChild(comments.createElementNS(W, 'w:p'))
        paragraph_id = f'{index:08X}'
        set_attr(elements(entry, W, 'p')[-1], W14, 'w14', 'paraId', paragraph_id)
        paragraphs[identifier] = paragraph_id
    tree = Document()
    root = tree.createElementNS(W15, 'w15:commentsEx')
    root.setAttribute('xmlns:w15', W15)
    tree.appendChild(root)
    for identifier in entries:
        value = values.get(identifier, {})
        record = tree.createElementNS(W15, 'w15:commentEx')
        set_attr(record, W15, 'w15', 'paraId', paragraphs[identifier])
        set_attr(record, W15, 'w15', 'done', '1' if value.get('done') else '0')
        parent = value.get('ancestor_identifier')
        if parent is not None:
            if parent not in paragraphs:
                raise ValueError('Comment reply has no matching ancestor')
            set_attr(record, W15, 'w15', 'paraIdParent', paragraphs[parent])
        root.appendChild(record)
    return serialize_xml(comments).decode(), serialize_xml(tree).decode()
