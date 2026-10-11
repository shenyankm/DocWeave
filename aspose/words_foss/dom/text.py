"""Word DOM text, including structural characters and logical inline stories."""
from xml.etree.ElementTree import Element

from aspose.words_foss.docx_reader.utils import _read_word_text
from aspose.words_foss.dom.nodes import XML, W, _elements, _find, _is, _text

_CONTROLS = {'tab': '\t', 'ptab': '', 'cr': '\r', 'noBreakHyphen': '\x1e',
             'softHyphen': '\x1f', 'footnoteRef': '\x02', 'endnoteRef': '\x02',
             'annotationRef': '\x05', 'separator': '\x03', 'continuationSeparator': '\x04'}
_EMPTY = {'bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd',
          'lastRenderedPageBreak', 'proofErr', 'permStart', 'permEnd', 'sectPr'}
_COMPOSITE = {'r', 'tbl', 'tr', 'tc', 'body', 'hdr', 'ftr', 'footnote', 'endnote',
              'comment', 'footnotes', 'endnotes', 'comments', 'ins', 'del', 'moveFrom', 'moveTo', 'txbxContent'}


def _paragraph_end(element):
    properties = _find(element, 'pPr')
    if properties is not None and _find(properties, 'sectPr') is not None:
        return '\x0c'
    parent = element.parentNode
    if parent is not None and (_is(parent, 'body') or _is(parent, 'tc')):
        children = [c for c in _elements(parent) if not _is(c, 'sectPr') and not _is(c, 'tcPr')]
        if children and element is children[-1]:
            return '\x0c' if _is(parent, 'body') else '\x07'
    return '\r'


def node_text(node, ancestors=()):
    element = node._element
    if element in ancestors:
        raise ValueError("Cyclic logical story reference")
    ancestors += (element,)
    if element.namespaceURI != W:
        raise NotImplementedError('Text extraction requires a supported Word node')
    name = element.localName
    if name in {'t', 'delText', 'instrText'}:
        if _elements(element):
            raise ValueError('Word text cannot contain nested elements')
        if name == 'instrText':
            return _text(element)
        attributes = {}
        space = element.getAttributeNS(XML, 'space') if element.hasAttributeNS(XML, 'space') else None
        root = element.ownerDocument.documentElement
        if _is(root, 'document') and root.getAttributeNS(XML, 'space') == 'preserve':
            space = 'preserve'
        if space is not None:
            attributes['{' + XML + '}space'] = space
        text = Element('t', attributes)
        text.text = _text(element)
        return _read_word_text(text)
    if name in _CONTROLS:
        return _CONTROLS[name]
    if name in _EMPTY:
        return ''
    if name == 'br':
        kind = element.getAttributeNS(W, 'type') or 'textWrapping'
        if kind not in {'textWrapping', 'page', 'column'}:
            raise ValueError('Invalid Word break type')
        return {'textWrapping': '\x0b', 'page': '\x0c', 'column': '\x0e'}[kind]
    if name == 'fldChar':
        kind = element.getAttributeNS(W, 'fldCharType')
        if kind not in {'begin', 'separate', 'end'}:
            raise ValueError('Invalid Word field delimiter')
        return {'begin': '\x13', 'separate': '\x14', 'end': '\x15'}[kind]
    if name == 'sym':
        try:
            return chr(int(element.getAttributeNS(W, 'char'), 16))
        except (ValueError, OverflowError):
            raise ValueError('Invalid Word symbol character') from None
    if name in {'drawing', 'pict'}:
        # Raster images have no text; only the explicitly embedded Word story contributes.
        return ''.join(node_text(node.owner_document._wrap(node.part_name, box), ancestors)
                       for box in element.getElementsByTagNameNS(W, 'txbxContent'))
    if name == 'fldSimple':
        return '\x13' + element.getAttributeNS(W, 'instr') + '\x14' + ''.join(node_text(c, ancestors) for c in node.child_nodes) + '\x15'
    if name == 'hyperlink':
        from aspose.words_foss._links import validate_hyperlink_history
        history = validate_hyperlink_history(element.getAttributeNS(W, 'history')
                    if element.hasAttributeNS(W, 'history') else None)
        target, separator, anchor = node.target.partition('#')
        instruction = ' HYPERLINK '
        if target:
            instruction += '"' + target + '" '
        if separator:
            instruction += '\\l "' + anchor + '" '
        for attribute, switch in [('tooltip', 'o'), ('tgtFrame', 't'), ('docLocation', 's')]:
            value = element.getAttributeNS(W, attribute)
            if value:
                instruction += '\\' + switch + ' "' + value + '" '
        if history not in {'1', 'true', 'on'}:
            instruction += ' \\h '
        return '\x13' + instruction + '\x14' + ''.join(node_text(c, ancestors) for c in node.child_nodes) + '\x15'
    if name == 'p' or name in _COMPOSITE:
        children = tuple(c for c in node.child_nodes if not _is(c._element, "sectPr"))
        value = ''.join(node_text(c, ancestors) for c in children)
        if name == 'p':
            return value + _paragraph_end(element)
        if name == 'tr':
            return value + '\x07'
        if name in {'body', 'tc'} and (not children or not _is(children[-1]._element, 'p')):
            return value + ('\x0c' if name == 'body' else '\x07')
        return value
    raise NotImplementedError('Text extraction for this Word structure is not supported')


def _header_footer_text(document, properties):
    from aspose.words_foss._opc import resolve_target
    from aspose.words_foss.docx_writer.constants import PKG_RELS_URI, R_URI
    references = [e for e in _elements(properties) if _is(e, 'headerReference') or _is(e, 'footerReference')]
    if not references:
        return ''
    package = document._package
    package.story_part_names()
    links = {e.getAttribute('Id'): e for e in package.tree('word/_rels/document.xml.rels').getElementsByTagNameNS(PKG_RELS_URI, 'Relationship')}
    stories = {}
    for reference in references:
        role = 'header' if _is(reference, 'headerReference') else 'footer'
        variant = reference.getAttributeNS(W, 'type')
        key = role, variant
        link = links.get(reference.getAttributeNS(R_URI, 'id'))
        if key in stories or variant not in {'even', 'default', 'first'} or link is None:
            raise ValueError('Invalid header/footer reference')
        if link.getAttribute('TargetMode') not in {'', 'Internal'} or link.getAttribute('Type') != 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + role:
            raise ValueError('Invalid header/footer relationship')
        name = resolve_target('word/document.xml', link.getAttribute('Target'))
        stories[key] = document.story(name)
    return ''.join(stories[role, variant].get_text() for role in ['footer', 'header']
                   for variant in ['even', 'default', 'first'] if (role, variant) in stories)


def document_text(document):
    """Sections contribute their body followed by their explicit header/footer nodes."""
    result, children = [], []
    for node in document.body.child_nodes:
        element = node._element
        properties = None
        if _is(element, 'sectPr'):
            properties = element
        else:
            children.append(node)
            paragraph_properties = _find(element, 'pPr') if _is(element, 'p') else None
            if paragraph_properties is not None:
                properties = _find(paragraph_properties, 'sectPr')
        if properties is not None:
            value = ''.join(child.get_text() for child in children)
            if not children or not _is(children[-1]._element, 'p'):
                value += '\x0c'
            result.append(value + _header_footer_text(document, properties))
            children = []
    if children:
        value = ''.join(child.get_text() for child in children)
        if not _is(children[-1]._element, 'p'):
            value += '\x0c'
        result.append(value)
    return ''.join(result) if result else '\x0c'
