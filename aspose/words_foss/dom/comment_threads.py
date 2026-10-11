"""Logical comment threads, independent of their serialized paragraph identifiers."""
from defusedxml.ElementTree import fromstring

from aspose.words_foss._comment_threads import (
    CONTENT_TYPE,
    RELATION,
    W14,
    W15,
    W,
    read_threads,
)
from aspose.words_foss._opc import resolve_target
from aspose.words_foss.dom.comment_utc import CT, REL, elements, set_attr


class CommentThreads:
    def __init__(self, graph):
        self.graph = graph
        self.parents = {}
        self.done = {}
        self.part = None
        package = graph.document._package
        name = 'word/_rels/document.xml.rels'
        links = elements(package.tree(name).documentElement, REL, 'Relationship') if name in package.part_names else []
        links = [link for link in links if link.getAttribute('Type') == RELATION]
        if len(links) > 1 or any(link.getAttribute('TargetMode') not in {'', 'Internal'} for link in links):
            raise ValueError('Expected one internal comment extension relationship')
        if not links:
            return
        self.part = resolve_target('word/document.xml', links[0].getAttribute('Target'))
        if self.part not in package.part_names or 'comments' not in graph.parts:
            raise ValueError('Comment extension has no matching story')
        comments = package.tree(graph.parts['comments']).documentElement
        values = read_threads(fromstring(comments.toxml(), forbid_dtd=True),
                              fromstring(package.tree(self.part).toxml(), forbid_dtd=True))
        for identifier, value in values.items():
            entry = graph.entries['comments', identifier]
            self.done[entry] = value['done']
            if value['ancestor_identifier'] is not None:
                self.parents[entry] = graph.entries['comments', value['ancestor_identifier']]

    def replies(self, entry):
        order = []
        documents = dict.fromkeys(anchor.ownerDocument for anchor in self.graph.anchors.values())
        for document in documents:
            for reference in document.getElementsByTagNameNS(W, 'commentReference'):
                child = self.graph.anchor_entries.get(reference.parentNode)
                if child is not None:
                    order.append(child)
        order.extend(child for (role, _), child in self.graph.entries.items() if role == 'comments' and child not in order)
        return [child for child in order if self.parents.get(child) is entry]

    def project(self, trees, mapping):
        if self.part is None and not self.done and not self.parents:
            return lambda: None
        from xml.dom.minidom import Document

        from aspose.words_foss.utils.xml_helpers import serialize_xml
        package = self.graph.document._package
        known = {entry for (role, _), entry in self.graph.entries.items() if role == 'comments'}
        entries = [(old, new) for old, new in mapping.items() if old in known and new.parentNode is not None]
        name = self.part
        if name is None:
            name = 'word/commentsExtended.xml'
            index = 0
            while name in package.part_names or name in trees:
                index += 1
                name = f'word/commentsExtended-{index}.xml'
        tree = Document()
        root = tree.createElementNS(W15, 'w15:commentsEx')
        root.setAttribute('xmlns:w15', W15)
        tree.appendChild(root)
        reserved = {paragraph.getAttributeNS(W14, 'paraId').upper() for _, new in entries
                    for paragraph in elements(new, W, 'p')}
        used, paragraphs, edits = set(), {}, []
        for old, new in entries:
            direct = elements(new, W, 'p')
            if not direct:
                raise ValueError('Comment requires a terminal paragraph')
            paragraph = direct[-1]
            raw = paragraph.getAttributeNS(W14, 'paraId').upper()
            if len(raw) != 8 or raw in used or any(c not in '0123456789ABCDEF' for c in raw):
                number = 1
                while f'{number:08X}' in used or f'{number:08X}' in reserved:
                    number += 1
                raw = f'{number:08X}'
            used.add(raw)
            set_attr(paragraph, W14, 'w14', 'paraId', raw)
            paragraphs[old] = raw
            source = elements(old, W, 'p')
            if source:
                edits.append((source[-1], raw))
        for old, _ in entries:
            node = tree.createElementNS(W15, 'w15:commentEx')
            set_attr(node, W15, 'w15', 'paraId', paragraphs[old])
            set_attr(node, W15, 'w15', 'done', '1' if self.done.get(old, False) else '0')
            parent = self.parents.get(old)
            if parent is not None:
                if parent not in paragraphs:
                    raise ValueError('Comment reply has no active ancestor')
                set_attr(node, W15, 'w15', 'paraIdParent', paragraphs[parent])
            root.appendChild(node)
        trees[name] = tree
        rel_name = 'word/_rels/document.xml.rels'
        if rel_name not in trees:
            trees[rel_name] = package.tree(rel_name).cloneNode(True)
        relations = trees[rel_name].documentElement
        for link in list(elements(relations, REL, 'Relationship')):
            if link.getAttribute('Type') == RELATION:
                relations.removeChild(link)
        identifiers = {link.getAttribute('Id') for link in elements(relations, REL, 'Relationship')}
        rid = 'comment-threads'
        while rid in identifiers:
            rid += '-1'
        link = relations.ownerDocument.createElementNS(REL, 'Relationship')
        link.setAttribute('xmlns', REL)
        for key, value in [('Id', rid), ('Type', RELATION), ('Target', name.removeprefix('word/'))]:
            link.setAttribute(key, value)
        relations.appendChild(link)
        if '[Content_Types].xml' not in trees:
            trees['[Content_Types].xml'] = package.tree('[Content_Types].xml').cloneNode(True)
        types = trees['[Content_Types].xml'].documentElement
        for item in list(elements(types, CT, 'Override')):
            if item.getAttribute('PartName') == '/' + name:
                types.removeChild(item)
        item = types.ownerDocument.createElementNS(CT, 'Override')
        item.setAttribute('xmlns', CT)
        item.setAttribute('PartName', '/' + name)
        item.setAttribute('ContentType', CONTENT_TYPE)
        types.appendChild(item)
        payloads = {n: serialize_xml(trees[n]) for n in (name, rel_name, '[Content_Types].xml')}
        def commit():
            package._commit_parts(payloads)
            self.part = name
            for paragraph, value in edits:
                set_attr(paragraph, W14, 'w14', 'paraId', value)
        return commit


class CommentReplies:
    """Live collection of visible replies, including pending API-created nodes."""
    def __init__(self, comment):
        self.comment = comment

    def _nodes(self):
        graph = self.comment.owner_document._annotations()
        return tuple(self.comment.owner_document._wrap(self.comment.part_name, entry)
                     for entry in graph.threads.replies(self.comment._element)
                     if entry in graph.anchors or entry in graph.pending)

    @property
    def count(self):
        return len(self)

    def __len__(self):
        return len(self._nodes())

    def __iter__(self):
        return iter(self._nodes())

    def __getitem__(self, index):
        return self._nodes()[index]

    def __eq__(self, other):
        return self._nodes() == other
