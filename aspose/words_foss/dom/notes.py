"""Logical inline annotations backed by their typed OPC content parts."""
from enum import IntEnum
import re

from aspose.words_foss.dom.nodes import W, _elements, _is, _find, _safe_structure, _new
from aspose.words_foss.dom.importing import _set_word_attribute

_ROLES = {"footnotes": "footnote", "endnotes": "endnote", "comments": "comment"}
_REFERENCES = {"footnoteReference": "footnotes", "endnoteReference": "endnotes",
               "commentReference": "comments"}


def annotation_id(value):
    """Compare decimal XML IDs without exposing malformed source content."""
    if not re.fullmatch(r'[+-]?[0-9]+', value.strip()):
        raise ValueError('Invalid annotation identifier')
    return int(value)


class FootnoteType(IntEnum):
    FOOTNOTE = 0
    ENDNOTE = 1


class NoteGraph:
    def __init__(self, document):
        self.document = document
        self.parts = {}
        self.entries = {}
        self.duplicate_entries = []
        self.anchors = {}
        self.comment_ids = {}
        self.anchor_entries = {}
        self.pending = {}
        self.templates = {}
        self.pending_trees = {}
        self.changed = False
        document._package._annotation_normalized_parts = set()
        for name in sorted(document._package.story_part_names()):
            root = document._package.tree(name).documentElement
            if root.localName not in _ROLES or root.namespaceURI != W:
                continue
            role = root.localName
            if role in self.parts:
                raise ValueError("Ambiguous annotation story relationship")
            self.parts[role] = name
            for entry in _elements(root):
                if not _is(entry, _ROLES[role]) or entry.getAttributeNS(W, "type") not in {"", "normal"}:
                    continue
                identifier = str(annotation_id(entry.getAttributeNS(W, "id")))
                key = role, identifier
                if key in self.entries:
                    self.duplicate_entries.append(entry)
                    self.changed = True
                    continue
                self.entries[key] = entry
                children = _elements(entry)
                if not children or not _is(children[-1], 'p'):
                    entry.appendChild(_new(entry, 'p'))
                    document._package._annotation_normalized_parts.add(name)
        from aspose.words_foss.dom.comment_utc import CommentUtc
        self.utc = CommentUtc(document._package, self.entries)
        named_trees = [(name, document._package.tree(name)) for name in sorted(document._package.story_part_names(), key=lambda name: (name != 'word/document.xml', name))]
        trees = [tree for name, tree in named_trees]
        # Split only recognized story content; unchanged OPC bytes remain the source.
        for name, tree in named_trees:
            for paragraph in list(tree.getElementsByTagNameNS(W, 'p')):
                if self._story_paragraph(paragraph) and self._split_runs(paragraph):
                    document._package._annotation_normalized_parts.add(name)
        for tree in trees:
            for paragraph in tree.getElementsByTagNameNS(W, 'p'):
                if not self._story_paragraph(paragraph):
                    continue
                for run in _elements(paragraph):
                    if not _is(run, 'r'):
                        continue
                    references = [child for child in _elements(run)
                                  if child.namespaceURI == W and child.localName in _REFERENCES]
                    if len(references) != 1:
                        continue
                    reference = references[0]
                    content = [item for item in _elements(run) if not _is(item, 'rPr')]
                    if (any(item is not reference for item in content)
                            and not (reference.localName in {'footnoteReference', 'endnoteReference'}
                                     and reference.getAttributeNS(W, 'customMarkFollows') in {'1', 'true', 'on'}
                                     and all(item is reference or _is(item, 't') for item in content))):
                        raise NotImplementedError('Mixed annotation run requires metadata preservation')
                    role = _REFERENCES[reference.localName]
                    entry = self.entries.get((role, str(annotation_id(reference.getAttributeNS(W, 'id')))))
                    if entry is None:
                        continue
                    if entry in self.anchors:
                        raise ValueError('An annotation requires one anchor')
                    self.anchors[entry] = run
                    self.anchor_entries[run] = entry
                    if role == 'comments':
                        self.comment_ids[entry] = len(self.comment_ids)
        for tree in trees:
            for name in ['commentRangeStart', 'commentRangeEnd']:
                for marker in tree.getElementsByTagNameNS(W, name):
                    if self._story_paragraph(marker.parentNode):
                        identifier = str(annotation_id(marker.getAttributeNS(W, 'id')))
                        owner = self.entries.get(('comments', identifier))
                        if owner in self.comment_ids:
                            marker._logical_comment_id = self.comment_ids[owner]

        from aspose.words_foss.dom.comment_threads import CommentThreads
        self.threads = CommentThreads(self)
        self._next_comment_id = len(self.comment_ids)
        self.changed |= bool(self.threads.parents)

    @staticmethod
    def _story_paragraph(paragraph):
        root = paragraph.ownerDocument.documentElement
        entries = {'document': 'body', 'hdr': None, 'ftr': None, **_ROLES}
        if root.namespaceURI != W or root.localName not in entries:
            return False
        entry = entries[root.localName]
        allowed = {'tc', 'tr', 'tbl', entry}
        found_entry = entry is None
        parent = paragraph.parentNode
        while parent is not root:
            if parent is None or parent.nodeType != parent.ELEMENT_NODE:
                return False
            if parent.namespaceURI != W or parent.localName not in allowed:
                return False
            found_entry |= parent.localName == entry
            parent = parent.parentNode
        return found_entry

    @staticmethod
    def _main_paragraph(paragraph):
        root = paragraph.ownerDocument.documentElement
        return _is(root, 'document') and NoteGraph._story_paragraph(paragraph)

    def _split_runs(self, paragraph):
        allowed = {'rPr', 't', 'tab', 'br', 'cr', *_REFERENCES}
        changed = False
        for run in list(_elements(paragraph)):
            if not _is(run, 'r'):
                continue
            children = _elements(run)
            if (not any(child.localName in _REFERENCES and child.namespaceURI == W for child in children)
                    or any(child.namespaceURI != W or child.localName not in allowed for child in children)
                    or not _safe_structure(run, properties=True)):
                continue
            content = [child for child in children if not _is(child, 'rPr')]
            if len(content) <= 1:
                continue
            groups, leading = [], []
            for child in run.childNodes:
                if child.nodeType != child.ELEMENT_NODE:
                    (groups[-1] if groups else leading).append(child)
                    continue
                if _is(child, 'rPr'):
                    continue
                if child.localName in _REFERENCES:
                    groups.append([child])
                elif groups and next(item for item in groups[-1] if item.nodeType == item.ELEMENT_NODE).localName not in _REFERENCES:
                    groups[-1].append(child)
                else:
                    groups.append([child])
            if leading:
                groups[0][:0] = leading
            # A custom mark's following text is part of its logical anchor run.
            for index in range(len(groups) - 2, -1, -1):
                first = next(item for item in groups[index] if item.nodeType == item.ELEMENT_NODE)
                following = [item for item in groups[index + 1] if item.nodeType == item.ELEMENT_NODE]
                if (first.localName in {'footnoteReference', 'endnoteReference'}
                        and first.getAttributeNS(W, 'customMarkFollows') in {'1', 'true', 'on'}
                        and len(following) == 1 and _is(following[0], 't')):
                    groups[index].extend(groups.pop(index + 1))
            if len(groups) == 1:
                continue
            formatting = _find(run, 'rPr')
            for group in groups:
                split = run.cloneNode(deep=False)
                if formatting is not None:
                    split.appendChild(formatting.cloneNode(deep=True))
                for child in group:
                    split.appendChild(child.cloneNode(deep=True))
                paragraph.insertBefore(split, run)
            paragraph.removeChild(run)
            changed = True
        return changed

    def node_for_run(self, run):
        entry = self.anchor_entries.get(run)
        if entry is None:
            return None
        role = entry.localName + 's'
        name = self.parts.get(role)
        if name is None:
            name = self.pending_trees[role][0]
        return self.document._wrap(name, entry)

    def contained_entries(self, element):
        return [self.anchor_entries[run] for run in element.getElementsByTagNameNS(W, 'r')
                if run in self.anchor_entries]

    def safe_content(self, element):
        if element.namespaceURI == W and element.localName in _REFERENCES:
            return element.parentNode in self.anchor_entries
        if _safe_structure(element):
            return True
        # Only recognized annotation references relax ordinary structural guards.
        shell = element.cloneNode(deep=False)
        return (_safe_structure(shell) and
                all(self.safe_content(child) for child in _elements(element)))

    def clone_container(self, node, deep):
        from aspose.words_foss.dom.nodes import _bind_namespace_context, _copy_row_warm
        if not self.safe_content(node._element):
            raise NotImplementedError('Cloning complex annotation content requires resource translation')
        entries = self.contained_entries(node._element)
        if any(not all(_safe_structure(child) for child in _elements(entry)) for entry in entries):
            raise NotImplementedError('Cloning annotation content requires resource translation')
        copied = node._element.cloneNode(deep=True)
        if not deep:
            from aspose.words_foss.dom.nodes import _PROPERTY_ROOTS
            for child in list(copied.childNodes):
                if child.nodeType != child.ELEMENT_NODE or child.localName not in _PROPERTY_ROOTS | {'tblGrid'}:
                    copied.removeChild(child)
        _bind_namespace_context(node._element, copied)
        _copy_row_warm(node._element, copied)
        range_mapping = {}
        if deep:
            original_runs = node._element.getElementsByTagNameNS(W, 'r')
            copied_runs = copied.getElementsByTagNameNS(W, 'r')
            for old_run, new_run in zip(original_runs, copied_runs):
                entry = self.anchor_entries.get(old_run)
                if entry is None:
                    continue
                old_node = self.node_for_run(old_run)
                new_node = self.clone(old_node, True)
                role = entry.localName + 's'
                if role == 'comments':
                    range_mapping.setdefault(old_node.id, self._next_comment_id)
                    self.comment_ids[new_node._element] = self._next_comment_id
                    self._next_comment_id += 1
                existing = [int(key[1]) for key in self.entries if key[0] == role]
                identifier = str(max(existing, default=1 if role != 'comments' else -1) + 1)
                _set_word_attribute(new_node._element, 'id', identifier)
                reference = next(child for child in _elements(new_run) if child.localName in _REFERENCES)
                _set_word_attribute(reference, 'id', identifier)
                self.entries[role, identifier] = new_node._element
                self.anchors[new_node._element] = new_run
                self.anchor_entries[new_run] = new_node._element
            for old, new in zip(node._element.getElementsByTagNameNS(W, '*'), copied.getElementsByTagNameNS(W, '*')):
                if hasattr(old, '_logical_comment_id'):
                    new._logical_comment_id = range_mapping.get(old._logical_comment_id, old._logical_comment_id)
        return self.document._wrap(node.part_name, copied)

    def insert_container(self, parent, node, reference):
        entries = self.contained_entries(node._element)
        if any(entry not in self.pending and (entry in self.threads.parents or self.threads.replies(entry))
               for entry in entries):
            raise NotImplementedError('Moving comment threads is not yet supported')
        if not self.safe_content(node._element):
            raise NotImplementedError('Moving complex annotation content requires resource translation')
        operations, additions, new_roles = [], {}, {}
        for entry in entries:
            if entry not in self.pending:
                continue
            role = self.pending[entry]
            name, tree, new_part = self._content_tree(role)
            if new_part and role not in new_roles:
                operations.extend(self._new_relationship(role, name, additions))
                additions[name] = tree
                new_roles[role] = name
            operations.append(('insert', tree.documentElement, entry, None))
        anchor = reference._element if reference is not None else None
        if _is(parent._element, 'body') and anchor is None:
            anchor = _find(parent._element, 'sectPr')
        operations.append(('insert', parent._element, node._element, anchor))
        self._transaction(operations, additions)
        self.parts.update(new_roles)
        for entry in entries:
            self.pending.pop(entry, None)
        self.changed = True
        return node

    def remove_container(self, node):
        if not self.safe_content(node._element):
            raise NotImplementedError('Deleting complex annotation content requires resource translation')
        parent = node.parent_node
        if parent is None:
            raise ValueError('Node has no parent')
        parent._validate_removal(node)
        entries = self.contained_entries(node._element)
        if any(entry in self.threads.parents or self.threads.replies(entry) for entry in entries):
            raise NotImplementedError('Removing comment threads is not yet supported')
        operations = [('remove', node._element)]
        for entry in entries:
            if entry.parentNode is not None:
                operations.append(('remove', entry))
        self._transaction(operations)
        for entry in entries:
            self.pending[entry] = entry.localName + 's'
        self.changed = True
        return node

    def parent(self, entry):
        if entry in self.threads.parents and entry not in self.anchors:
            return self.parent(self.threads.parents[entry])
        anchor = self.anchors.get(entry)
        if anchor is None or anchor.parentNode is None:
            return None
        return self.document._wrap(self.anchor_part(anchor), anchor.parentNode)

    def anchor_part(self, anchor):
        for name, tree in self.document._package._trees.items():
            if anchor.ownerDocument is tree:
                return name
        for name, tree in self.pending_trees.values():
            if anchor.ownerDocument is tree:
                return name
        raise ValueError('Annotation anchor has no document part')

    def _transaction(self, operations, new_trees=None):
        """Prepare all changed XML and bounds before modifying any held node."""
        from aspose.words_foss.utils.xml_helpers import serialize_xml
        from aspose.words_foss.dom.nodes import _copy_row_warm
        package = self.document._package
        new_trees = new_trees or {}
        trees = dict(package._trees)
        trees.update(new_trees)
        copies, mapping = {}, {}
        for name, tree in trees.items():
            copied = tree.cloneNode(deep=True)
            copies[name] = copied
            pairs = [(tree, copied)]
            while pairs:
                old, new = pairs.pop()
                mapping[old] = new
                pairs.extend(zip(old.childNodes, new.childNodes))
        changed = set(new_trees)
        owners = {tree: name for name, tree in trees.items()}
        for operation in operations:
            action, node, *args = operation
            changed.add(owners[node.ownerDocument])
            projected = mapping[node]
            if action == 'remove':
                projected.parentNode.removeChild(projected)
            elif action == 'insert':
                child, anchor = args
                copied_child = mapping.get(child)
                if copied_child is None:
                    copied_child = projected.ownerDocument.importNode(child, deep=True)
                    _copy_row_warm(child, copied_child)
                    pairs = [(child, copied_child)]
                    while pairs:
                        old, new = pairs.pop()
                        mapping[old] = new
                        pairs.extend(zip(old.childNodes, new.childNodes))
                projected.insertBefore(copied_child, mapping.get(anchor) if anchor is not None else None)
            elif action == 'attribute':
                key, value = args
                _set_word_attribute(projected, key, value)
            else:
                raise ValueError('Unknown annotation operation')
        payloads = {name: serialize_xml(copies[name]) for name in changed}
        package.validate_parts(payloads)
        # Every pointer, new node and serialized payload was prepared above.
        package._commit_parts({name: payloads[name] for name in new_trees})
        package._trees.update(new_trees)
        for operation in operations:
            action, node, *args = operation
            if action == 'remove':
                node.parentNode.removeChild(node)
            elif action == 'insert':
                child, anchor = args
                node.insertBefore(child, anchor)
            else:
                key, value = args
                _set_word_attribute(node, key, value)
        package._dirty.update(changed)

    def owns_pending(self, element):
        return element in self.pending

    def _content_tree(self, role):
        from xml.dom.minidom import Document
        if role in self.parts:
            name = self.parts[role]
            return name, self.document._package.tree(name), False
        if role not in self.pending_trees:
            name = 'word/' + role + '.xml'
            index = 0
            while name in self.document.part_names:
                index += 1
                name = f'word/{role}-{index}.xml'
            tree = Document()
            root = tree.createElementNS(W, 'w:' + role)
            root.setAttribute('xmlns:w', W)
            tree.appendChild(root)
            self.pending_trees[role] = name, tree
        name, tree = self.pending_trees[role]
        return name, tree, True

    def create(self, role, *, author='', initial='', date_time=None):
        if role == 'comments':
            from aspose.words_foss.dom.comment_metadata import normalize_date, normalize_text
            author, initial = normalize_text(author), normalize_text(initial)
            date_time = normalize_date(date_time, constructor=True)
        name, tree, _ = self._content_tree(role)
        node = _new(tree.documentElement, _ROLES[role])
        if role == 'comments':
            _set_word_attribute(node, 'author', author)
            _set_word_attribute(node, 'initials', initial)
            from aspose.words_foss.dom.comment_metadata import set_date
            set_date(node, date_time)
            self.comment_ids[node] = self._next_comment_id
            self._next_comment_id += 1
        self.pending[node] = role
        return self.document._wrap(name, node)

    def _new_relationship(self, role, name, additions):
        """Prepare typed relation/type additions without replacing any existing tree."""
        from aspose.words_foss.docx_writer.constants import CT_URI, PKG_RELS_URI
        package = self.document._package
        operations = []
        rels_name = 'word/_rels/document.xml.rels'
        if rels_name in package.part_names:
            rels = package.tree(rels_name).documentElement
        else:
            if rels_name not in additions:
                from xml.dom.minidom import Document
                tree = Document()
                root = tree.createElementNS(PKG_RELS_URI, 'Relationships')
                root.setAttribute('xmlns', PKG_RELS_URI)
                tree.appendChild(root)
                additions[rels_name] = tree
            rels = additions[rels_name].documentElement
        used = {item.getAttribute('Id') for item in _elements(rels)}
        identifier = 'annotation-' + role
        index = 0
        while identifier in used:
            index += 1
            identifier = f'annotation-{role}-{index}'
        link = rels.ownerDocument.createElementNS(PKG_RELS_URI, (rels.prefix + ':' if rels.prefix else '') + 'Relationship')
        link.setAttribute('Id', identifier)
        link.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + role)
        link.setAttribute('Target', name.removeprefix('word/'))
        operations.append(('insert', rels, link, None))
        types = package.tree('[Content_Types].xml').documentElement
        override = types.ownerDocument.createElementNS(CT_URI, (types.prefix + ':' if types.prefix else '') + 'Override')
        override.setAttribute('PartName', '/' + name)
        override.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + role + '+xml')
        operations.append(('insert', types, override, None))
        return operations

    def insert(self, parent, node, reference=None):
        from aspose.words_foss.dom.nodes import Paragraph, _copy_row_warm
        if not isinstance(parent, Paragraph):
            raise ValueError('An annotation must be inserted into a paragraph')
        if parent.owner_document is not self.document or node.owner_document is not self.document:
            raise ValueError('Cross-document insertion requires import_node')
        parent._editable()
        if reference is not None and reference.parent_node is not parent:
            raise ValueError('Reference must be a direct child')
        role = self.pending.get(node._element)
        if role is None:
            role = {'footnote': 'footnotes', 'endnote': 'endnotes', 'comment': 'comments'}[node._element.localName]
        entry = node._element
        ancestor = parent
        while ancestor is not None:
            if ancestor._element is entry:
                raise ValueError('Cannot insert a node into itself or its descendant')
            ancestor = ancestor.parent_node
        name, tree, new_part = self._content_tree(role)
        root = tree.documentElement
        if not all(_safe_structure(child) for child in _elements(entry)):
            raise NotImplementedError('Annotation content requires resource/range translation')
        existing = self.anchors.get(entry)
        if existing is not None and (entry in self.threads.parents or self.threads.replies(entry)):
            raise NotImplementedError('Moving comment threads is not yet supported')
        anchor = self.anchors.get(reference._element) if reference is not None else None
        if anchor is None and reference is not None:
            anchor = reference._element
        operations = []
        additions = {name: tree} if new_part else {}
        if new_part:
            operations.extend(self._new_relationship(role, name, additions))
        if existing is None:
            # The storage identifier is unique; a cloned Comment may retain a duplicate logical ID.
            identifiers = [annotation_id(item.getAttributeNS(W, 'id')) for item in _elements(root)
                           if item.hasAttributeNS(W, 'id')]
            identifiers.extend(int(key[1]) for key in self.entries if key[0] == role)
            identifier = str(max(identifiers, default=1 if role != 'comments' else -1) + 1)
            template = self.templates.get(entry)
            if template is None:
                run = _new(parent._element, 'r')
                marker = _new(run, {'footnotes': 'footnoteReference', 'endnotes': 'endnoteReference', 'comments': 'commentReference'}[role])
                run.appendChild(marker)
            else:
                from aspose.words_foss.dom.nodes import _bind_namespace_context
                run = parent._element.ownerDocument.importNode(template, deep=True)
                _bind_namespace_context(template, run, parent._element)
                marker = next(child for child in _elements(run) if child.localName in _REFERENCES)
            _set_word_attribute(marker, 'id', identifier)
            operations.append(('insert', root, entry, None))
            operations.append(('attribute', entry, 'id', identifier))
        else:
            run = existing
            if existing.ownerDocument is not parent._element.ownerDocument:
                from aspose.words_foss.dom.nodes import _bind_namespace_context
                if not _safe_structure(existing, properties=True):
                    raise NotImplementedError('Moving annotation anchors requires resource translation')
                run = parent._element.ownerDocument.importNode(existing, deep=True)
                _bind_namespace_context(existing, run, parent._element)
                operations.append(('remove', existing))
        operations.append(('insert', parent._element, run, anchor))
        self._transaction(operations, additions)
        self.parts[role] = name
        self.entries[role, str(annotation_id(entry.getAttributeNS(W, 'id')))] = entry
        if existing is not None and existing is not run:
            self.anchor_entries.pop(existing)
        self.anchors[entry] = run
        self.anchor_entries[run] = entry
        self.pending.pop(entry, None)
        self.changed = True
        return node.canonical_node

    def _detached_state(self, node, anchor):
        from aspose.words_foss.dom.nodes import _bind_namespace_context
        template = anchor.cloneNode(deep=True)
        _bind_namespace_context(anchor, template)
        return template, node.reference_mark if node._element.localName != 'comment' else None

    def remove(self, node):
        entry = node._element
        if entry in self.threads.parents or self.threads.replies(entry):
            raise NotImplementedError('Removing comment threads is not yet supported')
        anchor = self.anchors.get(entry)
        if anchor is None:
            raise ValueError('Node has no parent')
        if not all(_safe_structure(child) for child in _elements(entry)):
            raise NotImplementedError('Annotation content requires resource/range translation')
        role = {'footnote': 'footnotes', 'endnote': 'endnotes', 'comment': 'comments'}[entry.localName]
        template, mark = self._detached_state(node, anchor)
        operations = [('remove', anchor)]
        if entry.parentNode is not None:
            operations.append(('remove', entry))
        self._transaction(operations)
        self.templates[entry] = template
        if role != 'comments':
            entry._reference_mark = mark
        self.anchors.pop(entry)
        self.anchor_entries.pop(anchor)
        self.entries.pop((role, str(annotation_id(entry.getAttributeNS(W, 'id')))), None)
        self.pending[entry] = role
        self.threads.parents.pop(entry, None)
        self.changed = True
        return node

    def anchor_template(self, entry):
        anchor = self.anchors.get(entry, self.templates.get(entry))
        if anchor is not None and not _safe_structure(anchor, properties=True):
            raise NotImplementedError('Annotation anchor requires resource translation')
        return anchor

    def clone(self, node, deep):
        from aspose.words_foss.dom.nodes import _copy_row_warm, _bind_namespace_context
        if not isinstance(deep, bool):
            raise TypeError('deep must be a boolean')
        entry = node._element
        if entry in self.threads.parents or self.threads.replies(entry):
            raise NotImplementedError('Cloning comment threads is not yet supported')
        template = self.anchor_template(entry)
        if not all(_safe_structure(child) for child in _elements(entry)):
            raise NotImplementedError('Annotation content requires resource/range translation')
        copied = entry.cloneNode(deep=deep)
        _bind_namespace_context(entry, copied)
        _copy_row_warm(entry, copied)
        role = {'footnote': 'footnotes', 'endnote': 'endnotes', 'comment': 'comments'}[entry.localName]
        self.pending[copied] = role
        if template is not None:
            copied_template = template.cloneNode(deep=True)
            _bind_namespace_context(template, copied_template)
            self.templates[copied] = copied_template
        if role != 'comments':
            copied._reference_mark = node.reference_mark
        if role == 'comments':
            self.comment_ids[copied] = node.id
            if entry in self.threads.done:
                self.threads.done[copied] = self.threads.done[entry]
        return self.document._wrap(node.part_name, copied)

    def prepare_projection(self):
        """Pure serialization plans; commit public IDs/minimum content after successful IO."""
        from aspose.words_foss.utils.xml_helpers import serialize_xml
        from aspose.words_foss.dom.comment_metadata import project_date
        package = self.document._package
        if not self.changed and not self.threads.done:
            name = self.parts.get('comments')
            if name is None:
                return {}, lambda: None
            tree = package.tree(name)
            projected = tree.cloneNode(deep=True)
            changed = False
            for source, target in zip(_elements(tree.documentElement), _elements(projected.documentElement)):
                if _is(source, 'comment') and source.getAttributeNS(W, 'type') in {'', 'normal'}:
                    changed = project_date(source, target) or changed
            if not changed and not self.utc.parts and not self.utc.links:
                return {}, lambda: None
            from aspose.words_foss.dom.row_format import projected_row_tree
            trees = {name: projected}
            mapping = dict(zip(_elements(tree.documentElement), _elements(projected.documentElement)))
            commit_utc = self.utc.project(trees, mapping)
            payloads = {n: serialize_xml(projected_row_tree(t)) if t is not None else None for n, t in trees.items()}
            package.validate_parts(payloads, _allow_removed=True)
            return payloads, commit_utc
        names = {'word/document.xml', *self.parts.values(),
                 *(self.anchor_part(anchor) for anchor in self.anchors.values())}
        trees, mapping = {}, {}
        for name in names:
            tree = package.tree(name)
            cloned = tree.cloneNode(deep=True)
            trees[name] = cloned
            pairs = [(tree, cloned)]
            while pairs:
                old, new = pairs.pop()
                mapping[old] = new
                pairs.extend(zip(old.childNodes, new.childNodes))
        main = package.tree('word/document.xml')
        commits = []
        # Keep source storage intact until successful IO; the logical graph uses the first ID.
        for entry in self.duplicate_entries:
            projected = mapping[entry]
            projected.parentNode.removeChild(projected)
            commits.append(('duplicate', entry, None))
        # Native removes note/comment anchors outside the main story after successful IO.
        for entry, anchor in self.anchors.items():
            if (self.anchor_part(anchor) == 'word/document.xml'
                    and self._main_paragraph(anchor.parentNode)):
                continue
            ancestor = anchor
            while ancestor.parentNode is not None:
                ancestor = ancestor.parentNode
            if ancestor.nodeType != ancestor.DOCUMENT_NODE:
                continue
            projected_anchor = mapping[anchor]
            projected_anchor.parentNode.removeChild(projected_anchor)
            projected_entry = mapping.get(entry)
            if projected_entry is not None and projected_entry.parentNode is not None:
                projected_entry.parentNode.removeChild(projected_entry)
            node = self.document._wrap(self.parts[entry.localName + 's'], entry)
            template, mark = self._detached_state(node, anchor)
            commits.append(('outside', entry, (anchor, template, mark)))
        active = []
        for paragraph in main.getElementsByTagNameNS(W, 'p'):
            if self._main_paragraph(paragraph):
                for run in _elements(paragraph):
                    node = self.node_for_run(run) if _is(run, 'r') else None
                    if node is not None:
                        active.append(node._element)
        next_ids = {'footnote': 2, 'endnote': 2}
        comment_mapping, assigned, seen_comments = {}, {}, set()
        next_comment = 0
        # Native allocation follows inline order, including later-discarded range IDs.
        comment_ids = {}
        for paragraph in main.getElementsByTagNameNS(W, 'p'):
            if not self._main_paragraph(paragraph):
                continue
            for item in _elements(paragraph):
                if item.localName in {'commentRangeStart', 'commentRangeEnd'} and item.namespaceURI == W:
                    identifier = self.document._wrap('word/document.xml', item).id
                    if identifier not in assigned:
                        assigned[identifier] = next_comment
                        next_comment += 1
                elif item in self.anchor_entries:
                    entry = self.anchor_entries[item]
                    if entry.localName != 'comment':
                        continue
                    original = self.comment_ids[entry]
                    if original in assigned and original not in seen_comments:
                        identifier = assigned[original]
                    else:
                        identifier = next_comment
                        next_comment += 1
                        assigned.setdefault(original, identifier)
                    seen_comments.add(original)
                    comment_ids[entry] = identifier
                    comment_mapping.setdefault(original, identifier)
        for entry in active:
            if entry.localName == 'comment':
                identifier = comment_ids[entry]
                commits.append(('id', entry, identifier))
            else:
                identifier = next_ids[entry.localName]
                next_ids[entry.localName] += 1
            projected = mapping[entry]
            _set_word_attribute(projected, 'id', str(identifier))
            run = mapping[self.anchors[entry]]
            reference = next(child for child in _elements(run) if child.localName in _REFERENCES)
            _set_word_attribute(reference, 'id', str(identifier))
            if entry.localName != 'comment' and hasattr(entry, '_reference_mark'):
                mark = entry._reference_mark
                if mark:
                    _set_word_attribute(reference, 'customMarkFollows', '1')
                elif reference.hasAttributeNS(W, 'customMarkFollows'):
                    reference.removeAttributeNS(W, 'customMarkFollows')
                for child in list(_elements(run)):
                    if _is(child, 't'):
                        run.removeChild(child)
                if mark:
                    text = _new(run, 't')
                    text.appendChild(run.ownerDocument.createTextNode(mark))
                    run.appendChild(text)
            children = _elements(entry)
            if not children or not _is(children[-1], 'p'):
                paragraph = _new(entry, 'p')
                projected.appendChild(projected.ownerDocument.importNode(paragraph, deep=True))
                commits.append(('minimum', entry, paragraph))
        # Native discards orphan range nodes only after a successful save.
        for name in ['commentRangeStart', 'commentRangeEnd']:
            for marker in list(main.getElementsByTagNameNS(W, name)):
                if not self._main_paragraph(marker.parentNode):
                    continue
                value = self.document._wrap('word/document.xml', marker).id
                projected = mapping[marker]
                if value not in comment_mapping:
                    projected.parentNode.removeChild(projected)
                    commits.append(('remove-range', marker, None))
                else:
                    identifier = comment_mapping[value]
                    _set_word_attribute(projected, 'id', str(identifier))
                    commits.append(('range-id', marker, identifier))
        for (role, _), entry in self.entries.items():
            if role == 'comments' and entry in mapping:
                project_date(entry, mapping[entry])
        from aspose.words_foss.dom.row_format import projected_row_tree, validate_row_properties
        commit_utc = self.utc.project(trees, mapping)
        commit_threads = self.threads.project(trees, mapping)
        for tree in trees.values():
            if tree is not None:
                validate_row_properties(tree)
        payloads = {name: serialize_xml(projected_row_tree(tree)) if tree is not None else None for name, tree in trees.items()}
        package.validate_parts(payloads, _allow_removed=True)
        def commit():
            self._next_comment_id = next_comment
            for action, entry, value in commits:
                if action == 'duplicate':
                    if entry.parentNode is not None:
                        entry.parentNode.removeChild(entry)
                    self.duplicate_entries.remove(entry)
                    package._dirty.add(self.parts[entry.localName + 's'])
                elif action == 'outside':
                    value, template, mark = value
                    value.parentNode.removeChild(value)
                    if entry.parentNode is not None:
                        entry.parentNode.removeChild(entry)
                    self.templates[entry] = template
                    if entry.localName != 'comment':
                        entry._reference_mark = mark
                    self.anchors.pop(entry)
                    self.anchor_entries.pop(value)
                    self.pending[entry] = entry.localName + 's'
                    package._dirty.update({self.anchor_part(value), self.parts[entry.localName + 's']})
                elif action == 'id':
                    self.comment_ids[entry] = value
                elif action == 'range-id':
                    entry._logical_comment_id = value
                elif action == 'remove-range':
                    entry.parentNode.removeChild(entry)
                    package._dirty.add('word/document.xml')
                else:
                    entry.appendChild(value)
                    package._dirty.add(self.parts['comments' if entry.localName == 'comment' else entry.localName + 's'])
            commit_utc()
            commit_threads()
        return payloads, commit

    def _import_mode(self, node, mode):
        from aspose.words_foss.dom.importing import ImportFormatMode, _theme_payload
        from aspose.words_foss.dom.styles import StyleResolver
        if mode != ImportFormatMode.USE_DESTINATION_STYLES:
            source_styles = StyleResolver(node.owner_document)
            target_styles = StyleResolver(self.document)
            old = _find(source_styles.root, 'docDefaults') if source_styles.root is not None else None
            new = _find(target_styles.root, 'docDefaults') if target_styles.root is not None else None
            if ((old.toxml() if old is not None else None) != (new.toxml() if new is not None else None)
                    or _theme_payload(node.owner_document) != _theme_payload(self.document)
                    or (source_styles.root.toxml() if source_styles.root is not None else None)
                    != (target_styles.root.toxml() if target_styles.root is not None else None)):
                raise NotImplementedError('Source annotation formatting requires theme/default translation')

    def import_container(self, node, deep, mode):
        from aspose.words_foss.dom.importing import _prepare_import, ImportFormatMode
        from aspose.words_foss.dom.nodes import _bind_namespace_context, _copy_row_warm
        source_graph = node.owner_document._annotations()
        self._import_mode(node, mode)
        entries = source_graph.contained_entries(node._element) if deep else []
        if not source_graph.safe_content(node._element) or any(
                not all(_safe_structure(child) for child in _elements(entry)) for entry in entries):
            raise NotImplementedError('Annotation bundle requires resource/range translation')
        # One style plan covers both the container and all linked content parts.
        bundle = _new(node._element, 'tc')
        copied = node._element.cloneNode(deep=True)
        _bind_namespace_context(node._element, copied)
        _copy_row_warm(node._element, copied)
        if not deep:
            from aspose.words_foss.dom.nodes import _PROPERTY_ROOTS
            for child in list(copied.childNodes):
                if child.nodeType != child.ELEMENT_NODE or child.localName not in _PROPERTY_ROOTS | {'tblGrid'}:
                    copied.removeChild(child)
        bundle.appendChild(copied)
        for entry in entries:
            copied_entry = entry.cloneNode(deep=True)
            _bind_namespace_context(entry, copied_entry)
            _copy_row_warm(entry, copied_entry)
            bundle.appendChild(copied_entry)
        prepared, commit_styles = _prepare_import(self.document, node, True,
            ImportFormatMode.USE_DESTINATION_STYLES,
            copied=node.owner_document._wrap(node.part_name, bundle))
        imported = _elements(prepared._element)[0]
        imported_entries = _elements(prepared._element)[1:]
        runs = [run for run in imported.getElementsByTagNameNS(W, 'r')
                if any(child.localName in _REFERENCES and child.namespaceURI == W for child in _elements(run))]
        planned = []
        identifiers = {}
        for old, entry, run in zip(entries, imported_entries, runs):
            role = old.localName + 's'
            name, tree, _ = self._content_tree(role)
            value = identifiers.get(role, max((int(key[1]) for key in self.entries if key[0] == role),
                        default=1 if role != 'comments' else -1)) + 1
            identifiers[role] = value
            target_entry = tree.importNode(entry, deep=True)
            _bind_namespace_context(entry, target_entry)
            _copy_row_warm(entry, target_entry)
            if role != 'comments':
                target_entry._reference_mark = source_graph.node_for_run(source_graph.anchors[old]).reference_mark
            _set_word_attribute(target_entry, 'id', str(value))
            marker = next(child for child in _elements(run) if child.localName in _REFERENCES)
            _set_word_attribute(marker, 'id', str(value))
            planned.append((old, target_entry, run, role, str(value)))
        # Range IDs are logical IDs; original physical source IDs must not leak into the target.
        next_id = self._next_comment_id
        range_mapping = {}
        for old, entry, run, role, identifier in planned:
            if role == 'comments':
                range_mapping.setdefault(source_graph.comment_ids[old], next_id)
                next_id += 1
        for old, new in zip(node._element.getElementsByTagNameNS(W, '*'), imported.getElementsByTagNameNS(W, '*')):
            if hasattr(old, '_logical_comment_id'):
                new._logical_comment_id = range_mapping.get(old._logical_comment_id, old._logical_comment_id)
        prepared._element.removeChild(imported)
        commit_styles()
        for old, entry, run, role, identifier in planned:
            self.pending[entry] = role
            self.entries[role, identifier] = entry
            self.anchors[entry] = run
            self.anchor_entries[run] = entry
            if role == 'comments':
                self.comment_ids[entry] = self._next_comment_id
                self._next_comment_id += 1
        translated = {old: entry for old, entry, _, role, _ in planned if role == 'comments'}
        for old, entry in translated.items():
            if old in source_graph.threads.done:
                self.threads.done[entry] = source_graph.threads.done[old]
            parent = source_graph.threads.parents.get(old)
            if parent in translated:
                self.threads.parents[entry] = translated[parent]
        return self.document._wrap('word/document.xml', imported)

    def import_entry(self, node, deep, mode):
        from aspose.words_foss.dom.importing import _prepare_import, ImportFormatMode
        from aspose.words_foss.dom.nodes import _bind_namespace_context, _copy_row_warm
        self._import_mode(node, mode)
        source_graph = node.owner_document._annotations()
        template = source_graph.anchor_template(node._element)
        if not all(_safe_structure(child) for child in _elements(node._element)):
            raise NotImplementedError('Annotation content requires resource translation')
        bundle = _new(node._element, 'tc')
        copied = node._element.cloneNode(deep=deep)
        _bind_namespace_context(node._element, copied)
        _copy_row_warm(node._element, copied)
        bundle.appendChild(copied)
        if template is not None:
            copied_template = template.cloneNode(deep=True)
            _bind_namespace_context(template, copied_template)
            bundle.appendChild(copied_template)
        prepared, commit_styles = _prepare_import(self.document, node, True,
            ImportFormatMode.USE_DESTINATION_STYLES,
            copied=node.owner_document._wrap(node.part_name, bundle))
        source_entry = _elements(prepared._element)[0]
        role = node._element.localName + 's'
        name, tree, _ = self._content_tree(role)
        imported = tree.importNode(source_entry, deep=True)
        _bind_namespace_context(source_entry, imported)
        _copy_row_warm(source_entry, imported)
        planned_template = _elements(prepared._element)[1] if template is not None else None
        commit_styles()
        self.pending[imported] = role
        if planned_template is not None:
            self.templates[imported] = planned_template
        if role != 'comments':
            imported._reference_mark = node.reference_mark
        if role == 'comments':
            self.comment_ids[imported] = self._next_comment_id
            self._next_comment_id += 1
            if node._element in source_graph.threads.done:
                self.threads.done[imported] = source_graph.threads.done[node._element]
        return self.document._wrap(name, imported)
