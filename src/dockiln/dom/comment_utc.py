"""Typed modern comment UTC metadata and source-independent node values."""
import re
from datetime import datetime, timedelta, timezone
from xml.dom.minidom import Document

from dockiln._opc import resolve_target
from dockiln.dom.comment_metadata import W, _DATE, _parse_date, _date_offset, _local_wall, get_date

W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid'
CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
ROLES = {'commentsIds': (CID, 'commentId', 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds'),
         'commentsExtensible': (CEX, 'commentExtensible', 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible')}


def elements(node, namespace, local):
    return [c for c in node.childNodes if c.nodeType == c.ELEMENT_NODE
            and c.namespaceURI == namespace and c.localName == local]


def hex_id(raw):
    # Native trims ASCII whitespace and accepts leading zeroes beyond eight digits.
    lexical = raw.strip(' \t\r\n')
    if not lexical:
        return None
    lexical = lexical.lstrip('0') or '0'
    return int(lexical, 16) if re.fullmatch('[0-9a-fA-F]{1,8}', lexical) else None


def durable_id(node, namespace):
    if not node.hasAttributeNS(namespace, 'durableId'):
        return 0
    value = hex_id(node.getAttributeNS(namespace, 'durableId'))
    return 0x80000000 if value is None else value


def para_id(raw):
    value = hex_id(raw)
    return value if value is not None and 0 < value < 0x80000000 else None


def utc_date(raw):
    value, residual = _parse_date(raw)
    if value == datetime.min and not residual:
        lexical = raw.strip()
        if not (_DATE.fullmatch(lexical) and lexical.startswith('0001-01-01')
                and (len(lexical) == 10 or lexical[11:19] == '00:00:00')):
            return None
    value = value.replace(tzinfo=None)
    try:
        zone = _date_offset(raw.strip())
    except ValueError:
        return None
    if zone is None:
        return None if value == datetime.min and not residual else value
    offset = zone[1]
    try:
        result = (value - offset).replace(tzinfo=timezone.utc)
        return None if result == datetime.min.replace(tzinfo=timezone.utc) and not residual else result
    except OverflowError:
        if value.year == 1:
            # DateTime's source parser wraps a negative UTC tick count by one day.
            return (value + timedelta(days=1) - offset).replace(tzinfo=timezone.utc)
        return None


def get_utc(element):
    state = getattr(element, '_comment_utc_source', None)
    if state is not None and state[0] is element and state[1] == element.getAttributeNS(W, 'date'):
        return state[2]
    value = get_date(element)
    residual = getattr(element, '_comment_date_warm', (None, None, None, 0))[3]
    if value.replace(tzinfo=None) == datetime.min and not residual:
        return datetime.min
    if value.tzinfo is None:
        value = _local_wall(value)
    try:
        return value.astimezone(timezone.utc)
    except OverflowError:
        return (datetime.min if value.year == 1 else datetime.max).replace(tzinfo=timezone.utc)


def copy_utc(source, target):
    roles = getattr(source, '_comment_metadata_roles', None)
    if roles is not None and roles[0] is source:
        target._comment_metadata_roles = (target, roles[1])
    state = getattr(source, '_comment_utc_source', None)
    if state is not None and state[0] is source and state[1] == source.getAttributeNS(W, 'date'):
        target._comment_utc_source = (target, state[1], state[2])


def paragraphs(entry):
    result = []
    for p in entry.getElementsByTagNameNS(W, 'p'):
        parent = p.parentNode
        while parent is not entry and parent is not None:
            if parent.namespaceURI != W or parent.localName not in {'tbl', 'tr', 'tc'}:
                break
            parent = parent.parentNode
        if parent is entry:result.append(p)
    return result


def source_link(entry, ids):
    identifiers = {para_id(p.getAttributeNS(W14, 'paraId')) for p in paragraphs(entry)} - {None}
    return next((node for identifier, node in reversed(ids.items()) if identifier in identifiers), None)


def metadata_links(package):
    name = 'word/_rels/document.xml.rels'
    if name not in package.part_names:
        return []
    types = {value[2] for value in ROLES.values()}
    return [node for node in elements(package.tree(name).documentElement, REL, 'Relationship')
            if node.getAttribute('Type') in types]


def selected_links(package):
    name = 'word/_rels/document.xml.rels'
    if name not in package.part_names:
        return {}
    # 26.9 chooses the first ordinal relationship ID, after duplicate-ID overwrite.
    links = {}
    for link in elements(package.tree(name).documentElement, REL, 'Relationship'):
        if not link.getAttribute('Target'):
            continue
        identifier = link.getAttribute('Id')
        if not identifier:
            raise ValueError('Expected a non-empty relationship identifier')
        links[identifier] = link
    selected = {}
    for identifier in sorted(links, key=lambda value: value.encode('utf-16-be')):
        link = links[identifier]
        role = next((r for r, (_, _, t) in ROLES.items() if link.getAttribute('Type') == t), None)
        if role is not None and role not in selected:
            selected[role] = link
    return selected


def package_target(link, owner):
    try:
        return resolve_target(owner, link.getAttribute('Target'))
    except ValueError:
        if link.getAttribute('TargetMode') == 'External':
            return None
        raise


def typed_parts(package):
    result = {}
    for role, link in selected_links(package).items():
        part = package_target(link, 'word/document.xml')
        if part not in package.part_names:
            continue
        root = package.tree(part).documentElement
        if root.namespaceURI == ROLES[role][0] and root.localName == role:
            result[role] = part
    return result


def set_attr(node, namespace, prefix, local, value):
    node.setAttribute('xmlns:' + prefix, namespace)
    node.setAttributeNS(namespace, prefix + ':' + local, value)


class CommentUtc:
    def __init__(self, package, entries):
        self.package = package
        self.entries = entries
        self.source_comments = {entry for (role, _), entry in entries.items() if role == 'comments'}
        self.parts = typed_parts(package)
        self.links = metadata_links(package)
        self.roles = {role for role, link in selected_links(package).items()
                      if package_target(link, 'word/document.xml') in package.part_names}
        self.ids, self.ext = self._maps()
        for (role, _), entry in entries.items():
            if role != 'comments':
                continue
            entry._comment_metadata_roles = (entry, frozenset(self.roles))
            link = source_link(entry, self.ids)
            extension = self.ext.get(durable_id(link, CID)) if link is not None else None
            value = utc_date(extension.getAttributeNS(CEX, 'dateUtc')) if extension is not None else None
            if value is not None:
                entry._comment_utc_source = (entry, entry.getAttributeNS(W, 'date'), value)

    def _maps(self):
        ids, ext = {}, {}
        if 'commentsIds' in self.parts:
            for node in elements(self.package.tree(self.parts['commentsIds']).documentElement, CID, 'commentId'):
                identifier = para_id(node.getAttributeNS(CID, 'paraId'))
                ids.pop(identifier, None)
                ids[identifier] = node
        if 'commentsExtensible' in self.parts:
            for node in elements(self.package.tree(self.parts['commentsExtensible']).documentElement, CEX, 'commentExtensible'):
                ext[durable_id(node, CEX)] = node
        return ids, ext

    def project(self, trees, mapping):
        known = self.source_comments | {entry for (role, _), entry in self.entries.items() if role == 'comments'}
        entries = [(source, target) for source, target in mapping.items()
                   if source.nodeType == source.ELEMENT_NODE and source.namespaceURI == W
                   and source.localName == 'comment' and source in known and target.parentNode is not None]
        if not self.links and not self.parts and not any(getattr(s, '_comment_utc_source', None) or getattr(s, '_comment_metadata_roles', (None, ()))[1] for s, _ in entries):
            return lambda: None
        active_roles = set(self.roles)
        for source, _ in entries:
            provenance = getattr(source, '_comment_metadata_roles', None)
            if provenance is not None and provenance[0] is source:
                active_roles.update(provenance[1])
            if getattr(source, '_comment_utc_source', None):
                active_roles.add('commentsExtensible')
        if active_roles:
            active_roles.add('commentsIds')
        ids, ext = self._maps()
        selected_ids = {source_link(source, ids) for source, _ in entries} - {None}
        selected_ext = {ext.get(durable_id(node, CID)) for node in selected_ids} - {None}
        discarded_records = 0
        reserved_para = set(ids)
        reserved_durable = set(ext)
        if 'commentsIds' in self.parts:
            reserved_durable |= {durable_id(node, CID) for node in
                elements(self.package.tree(self.parts['commentsIds']).documentElement, CID, 'commentId')}
        names = dict(self.parts)
        for role, (ns, _, rel_type) in ROLES.items():
            if role not in active_roles:
                continue
            name = names.get(role)
            if name is None:
                name = 'word/' + role + '.xml'
                number = 0
                while name in self.package.part_names or name in trees:
                    number += 1;name = 'word/' + role + '-' + str(number) + '.xml'
                names[role] = name
                tree = Document();root = tree.createElementNS(ns, 'm:' + role)
                root.setAttribute('xmlns:m', ns);tree.appendChild(root)
                trees[name] = tree
                if '[Content_Types].xml' not in trees:
                    trees['[Content_Types].xml'] = self.package.tree('[Content_Types].xml').cloneNode(True)
                types = trees['[Content_Types].xml'].documentElement
                for override in list(elements(types, CT, 'Override')):
                    if resolve_target('', override.getAttribute('PartName')) == name:
                        types.removeChild(override)
                link = types.ownerDocument.createElementNS(CT, 'Override');link.setAttribute('xmlns', CT)
                link.setAttribute('PartName', '/' + name);link.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + role + '+xml');types.appendChild(link)
            elif name not in trees:
                trees[name] = self.package.tree(name).cloneNode(True)
            root = trees[name].documentElement
            if role in self.parts:
                originals = elements(self.package.tree(name).documentElement, ns, ROLES[role][1])
                copies = elements(root, ns, ROLES[role][1])
                selected = selected_ids if role == 'commentsIds' else selected_ext
                discarded_records += sum(original not in selected for original in originals)
                for copy in copies:root.removeChild(copy)
        rel_name = 'word/_rels/document.xml.rels'
        if rel_name not in trees:
            trees[rel_name] = self.package.tree(rel_name).cloneNode(True)
        rels = trees[rel_name].documentElement
        modern_types = {value[2] for value in ROLES.values()}
        for link in list(elements(rels, REL, 'Relationship')):
            if link.getAttribute('Type') in modern_types:
                rels.removeChild(link)
        relationship_ids = {link.getAttribute('Id') for link in elements(rels, REL, 'Relationship')}
        selected = selected_links(self.package)
        for role, name in names.items():
            original = selected.get(role)
            if original is not None and package_target(original, 'word/document.xml') != name:
                original = None
            rid = 'comment-metadata-' + role
            if original is not None and original.getAttribute('Id') not in relationship_ids:
                rid = original.getAttribute('Id')
            while rid in relationship_ids:
                rid += '-1'
            relationship_ids.add(rid)
            link = rels.ownerDocument.importNode(original, True) if original is not None else rels.ownerDocument.createElementNS(REL, 'Relationship')
            link.setAttribute('xmlns', REL)
            if link.hasAttribute('TargetMode'):
                link.removeAttribute('TargetMode')
            for key, value in [('Id', rid), ('Type', ROLES[role][2]), ('Target', name.removeprefix('word/'))]:
                link.setAttribute(key, value)
            rels.appendChild(link)
        # Only recognized losing metadata roots are eligible for package removal.
        candidates = set()
        for link in self.links:
            if not link.getAttribute('Target'):
                continue
            name = package_target(link, 'word/document.xml')
            if name in names.values() or name not in self.package.part_names:
                continue
            root = self.package.tree(name).documentElement
            if any(root.namespaceURI == ns and root.localName == role for role, (ns, _, _) in ROLES.items()):
                candidates.add(name)
        from dockiln._opc import relationships_path
        import posixpath
        retained = set()
        for name in self.package.part_names:
            if not name.endswith('.rels') or name in {relationships_path(n) for n in candidates}:
                continue
            directory, filename = posixpath.split(name)
            owner = posixpath.join(posixpath.dirname(directory), filename[:-5])
            tree = trees.get(name, self.package.tree(name))
            for link in elements(tree.documentElement, REL, 'Relationship'):
                if link.getAttribute('Target'):
                    retained.add(package_target(link, owner))
        removed = candidates - retained
        for name in removed:
            trees[name] = None
            related = relationships_path(name)
            if related in self.package.part_names:
                trees[related] = None
        if removed:
            if '[Content_Types].xml' not in trees:
                trees['[Content_Types].xml'] = self.package.tree('[Content_Types].xml').cloneNode(True)
            types = trees['[Content_Types].xml'].documentElement
            for override in list(elements(types, CT, 'Override')):
                if resolve_target('', override.getAttribute('PartName')) in {n for n, tree in trees.items() if tree is None}:
                    types.removeChild(override)
        discarded_records += max(0, len(self.links) - len(names)) + len(removed)
        reserved = reserved_para | {hex_id(p.getAttributeNS(W14, 'paraId')) for _, t in entries for p in paragraphs(t)}
        used_para = set()
        used_durable, edits = set(), []
        def allocate(value, used, forbidden=(), limit=0x100000000):
            if value is None or value == 0 or value in used or value >= limit:
                value = 1
                while value in used or value in forbidden:value += 1
            used.add(value);return value
        for source, target in entries:
            target_paragraphs = paragraphs(target)
            if not target_paragraphs:continue
            old_paragraphs = paragraphs(source)
            for index, paragraph in enumerate(target_paragraphs[:-1]):
                value = para_id(paragraph.getAttributeNS(W14, 'paraId'))
                if value is None:continue
                assigned = allocate(value, used_para, reserved, 0x80000000)
                if assigned != value:
                    text = f'{assigned:08X}'
                    set_attr(paragraph, W14, 'w14', 'paraId', text)
                    if index < len(old_paragraphs):edits.append((old_paragraphs[index], text))
            old_pid = hex_id(target_paragraphs[-1].getAttributeNS(W14, 'paraId'))
            pid = allocate(old_pid, used_para, reserved, 0x80000000)
            old_link = source_link(source, ids)
            did = allocate(durable_id(old_link, CID) if old_link is not None else None, used_durable, reserved_durable)
            set_attr(target_paragraphs[-1], W14, 'w14', 'paraId', f'{pid:08X}')
            if old_paragraphs:edits.append((old_paragraphs[-1], f'{pid:08X}'))
            for role, ns, local, old_node in [('commentsIds', CID, 'commentId', old_link),
                     ('commentsExtensible', CEX, 'commentExtensible', ext.get(durable_id(old_link, CID)) if old_link is not None else None)]:
                if role not in active_roles:
                    continue
                root = trees[names[role]].documentElement
                node = root.ownerDocument.importNode(old_node, True) if old_node is not None else root.ownerDocument.createElementNS(ns, 'm:' + local)
                set_attr(node, ns, 'm', 'durableId', f'{did:08X}')
                if role == 'commentsIds':set_attr(node, ns, 'm', 'paraId', f'{pid:08X}')
                else:
                    value = get_utc(source)
                    if value == datetime.min and value.tzinfo is None:
                        if old_node is not None:discarded_records += 1
                        continue
                    else:set_attr(node, ns, 'm', 'dateUtc', value.replace(tzinfo=None).isoformat(timespec='seconds') + 'Z')
                boundary = next(iter(elements(root, ns, 'extLst')), None)
                root.insertBefore(node, boundary)
        if discarded_records:
            from dockiln.diagnostics import ContentLossWarning, warn
            warn('Unassociated or non-winning modern comment metadata records are omitted',
                 ContentLossWarning, code='comment.utc_metadata_normalized')
        from dockiln.utils.xml_helpers import serialize_xml
        payloads = {name: serialize_xml(trees[name]) if trees[name] is not None else None for name in set(names.values()) | {n for n, tree in trees.items() if tree is None} | (set(trees) & {'word/_rels/document.xml.rels', '[Content_Types].xml'})}
        prepared_links = [link for link in elements(rels, REL, 'Relationship') if link.getAttribute('Type') in modern_types]
        def commit():
            self.package._commit_parts(payloads)
            self.parts = names
            self.roles = set(names)
            self.links = prepared_links
            self.source_comments |= known
            for node, value in edits:set_attr(node, W14, 'w14', 'paraId', value)
        return commit
