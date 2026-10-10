"""Original OPC payloads with lazy, namespace-preserving XML trees."""

from hashlib import sha256
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from defusedxml.minidom import parseString

from aspose.words_foss import _io
from aspose.words_foss._opc import resolve_target
from aspose.words_foss._flat_opc import decode, is_flat_opc
from aspose.words_foss._io import (
    atomic_output,
    check_input_size,
    read_bounded,
    validate_docx_archive,
)
from aspose.words_foss.dom.nodes import W, _elements, _is, _safe_structure
from aspose.words_foss.utils.xml_helpers import serialize_xml


def _empty_plain_table(element):
    formatting = {"tbl", "tblPr", "tblGrid", "gridCol", "tblStyle", "tblW", "tblLook",
                  "tblLayout", "jc", "tblInd", "tblCellSpacing", "tblCellMar", "tblBorders",
                  "top", "left", "bottom", "right", "start", "end", "insideH", "insideV",
                  "shd", "bidiVisual", "tblOverlap", "tblpPr", "tblStyleRowBandSize", "tblStyleColBandSize"}
    return (_is(element, "tbl") and _safe_structure(element) and
            all(_is(child, "tblPr") or _is(child, "tblGrid") for child in _elements(element)) and
            all(parent.namespaceURI == W and parent.localName in formatting and
                all(child.nodeType == child.ELEMENT_NODE or
                    child.nodeType == child.TEXT_NODE and not child.data.strip()
                    for child in parent.childNodes)
                for parent in [element, *element.getElementsByTagNameNS("*", "*")]))


STORY_ROOTS = {"header": "hdr", "footer": "ftr", "footnotes": "footnotes",
               "endnotes": "endnotes", "comments": "comments"}


class DocxPackage:
    def __init__(self, source):
        if hasattr(source, "read"):
            data = read_bounded(source)
        else:
            path = Path(source)
            check_input_size(path.stat().st_size)
            with path.open("rb") as stream:
                data = read_bounded(stream)
        self._trees = {}
        self._dirty = set()
        self._injected_default_theme = False
        self._style_projection_part = None
        self._style_property_overrides = set()
        if is_flat_opc(data):
            data = decode(data)
        with ZipFile(BytesIO(data)) as archive:
            validate_docx_archive(archive)
            self._entries = archive.infolist()
            self._payloads = {entry.filename: archive.read(entry) for entry in self._entries}
            self._comment = archive.comment
        self._original_hashes = {name: sha256(data).digest()
                                 for name, data in self._payloads.items()}
        if "word/document.xml" not in self._payloads:
            raise ValueError("Source is not a DOCX package")
        if any(name.lower().startswith("_xmlsignatures/") for name in self._payloads):
            raise ValueError("Signed DOCX packages cannot be edited")
        for name in self._payloads:
            if name.endswith(".rels"):
                tree = self.tree(name)
                if any(node.getAttribute("Type").startswith((
                    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/",
                    "http://purl.oclc.org/ooxml/package/relationships/digital-signature/",
                )) for node in tree.getElementsByTagNameNS("*", "Relationship")):
                    raise ValueError("Signed DOCX packages cannot be edited")

    @property
    def part_names(self):
        return tuple(self._payloads)

    def story_part_names(self):
        """Discover editable stories through internal relationships and part types."""
        names = {"word/document.xml"}
        rels_name = "word/_rels/document.xml.rels"
        if rels_name not in self.part_names:
            return names
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        types = self.tree("[Content_Types].xml")
        overrides = {resolve_target("", node.getAttribute("PartName")): node.getAttribute("ContentType")
                     for node in types.getElementsByTagNameNS(ct_ns, "Override")}
        defaults = {node.getAttribute("Extension").lower(): node.getAttribute("ContentType")
                    for node in types.getElementsByTagNameNS(ct_ns, "Default")}
        for link in self.tree(rels_name).getElementsByTagNameNS(rel_ns, "Relationship"):
            relation = link.getAttribute("Type")
            if relation not in {rel_type + kind for kind in STORY_ROOTS} or link.getAttribute("TargetMode") == "External":
                continue
            kind = relation.removeprefix(rel_type)
            if link.getAttribute("TargetMode") not in {"", "Internal"}:
                raise ValueError("Expected an internal story relationship")
            name = resolve_target("word/document.xml", link.getAttribute("Target"))
            if name not in self.part_names:
                raise ValueError("Story relationship points to a missing part")
            content_type = overrides.get(name, defaults.get(name.rsplit(".", 1)[-1].lower()))
            if content_type != f"application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml":
                raise ValueError("Expected a Word story content type")
            if not _is(self.tree(name).documentElement, STORY_ROOTS[kind]):
                raise ValueError("Expected a Word story part")
            names.add(name)
        return names

    def tree(self, name):
        if name not in self._trees:
            self._trees[name] = parseString(self._payloads[name], forbid_dtd=True)
        return self._trees[name]

    def payload(self, name):
        if name == self._style_projection_part:
            from aspose.words_foss.dom.styles import serialized_style_payload

            data = serialized_style_payload(self, name)
            if data is not None:
                return data
        tree = None
        if name in self.story_part_names():
            from aspose.words_foss.dom.row_format import projected_row_tree, validate_row_properties
            original = self.tree(name)
            validate_row_properties(original)
            projected = projected_row_tree(original)
            from aspose.words_foss.dom.comment_metadata import projected_comment_tree
            projected = projected_comment_tree(original, projected)
            if projected is not original:
                # Read-only logical annotation normalization must not rewrite source runs/content.
                if name not in self._dirty and name in getattr(self, '_annotation_normalized_parts', ()):
                    source = parseString(self._payloads[name], forbid_dtd=True)
                    projected = projected_comment_tree(source, projected_row_tree(source))
                tree = projected
        if name in self._dirty or tree is not None:
            if tree is None:
                tree = self.tree(name)
            # Native saving omits empty tables; retain the live detached/imported DOM state.
            if any(_empty_plain_table(node) for node in tree.getElementsByTagNameNS(W, "tbl")):
                tree = tree.cloneNode(deep=True)
                for node in list(tree.getElementsByTagNameNS(W, "tbl")):
                    if _empty_plain_table(node):
                        node.parentNode.removeChild(node)
            return serialize_xml(tree)
        return self._payloads[name]

    def preservation_report(self, overrides=None):
        """Compare uncompressed part bytes with the loaded package, including resources."""
        overrides = overrides or {}
        report = {"added": [], "modified": [], "removed": [], "unchanged": []}
        for name in sorted(set(self._original_hashes) | set(self._payloads) | set(overrides)):
            if name not in self._original_hashes and name in overrides:
                state = "added"
            elif name not in self._payloads:
                state = "removed"
            elif name not in self._original_hashes:
                state = "added"
            elif sha256(overrides[name] if name in overrides else self.payload(name)).digest() != self._original_hashes[name]:
                state = "modified"
            else:
                state = "unchanged"
            report[state].append(name)
        return report

    def validate_parts(self, parts):
        """Validate a prospective complete part update without changing live XML."""
        for name, data in parts.items():
            if (not name or name.startswith("/") or ".." in name.split("/") or
                    "\\" in name or not isinstance(data, bytes)):
                raise ValueError("Unsafe DOCX part")
            if len(data) > _io.MAX_PART_BYTES:
                raise ValueError("DOCX part exceeds the safety limit")
        names = set(self._payloads) | set(parts)
        if len(names) > _io.MAX_ZIP_ENTRIES:
            raise ValueError("DOCX has too many ZIP entries")
        size = sum(len(parts[name] if name in parts else self.payload(name)) for name in names)
        if size > _io.MAX_EXPANDED_BYTES:
            raise ValueError("DOCX expanded size exceeds the safety limit")

    def set_parts(self, parts):
        """Commit prebuilt resource parts together, after checking output size limits."""
        self.validate_parts(parts)
        self._commit_parts(parts)

    def _commit_parts(self, parts):
        for name, data in parts.items():
            if name not in self._payloads:
                entry = ZipInfo(name)
                entry.compress_type = ZIP_DEFLATED
                self._entries.append(entry)
            self._payloads[name] = data
            self._trees.pop(name, None)
            self._dirty.discard(name)

    def _default_theme_parts(self, overrides=None):
        from aspose.words_foss._theme import missing_theme_parts
        from defusedxml.ElementTree import fromstring

        overrides = overrides or {}
        payloads = {name: overrides[name] if name in overrides else self.payload(name)
                    for name in set(self._payloads) | set(overrides)}
        ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
        types = fromstring(payloads['[Content_Types].xml'], forbid_dtd=True)
        if types.tag != ct + 'Types':
            raise ValueError('Expected an OPC content types part')
        declared = {resolve_target('', item.get('PartName', '')): item.get('ContentType', '')
                    for item in types if item.tag == ct + 'Override'}
        defaults = {item.get('Extension', '').lower(): item.get('ContentType', '')
                    for item in types if item.tag == ct + 'Default'}
        xml_parts = {name for name in payloads
                     if (kind := declared.get(name, defaults.get(name.rsplit('.', 1)[-1].lower(), '')))
                     in {'application/xml', 'text/xml'} or kind.endswith('+xml')}
        # Content types, rather than filename suffixes, distinguish XML from opaque bytes.
        # The package-preserving API adds defaults when an actual theme reference needs them.
        theme_attributes = {f'{{{W}}}{name}' for name in ('asciiTheme', 'hAnsiTheme', 'cstheme', 'eastAsiaTheme')}
        drawing = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
        referenced = any(theme_attributes.intersection(node.attrib) or
                         (node.tag == f'{{{W}}}color' and f'{{{W}}}themeColor' in node.attrib) or
                         (node.tag == drawing + 'fillRef' and node.get('idx', '0') not in ('0', '1000'))
                         for name, data in payloads.items()
                         if name.startswith('word/') and name in xml_parts
                         for node in fromstring(data).iter()
                         if node.tag in {f'{{{W}}}rFonts', f'{{{W}}}color', drawing + 'fillRef'})
        return missing_theme_parts(payloads) if referenced else {}

    def to_bytes(self, overrides=None):
        overrides = overrides or {}
        theme_parts = self._default_theme_parts(overrides)
        entries = list(self._entries)
        for name in theme_parts:
            if name not in self._payloads:
                entry = ZipInfo(name)
                entry.compress_type = ZIP_DEFLATED
                entries.append(entry)
        if set(overrides) - set(self.part_names):
            self.validate_parts(overrides)
        if len(set(self.part_names) | set(theme_parts) | set(overrides)) > _io.MAX_ZIP_ENTRIES:
            raise ValueError("DOCX has too many ZIP entries")
        stream = BytesIO()
        expanded = 0
        with ZipFile(stream, "w") as output:
            output.comment = self._comment
            for entry in entries:
                data = (theme_parts[entry.filename] if entry.filename in theme_parts else
                        overrides[entry.filename] if entry.filename in overrides
                        else self.payload(entry.filename))
                if len(data) > _io.MAX_PART_BYTES:
                    raise ValueError("DOCX part exceeds the safety limit")
                expanded += len(data)
                if expanded > _io.MAX_EXPANDED_BYTES:
                    raise ValueError("DOCX expanded size exceeds the safety limit")
                output.writestr(entry, data)
            for name in sorted(set(overrides) - {entry.filename for entry in entries}):
                data = overrides[name]
                if len(data) > _io.MAX_PART_BYTES:
                    raise ValueError("DOCX part exceeds the safety limit")
                expanded += len(data)
                if expanded > _io.MAX_EXPANDED_BYTES:
                    raise ValueError("DOCX expanded size exceeds the safety limit")
                entry = ZipInfo(name);entry.compress_type = ZIP_DEFLATED
                output.writestr(entry, data)
        data = stream.getvalue()
        if len(data) > _io.MAX_INPUT_BYTES:
            raise ValueError("DOCX output exceeds the safety limit")
        return data

    def save(self, destination):
        theme_parts = self._default_theme_parts()
        self.validate_parts(theme_parts)
        data = self.to_bytes()
        with atomic_output(destination) as temporary:
            temporary.write_bytes(data)
        self._commit_parts(theme_parts)
        if theme_parts:
            self._injected_default_theme = True
