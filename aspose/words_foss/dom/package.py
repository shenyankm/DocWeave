"""Original OPC payloads with lazy, namespace-preserving XML trees."""

from io import BytesIO
from hashlib import sha256
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from defusedxml.minidom import parseString

from aspose.words_foss import _io
from aspose.words_foss._flat_opc import decode, is_flat_opc
from aspose.words_foss._io import (
    atomic_output,
    check_input_size,
    read_bounded,
    validate_docx_archive,
)


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

    def tree(self, name):
        if name not in self._trees:
            self._trees[name] = parseString(self._payloads[name], forbid_dtd=True)
        return self._trees[name]

    def payload(self, name):
        if name in self._dirty:
            return self.tree(name).toxml(encoding="utf-8")
        return self._payloads[name]

    def preservation_report(self):
        """Compare uncompressed part bytes with the loaded package, including resources."""
        report = {"added": [], "modified": [], "removed": [], "unchanged": []}
        for name in sorted(set(self._original_hashes) | set(self._payloads)):
            if name not in self._payloads:
                state = "removed"
            elif name not in self._original_hashes:
                state = "added"
            elif sha256(self.payload(name)).digest() != self._original_hashes[name]:
                state = "modified"
            else:
                state = "unchanged"
            report[state].append(name)
        return report

    def set_parts(self, parts):
        """Commit prebuilt resource parts together, after checking output size limits."""
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
        for name, data in parts.items():
            if name not in self._payloads:
                entry = ZipInfo(name)
                entry.compress_type = ZIP_DEFLATED
                self._entries.append(entry)
            self._payloads[name] = data
            self._trees.pop(name, None)
            self._dirty.discard(name)

    def to_bytes(self):
        stream = BytesIO()
        with ZipFile(stream, "w") as output:
            output.comment = self._comment
            for entry in self._entries:
                output.writestr(entry, self.payload(entry.filename))
        return stream.getvalue()

    def save(self, destination):
        with atomic_output(destination) as temporary:
            temporary.write_bytes(self.to_bytes())
