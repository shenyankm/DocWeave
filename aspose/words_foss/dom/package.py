"""Original OPC payloads with lazy, namespace-preserving XML trees."""

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from defusedxml.minidom import parseString

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
        with ZipFile(BytesIO(data)) as archive:
            validate_docx_archive(archive)
            self._entries = archive.infolist()
            self._payloads = {entry.filename: archive.read(entry) for entry in self._entries}
            self._comment = archive.comment
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
