"""OPC zip-package writer for DOCX files.

Builds the four ancillary parts every Word document needs:

* ``[Content_Types].xml`` — content-type registry.
* ``_rels/.rels`` — package-level relationships.
* ``word/_rels/document.xml.rels`` — document-level relationships
  (hyperlinks, styles, numbering).
* The ZIP archive itself.

Keeping packaging in its own module means the renderer never touches
``zipfile`` directly, and adding new parts (headers, footers, images)
later is a localised change.
"""


import zipfile
from io import BytesIO
from posixpath import dirname, relpath
from urllib.parse import quote
from xml.etree import ElementTree as ET

from defusedxml import DefusedXmlException
from defusedxml.ElementTree import fromstring
from pathlib import Path
from typing import IO, Optional, Union

from aspose.words_foss._opc import related_part_snapshot, relationships_path, resolve_target
from aspose.words_foss._io import check_input_size
from aspose.words_foss.light_document_model import SourceTheme, SourceFontTable

from aspose.words_foss.docx_writer.constants import (
    CT_CORE_PROPS,
    CT_DOCUMENT,
    CT_EXT_PROPS,
    CT_FONT_TABLE,
    CT_FOOTER,
    CT_HEADER,
    CT_NUMBERING,
    CT_RELS,
    CT_SETTINGS,
    CT_STYLES,
    CT_URI,
    PKG_RELS_URI,
    REL_CORE_PROPS,
    REL_EXT_PROPS,
    REL_FONT_TABLE,
    REL_FOOTER,
    REL_HEADER,
    REL_HYPERLINK,
    REL_NUMBERING,
    REL_OFFICE_DOCUMENT,
    REL_SETTINGS,
    REL_STYLES,
)
from aspose.words_foss.docx_writer.doc_props_part import (
    render_app_xml,
    render_core_xml,
    render_font_table_xml,
)
from aspose.words_foss.docx_writer.drawing import REL_IMAGE, ImageEntry
from aspose.words_foss.docx_writer.headers_footers import (
    FOOTER_PART_PATH,
    FOOTER_REL_ID,
    FOOTER_TARGET,
    HEADER_PART_PATH,
    HEADER_REL_ID,
    HEADER_TARGET,
)
from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el, indent_xml
from aspose.words_foss.docx_writer.drawing import _ext_to_content_type


def _content_types(
    has_numbering: bool,
    has_header: bool,
    has_footer: bool,
    has_settings: bool,
    image_extensions: set[str],
    has_theme: bool = False,
) -> str:
    """Build ``[Content_Types].xml``.

    Images live under ``word/media/imageN.<ext>``; we register one
    ``<Default>`` per distinct extension (``png``, ``jpeg`` …) so Word
    resolves their content type via the file suffix instead of needing
    a per-image ``<Override>``.
    """

    children = [
        el("Default", {"Extension": "rels", "ContentType": CT_RELS}),
        el(
            "Default",
            {"Extension": "xml", "ContentType": "application/xml"},
        ),
    ]
    for ext in sorted(image_extensions):
        children.append(
            el(
                "Default",
                {"Extension": ext, "ContentType": _ext_to_content_type(ext)},
            )
        )
    children.extend(
        [
            el(
                "Override",
                {"PartName": "/word/document.xml", "ContentType": CT_DOCUMENT},
            ),
            el(
                "Override",
                {"PartName": "/word/styles.xml", "ContentType": CT_STYLES},
            ),
        ]
    )
    if has_theme:
        children.append(el('Override', {'PartName': '/word/theme/theme1.xml',
                                       'ContentType': 'application/vnd.openxmlformats-officedocument.theme+xml'}))
    if has_numbering:
        children.append(
            el(
                "Override",
                {"PartName": "/word/numbering.xml", "ContentType": CT_NUMBERING},
            )
        )
    if has_header:
        children.append(
            el(
                "Override",
                {"PartName": f"/{HEADER_PART_PATH}", "ContentType": CT_HEADER},
            )
        )
    if has_footer:
        children.append(
            el(
                "Override",
                {"PartName": f"/{FOOTER_PART_PATH}", "ContentType": CT_FOOTER},
            )
        )
    if has_settings:
        children.append(
            el(
                "Override",
                {"PartName": "/word/settings.xml", "ContentType": CT_SETTINGS},
            )
        )
    # Standard ancillary parts — always emitted because MS Word's
    # stricter loader (Word for Mac, Word 2016+, Office Online)
    # rejects packages that omit them even though ECMA-376 marks
    # them "Recommended".
    children.append(
        el("Override", {"PartName": "/word/fontTable.xml", "ContentType": CT_FONT_TABLE})
    )
    children.append(
        el("Override", {"PartName": "/docProps/core.xml", "ContentType": CT_CORE_PROPS})
    )
    children.append(
        el("Override", {"PartName": "/docProps/app.xml", "ContentType": CT_EXT_PROPS})
    )
    return XML_DECL + el("Types", {"xmlns": CT_URI}, children)


def _pkg_rels() -> str:
    """Top-level ``_rels/.rels`` linking the package root to the office
    document plus the two ``docProps/*`` parts MS Word expects to find
    referenced from the package root."""
    children = [
        el(
            "Relationship",
            {
                "Id": "rId1",
                "Type": REL_OFFICE_DOCUMENT,
                "Target": "word/document.xml",
            },
        ),
        el(
            "Relationship",
            {
                "Id": "rIdCoreProps",
                "Type": REL_CORE_PROPS,
                "Target": "docProps/core.xml",
            },
        ),
        el(
            "Relationship",
            {
                "Id": "rIdExtProps",
                "Type": REL_EXT_PROPS,
                "Target": "docProps/app.xml",
            },
        ),
    ]
    return XML_DECL + el("Relationships", {"xmlns": PKG_RELS_URI}, children)


def _doc_rels(
    hyperlinks: dict[str, str],
    has_numbering: bool,
    has_header: bool,
    has_footer: bool,
    has_settings: bool,
    images: list[ImageEntry],
    has_theme: bool = False,
) -> str:
    """Build ``word/_rels/document.xml.rels`` from accumulated relationships."""
    children = [
        el(
            "Relationship",
            {
                "Id": "rIdStyles",
                "Type": REL_STYLES,
                "Target": "styles.xml",
            },
        ),
    ]
    if has_theme:
        children.append(el('Relationship', {'Id': 'rIdTheme',
                                            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
                                            'Target': 'theme/theme1.xml'}))
    if has_numbering:
        children.append(
            el(
                "Relationship",
                {
                    "Id": "rIdNumbering",
                    "Type": REL_NUMBERING,
                    "Target": "numbering.xml",
                },
            )
        )
    if has_header:
        children.append(
            el(
                "Relationship",
                {
                    "Id": HEADER_REL_ID,
                    "Type": REL_HEADER,
                    "Target": HEADER_TARGET,
                },
            )
        )
    if has_footer:
        children.append(
            el(
                "Relationship",
                {
                    "Id": FOOTER_REL_ID,
                    "Type": REL_FOOTER,
                    "Target": FOOTER_TARGET,
                },
            )
        )
    if has_settings:
        children.append(
            el(
                "Relationship",
                {
                    "Id": "rIdSettings",
                    "Type": REL_SETTINGS,
                    "Target": "settings.xml",
                },
            )
        )
    # MS Word looks up the font table via this relationship and
    # falls back to default substitutions when it's missing —
    # several Office for Mac builds *require* the entry to be
    # present even if the fontTable.xml itself is empty.
    children.append(
        el(
            "Relationship",
            {
                "Id": "rIdFontTable",
                "Type": REL_FONT_TABLE,
                "Target": "fontTable.xml",
            },
        )
    )
    for image in images:
        # Targets in ``word/_rels/document.xml.rels`` are relative to
        # ``word/`` — strip the leading ``word/`` from media_path.
        target = image.media_path.removeprefix("word/")
        children.append(
            el(
                "Relationship",
                {
                    "Id": image.rel_id,
                    "Type": REL_IMAGE,
                    "Target": target,
                },
            )
        )
    for url, rid in hyperlinks.items():
        children.append(
            el(
                "Relationship",
                {
                    "Id": rid,
                    "Type": REL_HYPERLINK,
                    "Target": url,
                    "TargetMode": "External",
                },
            )
        )
    return XML_DECL + el("Relationships", {"xmlns": PKG_RELS_URI}, children)


def _hf_part_rels(hyperlinks: dict[str, str], images: list[ImageEntry]) -> str:
    """Build a sidecar ``.rels`` file for a header / footer part.

    Headers and footers can carry external hyperlinks and embedded
    images just like the document body; each part needs its own
    ``<Target>`` table because ``r:id`` references are resolved
    against the part's own rels file, not the document's.  Image
    targets are written relative to the part's location
    (``word/header1.xml`` resolves ``Target="media/image.png"`` as
    ``word/media/image.png``).
    """
    children: list[str] = []
    for url, rid in hyperlinks.items():
        children.append(
            el(
                "Relationship",
                {
                    "Id": rid,
                    "Type": REL_HYPERLINK,
                    "Target": url,
                    "TargetMode": "External",
                },
            )
        )
    for image in images:
        target = image.media_path.removeprefix("word/")
        children.append(
            el(
                "Relationship",
                {
                    "Id": image.rel_id,
                    "Type": REL_IMAGE,
                    "Target": target,
                },
            )
        )
    return XML_DECL + el("Relationships", {"xmlns": PKG_RELS_URI}, children)


def _import_source_resources(source: SourceTheme | SourceFontTable, text_parts, binary_parts, *,
                             main_destination, resource_directory):
    """Validate the captured graph, then rebind it under an unused resource directory."""
    check_input_size(max(sum(len(part.data) for part in source.parts),
                         len(source.data) + len(source.relationships or b'')))
    if not source.parts:
        if source.relationships is not None:
            root = fromstring(source.relationships, forbid_dtd=True)
            if root.tag != f'{{{PKG_RELS_URI}}}Relationships':
                raise ValueError('Expected OPC source relationships')
            if len(root):
                raise ValueError('Source relationship resources are missing from the source snapshot')
        if any(key.startswith('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}')
               for node in fromstring(source.data, forbid_dtd=True).iter() for key in node.attrib):
            raise ValueError('Source relationship resources are missing from the source snapshot')
        return
    parts = {}
    part_names = set()
    types = ET.Element(f'{{{CT_URI}}}Types')
    for part in source.parts:
        if (resolve_target('', quote('/' + part.name, safe='/')) != part.name
                or part.name.casefold() in part_names
                or part.name == '[Content_Types].xml'):
            raise ValueError('Invalid or duplicate source resource part name')
        parts[part.name] = part.data
        part_names.add(part.name.casefold())
        ET.SubElement(types, f'{{{CT_URI}}}Override', PartName=quote('/' + part.name, safe='/'),
                      ContentType=part.content_type)
    if parts.get(source.part_name) != source.data or parts.get(relationships_path(source.part_name)) != source.relationships:
        raise ValueError('Source resource snapshot declarations disagree')
    stream = BytesIO()
    with zipfile.ZipFile(stream, 'w') as archive:
        archive.writestr('[Content_Types].xml', ET.tostring(types))
        for name, data in parts.items():
            archive.writestr(name, data)
    with zipfile.ZipFile(stream) as archive:
        graph = related_part_snapshot(archive, source.part_name)
    if {name for name, _, _ in graph} != set(parts):
        raise ValueError('Unrelated parts in the source resource snapshot')
    occupied = {name.casefold() for name, _ in text_parts + binary_parts}
    number = 1
    while any(name.startswith(f'{resource_directory}{number}/') for name in occupied):
        number += 1
    prefix = f'{resource_directory}{number}/'
    mapping = {name: prefix + name for name, _, _ in graph if not name.endswith('.rels')}
    mapping[source.part_name] = main_destination
    content_types = fromstring(text_parts[0][1], forbid_dtd=True)
    for name, content_type, data in graph:
        if name.endswith('.rels'):
            continue
        destination = mapping[name]
        rels_name = relationships_path(name)
        root = fromstring(parts[rels_name], forbid_dtd=True) if rels_name in parts else None
        ids = {item.get('Id') for item in root} if root is not None else set()
        if name == source.part_name or content_type.endswith('+xml') or content_type in {'application/xml', 'text/xml'}:
            xml = fromstring(data, forbid_dtd=True)
            if any(value not in ids for node in xml.iter() for key, value in node.attrib.items()
                   if key.startswith('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}')):
                raise ValueError('Missing source resource relationship ID')
        if root is not None:
            for item in root:
                if item.get('TargetMode') != 'External':
                    target = mapping[resolve_target(name, item.get('Target'))]
                    item.set('Target', quote(relpath(target, dirname(destination)), safe='/'))
            binary_parts.append((relationships_path(destination), ET.tostring(root, encoding='utf-8')))
        if name != source.part_name:
            binary_parts.append((destination, data))
            ET.SubElement(content_types, f'{{{CT_URI}}}Override', PartName=quote('/' + destination, safe='/'),
                          ContentType=content_type)
    text_parts[0] = ('[Content_Types].xml', ET.tostring(content_types, encoding='unicode'))

def _strip_embedded_fonts(text_parts, binary_parts, start):
    """Strip embedding only after the complete original resource graph was validated."""
    parts = dict(binary_parts[start:])
    root_name = 'word/fontTable.xml'
    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    root = fromstring(parts[root_name], forbid_dtd=True)
    for font in root.findall(w + 'font'):
        for child in list(font):
            if child.tag in {w + 'embedRegular', w + 'embedBold', w + 'embedItalic', w + 'embedBoldItalic'}:
                font.remove(child)
    parts[root_name] = ET.tostring(root, encoding='utf-8')
    rels_name = relationships_path(root_name)
    if rels_name in parts:
        edges = fromstring(parts[rels_name], forbid_dtd=True)
        for edge in list(edges):
            if edge.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font':
                edges.remove(edge)
        if len(edges):
            parts[rels_name] = ET.tostring(edges, encoding='utf-8')
        else:
            del parts[rels_name]
    keep = set()
    pending = [root_name]
    while pending:
        name = pending.pop()
        if name in keep:
            continue
        keep.add(name)
        rels = relationships_path(name)
        if rels in parts:
            keep.add(rels)
            for edge in fromstring(parts[rels], forbid_dtd=True):
                if edge.get('TargetMode') != 'External':
                    pending.append(resolve_target(name, edge.get('Target')))
    removed = {name for name, _ in binary_parts[start:]} - keep
    binary_parts[start:] = [(name, parts[name]) for name, _ in binary_parts[start:] if name in keep]
    types = fromstring(text_parts[0][1], forbid_dtd=True)
    for item in list(types):
        if item.tag == f'{{{CT_URI}}}Override' and resolve_target('', item.get('PartName')) in removed:
            types.remove(item)
    text_parts[0] = ('[Content_Types].xml', ET.tostring(types, encoding='unicode'))


def _import_theme_resources(source: SourceTheme, text_parts, binary_parts):
    _import_source_resources(source, text_parts, binary_parts,
                             main_destination='word/theme/theme1.xml',
                             resource_directory='word/theme/resources')



def _build_parts(
    *,
    document_xml: str,
    styles_xml: str,
    numbering_xml: Optional[str],
    settings_xml: Optional[str],
    header_xml: Optional[str],
    header_hyperlinks: dict[str, str],
    header_images: list[ImageEntry],
    footer_xml: Optional[str],
    footer_hyperlinks: dict[str, str],
    footer_images: list[ImageEntry],
    hyperlinks: dict[str, str],
    images: list[ImageEntry],
    theme_xml: bytes | None = None,
    source_theme: SourceTheme | None = None,
    source_font_table: SourceFontTable | None = None,
    embed_fonts: bool = False,
) -> tuple[list[tuple[str, str]], list[tuple[str, bytes]]]:
    """Assemble ordered (text_parts, binary_parts) lists for the zip."""
    if source_theme is not None and source_theme.data != theme_xml:
        raise ValueError('Source theme XML disagrees with its resource snapshot')
    if theme_xml is not None:
        try:
            theme = fromstring(theme_xml, forbid_dtd=True)
        except (ET.ParseError, DefusedXmlException):
            raise ValueError('Invalid source theme XML') from None
        if theme.tag != '{http://schemas.openxmlformats.org/drawingml/2006/main}theme':
            raise ValueError('Expected a DrawingML theme')
        if source_theme is None and any(key.startswith('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}')
               for node in theme.iter() for key in node.attrib):
            raise NotImplementedError('Theme relationship resources are not yet supported by conversion')
    has_numbering = bool(numbering_xml)
    has_header = bool(header_xml)
    has_footer = bool(footer_xml)
    has_settings = bool(settings_xml)
    all_images = images + header_images + footer_images
    image_extensions = {image.media_path.rsplit(".", 1)[1].lower() for image in all_images}
    text_parts: list[tuple[str, str]] = [
        (
            "[Content_Types].xml",
            _content_types(has_numbering, has_header, has_footer, has_settings, image_extensions, has_theme=theme_xml is not None),
        ),
        ("_rels/.rels", _pkg_rels()),
        ("word/document.xml", document_xml),
        (
            "word/_rels/document.xml.rels",
            _doc_rels(hyperlinks, has_numbering, has_header, has_footer, has_settings, images, has_theme=theme_xml is not None),
        ),
        ("word/styles.xml", styles_xml),
    ]
    if has_numbering and numbering_xml:
        text_parts.append(("word/numbering.xml", numbering_xml))
    if has_settings and settings_xml:
        text_parts.append(("word/settings.xml", settings_xml))
    # Ancillary parts (always emitted) — see ``doc_props_part`` for why.
    if source_font_table is None:
        text_parts.append(("word/fontTable.xml", render_font_table_xml()))
    text_parts.append(("docProps/core.xml", render_core_xml()))
    text_parts.append(("docProps/app.xml", render_app_xml()))
    if has_header and header_xml:
        text_parts.append((HEADER_PART_PATH, header_xml))
        if header_hyperlinks or header_images:
            text_parts.append(
                (
                    f"word/_rels/{HEADER_TARGET}.rels",
                    _hf_part_rels(header_hyperlinks, header_images),
                )
            )
    if has_footer and footer_xml:
        text_parts.append((FOOTER_PART_PATH, footer_xml))
        if footer_hyperlinks or footer_images:
            text_parts.append(
                (
                    f"word/_rels/{FOOTER_TARGET}.rels",
                    _hf_part_rels(footer_hyperlinks, footer_images),
                )
            )
    binary_parts: list[tuple[str, bytes]] = [
        (image.media_path, image.image_bytes) for image in all_images
    ]
    if source_font_table is not None:
        try:
            check_input_size(max(sum(len(part.data) for part in source_font_table.parts),
                                 len(source_font_table.data) + len(source_font_table.relationships or b'')))
            fonts = fromstring(source_font_table.data, forbid_dtd=True)
            w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            if fonts.tag != w + 'fonts':
                raise ValueError('Expected a WordprocessingML font table')
            from aspose.words_foss.fonts import _root
            _root(source_font_table.data)
            font_start = len(binary_parts)
            binary_parts.append(('word/fontTable.xml', source_font_table.data))
            _import_source_resources(source_font_table, text_parts, binary_parts,
                                     main_destination='word/fontTable.xml',
                                     resource_directory='word/fonts/resources')
            if not embed_fonts:
                _strip_embedded_fonts(text_parts, binary_parts, font_start)
        except (ET.ParseError, DefusedXmlException):
            raise ValueError('Invalid source font table resource XML') from None
    if theme_xml is not None:
        binary_parts.append(('word/theme/theme1.xml', theme_xml))
        if source_theme is not None:
            try:
                _import_theme_resources(source_theme, text_parts, binary_parts)
            except (ET.ParseError, DefusedXmlException):
                raise ValueError('Invalid source theme resource XML') from None
    return text_parts, binary_parts


def write_docx_package(
    target: Union[str, Path, IO[bytes]],
    *,
    document_xml: str,
    styles_xml: str,
    numbering_xml: Optional[str],
    settings_xml: Optional[str] = None,
    theme_xml: bytes | None = None,
    source_theme: SourceTheme | None = None,
    source_font_table: SourceFontTable | None = None,
    embed_fonts: bool = False,
    hyperlinks: dict[str, str],
    images: Optional[list[ImageEntry]] = None,
    header_xml: Optional[str] = None,
    header_hyperlinks: Optional[dict[str, str]] = None,
    header_images: Optional[list[ImageEntry]] = None,
    footer_xml: Optional[str] = None,
    footer_hyperlinks: Optional[dict[str, str]] = None,
    footer_images: Optional[list[ImageEntry]] = None,
    compression: int = zipfile.ZIP_DEFLATED,
    compresslevel: Optional[int] = None,
    allow_zip64: bool = False,
    pretty_format: bool = False,
) -> None:
    """Write the assembled parts to ``target`` as a DOCX zip.

    ``target`` may be a path (``str`` / :class:`Path`) or any binary
    file-like object — :class:`zipfile.ZipFile` accepts both, so
    callers wanting in-memory output can pass an :class:`io.BytesIO`.
    Path targets get their parent directory auto-created; file-like
    targets are written to as-is.

    ``compresslevel`` is forwarded to :class:`zipfile.ZipFile` and
    controls the deflate compression effort (0–9).  ``None`` lets the
    standard library choose its default.

    ``allow_zip64`` enables ZIP64 extensions for archives exceeding the
    classic 4 GB / 65 535-entry ZIP limits.
    """
    text_parts, binary_parts = _build_parts(
        theme_xml=theme_xml,
        source_theme=source_theme,
        source_font_table=source_font_table,
        embed_fonts=embed_fonts,
        document_xml=document_xml,
        styles_xml=styles_xml,
        numbering_xml=numbering_xml,
        settings_xml=settings_xml,
        header_xml=header_xml,
        header_hyperlinks=header_hyperlinks or {},
        header_images=list(header_images or []),
        footer_xml=footer_xml,
        footer_hyperlinks=footer_hyperlinks or {},
        footer_images=list(footer_images or []),
        hyperlinks=hyperlinks,
        images=list(images or []),
    )

    if pretty_format:
        text_parts = [(name, indent_xml(content)) for name, content in text_parts]

    if isinstance(target, (str, Path)):
        path = Path(target)
        path.parent.mkdir(parents=True, exist_ok=True)
        target = path  # zipfile resolves Path internally

    zf_kwargs: dict[str, object] = {}
    if compresslevel is not None:
        zf_kwargs["compresslevel"] = compresslevel

    with zipfile.ZipFile(
        target, "w", compression, allowZip64=allow_zip64, **zf_kwargs
    ) as zf:
        for name, content in text_parts:
            zf.writestr(name, content)
        for name, payload in binary_parts:
            zf.writestr(name, payload)
