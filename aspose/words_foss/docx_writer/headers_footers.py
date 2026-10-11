"""Render explicit section header/footer stories with independent resources."""


from typing import Mapping, Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer.bookmarks import BookmarkState
from aspose.words_foss.docx_writer.constants import MC_URI, O_URI, R_URI, V_URI, W_URI
from aspose.words_foss.docx_writer.drawing import A_URI, PIC_URI, WP_URI, WPS_URI, ImageRenderState
from aspose.words_foss.docx_writer.paragraphs import render_paragraph
from aspose.words_foss.docx_writer.tables import render_table
from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el


def _render_part(
    children: list[ldm.Paragraph | ldm.Table],
    rels: dict,
    *,
    root_tag: str,
    style_pf_map: Mapping[str, ldm.ParagraphFormat],
    style_id_map: Mapping[str, str],
    style_font_map: Mapping[str, ldm.Font],
    image_state: "ImageRenderState | None" = None,
    num_id_map: Mapping[int, int] | None = None,
) -> str:
    """Render a single ``<w:hdr>`` or ``<w:ftr>`` XML payload.

    ``children`` is an ordered list of paragraphs and tables preserving
    the original document order.  Empty lists still emit the root element
    with a single empty ``<w:p/>`` — Word rejects header / footer parts
    that contain no paragraphs.

    ``image_state`` accumulates images referenced from inside the part;
    the caller passes a fresh instance per part so the resulting
    ``ImageEntry`` rows can be wired into the part's sidecar ``.rels``
    file and added to ``word/media/`` with non-colliding filenames.
    """
    rendered: list[str] = []
    if image_state is None:
        image_state = ImageRenderState()
    bookmark_state = BookmarkState()
    for child in children:
        if isinstance(child, ldm.Paragraph):
            rendered.append(
                render_paragraph(
                    child,
                    rels,
                    num_id_map=num_id_map or {},
                    image_state=image_state,
                    bookmark_state=bookmark_state,
                    style_pf_map=style_pf_map,
                    style_id_map=style_id_map,
                    style_font_map=style_font_map,
                )
            )
        elif isinstance(child, ldm.Table):
            rendered.append(
                render_table(
                    child,
                    rels,
                    num_id_map=num_id_map or {},
                    image_state=image_state,
                    bookmark_state=bookmark_state,
                    style_pf_map=style_pf_map,
                    style_id_map=style_id_map,
                    style_font_map=style_font_map,
                )
            )
    for _, bm_id in bookmark_state.drain_open():
        rendered.append(el("w:bookmarkEnd", {"w:id": bm_id}))
    if not rendered:
        rendered.append(el("w:p"))
    root = el(
        root_tag,
        {
            "xmlns:w": W_URI,
            "xmlns:r": R_URI,
            "xmlns:wp": WP_URI,
            "xmlns:a": A_URI,
            "xmlns:pic": PIC_URI,
            "xmlns:wps": WPS_URI,
            "xmlns:mc": MC_URI,
            "xmlns:v": V_URI,
            "xmlns:o": O_URI,
            # ``mc:Ignorable`` only lists prefixes whose namespace is
            # declared above; Word rejects the package otherwise.
            "mc:Ignorable": "wps",
        },
        rendered,
    )
    return XML_DECL + root


def render_header_xml(
    children: list[ldm.Paragraph | ldm.Table],
    rels: dict,
    style_pf_map: Mapping[str, ldm.ParagraphFormat],
    style_id_map: Mapping[str, str],
    style_font_map: Mapping[str, ldm.Font],
    image_state: "ImageRenderState | None" = None,
    num_id_map: Mapping[int, int] | None = None,
) -> str:
    return _render_part(
        children,
        rels,
        root_tag="w:hdr",
        style_pf_map=style_pf_map,
        style_id_map=style_id_map,
        style_font_map=style_font_map,
        image_state=image_state,
        num_id_map=num_id_map,
    )


def render_footer_xml(
    children: list[ldm.Paragraph | ldm.Table],
    rels: dict,
    style_pf_map: Mapping[str, ldm.ParagraphFormat],
    style_id_map: Mapping[str, str],
    style_font_map: Mapping[str, ldm.Font],
    image_state: "ImageRenderState | None" = None,
    num_id_map: Mapping[int, int] | None = None,
) -> str:
    return _render_part(
        children,
        rels,
        root_tag="w:ftr",
        style_pf_map=style_pf_map,
        style_id_map=style_id_map,
        style_font_map=style_font_map,
        image_state=image_state,
        num_id_map=num_id_map,
    )


def header_reference(rid: str) -> str:
    """``<w:headerReference w:type="default" r:id="..."/>`` for ``<w:sectPr>``."""
    return el("w:headerReference", {"w:type": "default", "r:id": rid})


def footer_reference(rid: str) -> str:
    """``<w:footerReference w:type="default" r:id="..."/>`` for ``<w:sectPr>``."""
    return el("w:footerReference", {"w:type": "default", "r:id": rid})


HEADER_REL_ID = "rIdHeader1"
FOOTER_REL_ID = "rIdFooter1"

# Filenames inside the ``.docx`` zip — kept in sync with the ``Target``
# attribute in ``word/_rels/document.xml.rels``.
HEADER_PART_PATH = "word/header1.xml"
FOOTER_PART_PATH = "word/footer1.xml"
HEADER_TARGET = "header1.xml"
FOOTER_TARGET = "footer1.xml"


def needs_header(doc: ldm.Document) -> bool:
    if doc.header_paragraphs:
        return True
    for sec in doc.sections:
        for hf in sec.headers_footers:
            if hf.header_footer_type == 0 and hf.children:
                return True
    return False


def needs_footer(doc: ldm.Document) -> bool:
    if doc.footer_paragraphs:
        return True
    for sec in doc.sections:
        for hf in sec.headers_footers:
            if hf.header_footer_type == 1 and hf.children:
                return True
    return False


# ----------------------------------------------------------------------
# Filtering helper: the writer threads ``unused_kw`` through ``render_paragraph``
# but they aren't used by the headers/footers writer; this stub keeps the
# import surface symmetrical with ``document_part``.
# ----------------------------------------------------------------------


def _unused() -> Optional[str]:  # pragma: no cover - documentation glue
    return None


def render_section_stories(doc, style_pf_map, style_id_map, style_font_map):
    """Render each explicit header/footer with its own part and relationships."""
    from .drawing import pt_to_mm
    from .numbering_part import build_num_id_map

    parts, references = [], []
    num_id_map = build_num_id_map(doc.lists)
    counts = {"header": 0, "footer": 0}
    variants = {0: "default", 1: "default", 2: "first", 3: "first", 4: "even", 5: "even"}
    for section in doc.sections:
        refs, seen = [], set()
        for story in section.headers_footers:
            kind = story.header_footer_type
            if kind not in variants:
                raise ValueError("Unsupported header/footer type")
            if kind in seen:
                raise ValueError("Duplicate section header/footer type")
            seen.add(kind)
            role = "footer" if kind % 2 else "header"
            counts[role] += 1
            number = counts[role]
            target, rid = f"{role}{number}.xml", f"rId{role.title()}{number}"
            refs.append(el(f"w:{role}Reference", {"w:type": variants[kind], "r:id": rid}))
            setup = section.page_setup
            baseline = (setup.page_height - setup.footer_distance if role == "footer"
                        else setup.header_distance)
            images = ImageRenderState(media_name_prefix=f"{role}{number}_image",
                rel_id_prefix=f"rId{role.title()}{number}Img", doc_pr_offset=(len(parts) + 1) * 100000,
                section_top_margin_mm=pt_to_mm(baseline),
                section_left_margin_mm=pt_to_mm(setup.left_margin))
            links = {}
            renderer = render_footer_xml if role == "footer" else render_header_xml
            xml = renderer(story.children, links, style_pf_map, style_id_map,
                           style_font_map, image_state=images, num_id_map=num_id_map)
            parts.append((target, rid, role, xml, links, images.images))
        references.append("".join(refs))
    return parts, references
