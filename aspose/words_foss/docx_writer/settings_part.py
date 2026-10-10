"""Generates ``word/settings.xml``.

Always emitted: the Aspose blank baseline carries a ``<w:compat>``
block that Word2010+ expects on every package.  LDM-tracked values
(default tabs and manual-break justification) are added when they
differ from Word's compiled-in default.
"""


from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer.blank_template import settings_compat
from aspose.words_foss.docx_writer.constants import W_URI, pt_to_twips
from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el


_DEFAULT_TAB_STOP_PT = 36.0


def needs_settings(doc: ldm.Document) -> bool:
    """Always True — the compat block belongs in every package."""
    return True


def render_settings_xml(doc: ldm.Document) -> str:
    """Build ``word/settings.xml`` payload."""
    children: list[str] = []
    if abs(doc.default_tab_stop - _DEFAULT_TAB_STOP_PT) > 0.01:
        children.append(
            el("w:defaultTabStop", {"w:val": pt_to_twips(doc.default_tab_stop)})
        )
    compat = settings_compat()
    tracked = el("w:compatSetting", {"w:name": "compatibilityMode",
        "w:uri": "http://schemas.microsoft.com/office/word", "w:val": str(doc.compatibility_mode)})
    if doc.do_not_expand_shift_return:
        tracked += el("w:doNotExpandShiftReturn")
    compat = (compat or "<w:compat></w:compat>").replace("</w:compat>", tracked + "</w:compat>")
    children.append(compat)
    if doc.theme_font_languages is not None:
        languages = doc.theme_font_languages
        attrs = {key: value for key, value in (
            ('w:val', languages.latin), ('w:eastAsia', languages.east_asian),
            ('w:bidi', languages.complex_script)) if value is not None}
        children.append(el('w:themeFontLang', attrs))
    if doc.font_embedding is not None:
        for field, tag in [('embed_true_type_fonts', 'embedTrueTypeFonts'),
                           ('do_not_embed_system_fonts', 'doNotEmbedSystemFonts'),
                           ('embed_system_fonts', 'embedSystemFonts'),
                           ('save_subset_fonts', 'saveSubsetFonts')]:
            value = getattr(doc.font_embedding, field)
            if value is not None:
                children.append(el('w:' + tag, None if value else {'w:val': '0'}))
    root = el("w:settings", {"xmlns:w": W_URI}, children)
    return XML_DECL + root
