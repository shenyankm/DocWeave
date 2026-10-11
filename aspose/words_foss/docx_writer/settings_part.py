"""Generates ``word/settings.xml``.

Loaded compatibility declarations retain their source order. Programmatic
documents keep the separate blank-template baseline; other settings use
the existing model fields.
"""


from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer.blank_template import settings_compat
from aspose.words_foss.docx_writer.constants import W_URI, pt_to_twips
from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el
from defusedxml.ElementTree import fromstring
from xml.etree import ElementTree as ET


_DEFAULT_TAB_STOP_PT = 36.0

# The recognized on/off compatibility declarations in WordprocessingML.
_COMPAT_BOOLEAN_TAGS = (
    'adjustLineHeightInTable',
    'alignTablesRowByRow',
    'allowSpaceOfSameStyleInTable',
    'applyBreakingRules',
    'autoSpaceLikeWord95',
    'autofitToFirstFixedWidthCell',
    'balanceSingleByteDoubleByteWidth',
    'cachedColBalance',
    'convMailMergeEsc',
    'displayHangulFixedWidth',
    'doNotAutofitConstrainedTables',
    'doNotBreakConstrainedForcedTable',
    'doNotBreakWrappedTables',
    'doNotExpandShiftReturn',
    'doNotLeaveBackslashAlone',
    'doNotSnapToGridInCell',
    'doNotSuppressIndentation',
    'doNotSuppressParagraphBorders',
    'doNotUseEastAsianBreakRules',
    'doNotUseHTMLParagraphAutoSpacing',
    'doNotUseIndentAsNumberingTabStop',
    'doNotVertAlignCellWithSp',
    'doNotVertAlignInTxbx',
    'doNotWrapTextWithPunct',
    'footnoteLayoutLikeWW8',
    'forgetLastTabAlignment',
    'growAutofit',
    'layoutRawTableWidth',
    'layoutTableRowsApart',
    'lineWrapLikeWord6',
    'mwSmallCaps',
    'noColumnBalance',
    'noExtraLineSpacing',
    'noLeading',
    'noSpaceRaiseLower',
    'noTabHangInd',
    'printBodyTextBeforeHeader',
    'printColBlack',
    'selectFldWithFirstOrLastChar',
    'shapeLayoutLikeWW8',
    'showBreaksInFrames',
    'spaceForUL',
    'spacingInWholePoints',
    'splitPgBreakAndParaMark',
    'subFontBySize',
    'suppressBottomSpacing',
    'suppressSpBfAfterPgBrk',
    'suppressSpacingAtTopOfPage',
    'suppressTopSpacing',
    'suppressTopSpacingWP',
    'swapBordersFacingPages',
    'truncateFontHeightsLikeWP6',
    'ulTrailSpace',
    'underlineTabInNumList',
    'useAltKinsokuLineBreakRules',
    'useAnsiKerningPairs',
    'useFELayout',
    'useNormalStyleForList',
    'usePrinterMetrics',
    'useSingleBorderforContiguousCells',
    'useWord2002TableStyleRules',
    'useWord97LineBreakRules',
    'wpJustification',
    'wpSpaceWidth',
    'wrapTrailSpaces',
)



def _render_source_compat(doc: ldm.Document) -> str:
    source = doc.source_settings
    namespace = '{' + W_URI + '}'
    root = fromstring(source.data, forbid_dtd=True) if source.data is not None else ET.Element(namespace + 'compat')
    if root.tag != namespace + 'compat':
        raise ValueError('Source compatibility XML must contain w:compat')
    if doc.compatibility_mode != source.compatibility_mode:
        for child in list(root):
            if (child.tag == namespace + 'compatSetting' and
                    child.get(namespace + 'name') == 'compatibilityMode' and
                    child.get(namespace + 'uri') == 'http://schemas.microsoft.com/office/word'):
                root.remove(child)
        ET.SubElement(root, namespace + 'compatSetting', {
            namespace + 'name': 'compatibilityMode',
            namespace + 'uri': 'http://schemas.microsoft.com/office/word',
            namespace + 'val': str(doc.compatibility_mode)})
    if doc.do_not_expand_shift_return != source.do_not_expand_shift_return:
        for child in list(root):
            if child.tag == namespace + 'doNotExpandShiftReturn':
                root.remove(child)
        ET.SubElement(root, namespace + 'doNotExpandShiftReturn', {
            namespace + 'val': '1' if doc.do_not_expand_shift_return else '0'})
    if source.data is None and not len(root):
        return ''
    return ET.tostring(root, encoding='unicode')


def needs_settings(doc: ldm.Document) -> bool:
    """The writer emits a settings part for its supported model fields."""
    return True


def _normalize_boolean_declarations(compat: str) -> str:
    """Project native last-winner/false omission without changing the snapshot."""
    from aspose.words_foss.docx_reader.utils import parse_onoff
    from aspose.words_foss.diagnostics import ContentLossWarning, warn

    wrapper = fromstring('<root xmlns:w="' + W_URI + '">' + compat + '</root>', forbid_dtd=True)
    if not len(wrapper):
        return compat
    block = wrapper[0]
    removed = False
    for tag in _COMPAT_BOOLEAN_TAGS:
        nodes = block.findall('{' + W_URI + '}' + tag)
        if not nodes:
            continue
        winner = nodes[-1] if parse_onoff(nodes[-1]) else None
        for node in nodes:
            if node is not winner:
                block.remove(node)
                removed = True
    if removed:
        warn('Compatibility declarations were normalized for saving.', ContentLossWarning,
             code='settings.compatibility_options_normalized')
    return ET.tostring(block, encoding='unicode')


def render_settings_xml(doc: ldm.Document) -> str:
    """Build ``word/settings.xml`` payload."""
    children: list[str] = []
    if abs(doc.default_tab_stop - _DEFAULT_TAB_STOP_PT) > 0.01:
        children.append(
            el("w:defaultTabStop", {"w:val": pt_to_twips(doc.default_tab_stop)})
        )
    if doc.source_settings is not None:
        compat = _render_source_compat(doc)
    else:
        compat = settings_compat()
        tracked = el("w:compatSetting", {"w:name": "compatibilityMode",
            "w:uri": "http://schemas.microsoft.com/office/word", "w:val": str(doc.compatibility_mode)})
        if doc.do_not_expand_shift_return:
            tracked += el("w:doNotExpandShiftReturn")
        compat = (compat or "<w:compat></w:compat>").replace("</w:compat>", tracked + "</w:compat>")
    children.append(_normalize_boolean_declarations(compat))
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
