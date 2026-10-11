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
    return ET.tostring(block,encoding='unicode')


def _render_preferences(doc, compat):
    from aspose.words_foss.settings import _COMPAT_TAGS, WORD_URI
    from aspose.words_foss.diagnostics import ContentLossWarning, warn
    state = ldm.CompatibilityOptionsState.model_validate(doc.compatibility_options_state.model_dump())
    namespace = '{'+W_URI+'}'
    root = fromstring('<root xmlns:w="'+W_URI+'">'+compat+'</root>', forbid_dtd=True)
    block = root[0] if len(root) else ET.Element(namespace+'compat')
    removed = False
    for name,tag in _COMPAT_TAGS.items():
        value = doc.do_not_expand_shift_return if name == 'do_not_expand_shift_return' else state.flags[name]
        if state.optimized_for is None and name not in state.edited_flags:
            if value:
                continue
        nodes = block.findall(namespace+tag)
        winner = nodes[-1] if value and nodes else None
        for child in nodes:
            if child is not winner:
                block.remove(child)
                removed = True
        if value:
            if winner is not None:
                winner.attrib.pop(namespace+'val',None)
            else:
                ET.SubElement(block,namespace+tag)
    if state.optimized_for is not None:
        original = doc.source_settings
        declared = False
        if original is not None and original.data is not None:
            source = fromstring(original.data, forbid_dtd=True)
            declared = any(child.tag == namespace + 'compatSetting' for child in source)
        for child in list(block):
            if child.tag == namespace + 'compatSetting':
                block.remove(child)
                removed = True
        version = state.optimized_for
        if version >= 5 or declared:
            values = []
            if version >= 2:
                values.append(('compatibilityMode', str(doc.compatibility_mode)))
            values.extend((name, '0' if name == 'overrideTableStyleFontSizeAndJustification'
                           and version == 3 else '1') for name in (
                'overrideTableStyleFontSizeAndJustification', 'enableOpenTypeFeatures',
                'doNotFlipMirrorIndents', 'differentiateMultirowTableHeaders'))
            for name, value in values:
                ET.SubElement(block, namespace + 'compatSetting', {
                    namespace + 'name': name, namespace + 'uri': WORD_URI,
                    namespace + 'val': value})
    if removed:
        warn('Compatibility declarations were normalized for saving.', ContentLossWarning,
             code='settings.compatibility_options_normalized')
    return ET.tostring(block,encoding='unicode')


def _complete_compatibility_settings(compat: str) -> str:
    """A loaded typed setting collection supplies native newer defaults by name."""
    namespace = '{' + W_URI + '}'
    wrapper = fromstring('<root xmlns:w="' + W_URI + '">' + compat + '</root>', forbid_dtd=True)
    if not len(wrapper):
        return compat
    block = wrapper[0]
    records = block.findall(namespace + 'compatSetting')
    if records:
        import re
        from math import isfinite
        names = {record.get(namespace + 'name') for record in records}
        mode_value = next((record.get(namespace + 'val', '') for record in reversed(records)
                           if record.get(namespace + 'name') == 'compatibilityMode'), '')
        number = float(mode_value) if re.fullmatch(
            r'[ \t\r\n]*[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?[ \t\r\n]*',
            mode_value) else 0.0
        mode = int(number) % 4294967296 if isfinite(number) else 0
        for name in ('overrideTableStyleFontSizeAndJustification', 'enableOpenTypeFeatures',
                     'doNotFlipMirrorIndents', 'differentiateMultirowTableHeaders'):
            if name not in names:
                ET.SubElement(block, namespace + 'compatSetting', {
                    namespace + 'name': name,
                    namespace + 'uri': 'http://schemas.microsoft.com/office/word',
                    namespace + 'val': '0' if name == 'overrideTableStyleFontSizeAndJustification'
                    and mode not in (11, 15) else '1'})
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
    if doc.compatibility_options_state is not None:
        compat = _render_preferences(doc,compat)
    compat = _normalize_boolean_declarations(compat)
    children.append(_complete_compatibility_settings(compat))
    state = doc.compatibility_options_state
    source = doc.source_settings
    if state is not None and (state.optimized_for is not None or 'ui_compat_97_to_2003' in state.edited_flags):
        ui = state.flags['ui_compat_97_to_2003']
    else:
        ui = source.ui_compat_97_to_2003 if source is not None else None
    ui_node = None
    if source is not None and source.ui_compat_data is not None:
        ui_node = fromstring(source.ui_compat_data,forbid_dtd=True)
        if ui_node.tag != '{'+W_URI+'}uiCompat97To2003':
            raise ValueError('Source compatibility UI XML must contain w:uiCompat97To2003')
        if not ui:
            from aspose.words_foss.diagnostics import ContentLossWarning, warn
            warn('Compatibility declarations were normalized for saving.', ContentLossWarning,
                 code='settings.compatibility_options_normalized')
    if ui:
        if ui_node is not None:
            ui_node.attrib.pop('{'+W_URI+'}val',None)
            children.append(ET.tostring(ui_node,encoding='unicode'))
        else:
            children.append(el('w:uiCompat97To2003'))
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
