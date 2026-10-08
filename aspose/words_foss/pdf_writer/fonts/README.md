# Bundled PDF fonts — Unofficial Fork Enhancement

These resources were added by this repository's unofficial enhanced fork of
Aspose.Words FOSS for Python. They were not present in the imported upstream
revision `2d2efee2787cb9e56d071d17f8d7b740dce8b784`; they are not official
Aspose fonts or a claim about fonts in later upstream releases. See the
[project's fork status and differences](../../../../README.md#fork-status-and-upstream-differences).

`DocumentSansSC-*.woff` are losslessly compressed, renamed derivatives of **Noto Sans SC 2.004**,
licensed under the SIL Open Font License 1.1; see `OFL.txt`. The original
copyright and license also remain in each font's `name` table.

Source: https://github.com/notofonts/noto-cjk/blob/main/Sans/Variable/TTF/Subset/NotoSansSC-VF.ttf

Source SHA-256:
`d68bafcb48a2707749396aa12bbbd833cb70401f3a9a689fd2902c7e0d295964`

The static Regular and Bold faces were generated with fontTools 4.66.1
(`fontTools.varLib.instancer.instantiateVariableFont`, fixing `wght` at 400
and 700 respectively). Oblique and BoldOblique derive from those faces
by decomposing glyphs and applying a 12-degree rightward shear:
`(x, y) -> (x + tan(12 degrees) * y, y)`. Their italic angle is -12 degrees,
with the italic flags set in the `head` and `OS/2` tables. Font-family and
PostScript names were changed to Document Sans SC / DocumentSansSC.

The four static faces were losslessly compressed to WOFF with fontTools
(`font.flavor = "woff"; font.save(path)`). Glyph order, Unicode mappings and
horizontal metrics were checked against the original TTF faces. Resource
size decreased from approximately 41.1 MiB to 25.0 MiB; this does not imply
the same reduction in already ZIP-compressed wheel downloads. WOFF needs
fpdf2 2.8.9 or newer; the PDF embeds ordinary font subsets, not WOFF data.

All source glyphs were retained; the resources are not limited to the
characters used in the regression test. This provides common Simplified
and Traditional Chinese characters plus Latin text and punctuation.
It is not an all-language font or an emoji font.

The PDF writer substitutes this family for source fonts, including code
blocks. Bold uses the 700-weight face; italic uses the derived oblique
face. Original font family, monospace metrics and exact Word pagination
are not preserved. Only styles appearing in the document are registered;
fpdf2 subsets the fonts when embedding them in PDFs. Missing glyphs emit
a warning unless covered by explicitly configured trusted fallback fonts.
