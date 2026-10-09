# 增强版转换、诊断与安全边界

本页描述本 fork 的 `26.7.0.post2`，不是官方 Aspose 的功能保证。
内存输出、结构化内容和模板用法见 [项目 README](../README.zh-CN.md)，原包编辑见 [DOM 指南](docx-dom.md)；
同类项目的参考方向见 [生态对比](ecosystem-adoption.md)，固定商业基准、差距和复验材料见
[商业对齐计划](commercial-alignment-plan.md)。

## 保存选项与实际格式

`OoxmlSaveOptions.save_format` 只接受 DOCX，`MarkdownSaveOptions.save_format` 只接受 Markdown。
可使用 `SaveFormat`、对应整数或已有字符串别名（大小写不敏感，如 `docx`、`markdown`、`md`）。
不匹配或未知值在渲染前抛出 `ValueError`，文件输出保留原文件，内存输出不返回错误格式的字节。
`OoxmlSaveOptions(None)` 仍使用默认 DOCX；创建后显式改为 `None` 属于无效值。
扩展名不覆盖显式选项。此规则也适用于直接使用 `LdmDocxWriter`，以及创建 writer 后修改选项。
这是行为修正：此前不匹配的 `save_format` 被静默忽略；切换格式应传入对应的 `SaveFormat` 或选项类。

三个保存选项类的 15 个枚举字段接受已定义成员、对应整数和大小写不敏感的成员名；
未知字符串、越界整数及错误类型现在在赋值时抛出 `ValueError`，保留原属性值。
此前部分请求被静默忽略或在保存时才失败；调用方应在赋值与保存边界处理异常。
直接 `ConversionOptions` 构造的五个共享枚举使用相同校验，详见[逐字段审计](#保存选项逐字段审计)。

有效但未实现的请求不会因枚举校验而成为已实现功能：`NON_COMPATIBLE_TABLES` 仍等同 `NONE`，
`Zip64Mode.ALWAYS` 仍等同 `IF_NECESSARY`，Markdown 非默认 `image_resolution` 不会重采样图片。
这三种请求现在产生 `ConversionWarning` 及 `markdown.unsupported_option` / `docx.unsupported_option`
诊断；默认设置不增加警告。需要拒绝这些退回行为时，可将 `ConversionWarning` 升级为错误，
输出前失败会保留原主文件。`--strict` 针对内容损失，并非拒绝所有选项退回的通用开关。

### 布尔保存开关

`PdfSaveOptions` 的八个、`MarkdownSaveOptions` 的三个及 `OoxmlSaveOptions.pretty_format`
布尔字段仅接受 `True` / `False`。字符串（包括 `"false"` / `"true"`）、0/1、None 和其他
非布尔值在赋值时抛出 `ValueError`，不改变原属性，也不污染 PDF 显式请求诊断集合。
直接 `ConversionOptions` 的八个布尔字段在构造和赋值时同样拒绝错误类型。
从配置读取这些开关时，调用方应先完成有明确规则的解析，不要依赖 Python 的一般真值判断。

此前非空字符串 `"false"` 会启用 PDF 结构、Markdown 下划线或 DOCX 美化输出；
现在不会接受该赋值。默认值、公开实例字典字段、复制及有效 True/False 的实际行为保留。
尚未实现的 PDF 开关继续发出 unsupported_option 诊断，类型通过不意味着功能已支持。
普通 Markdown 的下划线使用 `++text++` 扩展语法，HTML 表格使用 u 标签；
普通 Markdown 扩展不是跨阅读器兼容保证。本次未改变该输出语法。
真实 DOCX、独立 pypdf/ZIP/python-docx 回读及前后记录见[布尔契约](benchmarks/boolean-save-contract.json)。

### PDF 打开缩放因子

`PdfSaveOptions.zoom_factor` 仅在 `zoom_behavior=ZOOM_FACTOR` 时消费，接受有限非负
`int` / `float`，拒绝 bool、字符串、None、负数和非有限值；保存时抛出 `ValueError`，
保留原目标文件。其他缩放模式不解释该字段。0 保留现有默认打开行为，正数表示百分比，
例如 125.5 写入 `/XYZ` 的缩放比例为 1.255。

正值在 Catalog 序列化前以普通十进制写入，避免底层科学计数法产生无法解析的 PDF。
当前 writer 将比例数字限制为最多 47 个字符，超出时提前拒绝：这是根据独立 pypdf
数值解析预算设置的兼容边界，不是 PDF 标准规定的统一数值上限。
依赖 fpdf2 私有 Catalog hook，升级依赖时须复验六种模式和大纲行为。
独立解析、正文保留和默认输出对照不代表已验证各阅读器的实际打开缩放效果；
复验入口见[缩放契约记录](benchmarks/pdf-zoom-contract.json)。

### PDF 大纲输入与补层上界

保存时统一校验 `OutlineOptions` 的六个字段：标题深度为非负整数（既有标题输出仍夹紧到六），
展开深度、默认书签深度和每个名称的书签深度为 0..9 整数；所有深度拒绝 bool。
书签映射要求字符串键的 Mapping，两个 `create_*` 开关只接受 bool。
错误在渲染和补层分配前抛出 `ValueError`，适用于内存、文件和直接 writer，保留原目标文件。
创建 writer 后改字段或原地修改映射，同样会在下一次保存重新校验。

0..9 书签深度与[商业版公开范围](https://reference.aspose.com/words/net/aspose.words.saving/outlineoptions/defaultbookmarksoutlinelevel/)
一致，但不是完整兼容承诺：显式映射 0 省略该书签；默认深度 0 且映射为空时，
仍按本项目 `export_bookmarks_outline` 退回到一级或省略。隐藏名称不因显式映射而输出。
此前超大书签深度可能在 `create_missing_outline_levels=True` 时循环创建大量空节点；
现在超界值提前拒绝，不把补层关闭或静默夹紧当作请求已生效。
真实 DOCX、独立 pypdf 导航和受控分配探针见[复验记录](benchmarks/pdf-outline-contract.json)。
实际阅读器导航操作及完整保存矩阵仍未验收。

### Markdown 页眉页脚损失

Markdown 当前只导出默认页眉/页脚中顶层段落的图片，省略文字、表格、注引用、横线和变体图片。
现在这些省略会产生 `ContentLossWarning` 和 `markdown.header_footer_content_omitted` 诊断；
隐藏文字、纯字段指令及空段落不触发该诊断。诊断不包含被省略的文字，也不修改源模型。
默认输出仍保持既有 Markdown 内容；需要保留页眉页脚应使用原包编辑或支持该内容的输出路径。
`--strict` 或将 `ContentLossWarning` 升级为错误的调用会在输出前失败，保留原目标文件。
这补齐了有损转换提示，并未实现 Markdown 页眉页脚渲染。

### HTML 表格与字段可见性

`export_as_html=TABLES` 使用实际 HTML 行/列跨度和格式标签，保留嵌套表格及段落顺序、
普通编码链接和图片；文字与属性转义，危险 URI 不激活并有损失诊断。
孤立的纵向合并续格保留内容并报警。默认仍为管道 Markdown，合并/嵌套几何损失继续报警。
HTML 表格中的复杂字段保留可见结果，但不导出字段动作/目标；注标签和定义可读，
语义注锚点尚未支持；非图片形状也未支持。这些限制都有 `ContentLossWarning`，不能视为功能完成。
HTML 输出不等同 Word 的字体、页面布局和完整交互语义。

共享可见性过滤现在按每层字段的指令/结果状态处理嵌套字段：内部字段结束不会让外层指令
文字进入 Markdown/PDF 输出。指令区的注引用不会触发未显示注正文的导出；可见嵌套结果和
用内部字段结果组成的 Markdown 超链接目标仍保留。原包和结构化源内容仍可保留原始字段数据。

Markdown 的正文、表格和默认页眉页脚图片，以及 PDF 的内联、浮动、锚定、定位图片和文本框，
现在均排除字段指令区节点；含图段落的文字也使用同一可见性边界。PDF 图片测量、预绘制和
纯分页识别同步过滤，避免不绘制的指令图片仍改变布局。可见结果图片继续输出，外置 Markdown
图片目录不写入指令图片；源模型和原包内容不会因此被删除。见[复验记录](benchmarks/field-shape-visibility.json)。

### 外置图片文件名

外置 Markdown/HTML 表格图片保留可用的原文件名；同名但字节不同的图片追加序号，
相同字节复用，已有无关文件和目标符号链接不会被覆盖。主 Markdown 路径不会被作为
图片目标；新图片使用原子文件写入。目录已有冲突时生成名称可能变化，链接随名称更新。
渲染、编码、后续图片写入及主文件发布的普通异常会回滚本次新图片，已有及复用文件保留。
独占占用标记保持到主文件提交结束，其他导出避开仍可能回滚的图片；本次重复字节仍复用。
单导出失败、受控并发占用及失败/成功交错已验证，见[回滚复验](benchmarks/markdown-image-rollback.json)
和[同名图片复验](benchmarks/markdown-image-collisions.json)。这不是断电或强制终止安全的多文件
原子事务；中断可能留下图片或占用标记，新建空目录也可能保留。
本机文件系统上的独立进程占用、普通异常回滚和强制终止后再次导出已验证，
见[进程复验](benchmarks/markdown-image-processes.json)。被终止进程的残留保留，后续导出避开
占用名称；已有占用标记不会被自动删除。自动恢复、断电及网络文件系统行为仍未验收。

### 图片替代文本与 URI

普通 Markdown 图片的替代文本按字面转义，避免 `]`、HTML 标签和嵌套图片片段变成新语法；
CR/LF/TAB 使用字符引用，保留文本并避免空行拆开图片。正文、管道表格和默认页眉页脚图片
复用该路径。缺失格式元数据的内联栅格图片使用现有 Pillow 识别 MIME，不重编码图片或修改模型；
无法识别并且没有有效类型的内联图片会产生内容损失诊断。

外置图片的物理文件名保持不变，链接使用 `/`，空格、`#`、`?`、百分号及其他路径字符做 URI 编码。
`images_folder_alias` 是 URL 前缀：在其 path 后添加编码文件名，保留 query/fragment；
别名中需要作为路径字符的 `#` / `?` 应写成 `%23` / `%3F`，已有百分号编码不会重复编码。
激活的别名必须是没有原始 ASCII 控制字符的字符串；URL 解析报错时在创建副产物前拒绝。
直接 writer 没有主输出路径而使用绝对图片目录时，生成 `file:` URI。

普通 Markdown 与 HTML 表格图片均拒绝危险 scheme，并产生 `ContentLossWarning`；
严格模式沿用已有回滚，保留原主文件和已有图片。这里只检查 URI/输出语法，不访问远程资源，
也不提供外部文件访问授权或完整 HTML/SVG 净化。

[复验记录](benchmarks/markdown-image-links.json)包含真实 DOCX、独立 Markdown token 和 Pandoc HTML 对照。
`markdown-it-py 4.2.0` 的图片 alt 渲染会漏掉 `text_special` token 中的转义标点；
该复验使用解析 token 检查语法，并用 Pandoc 3.8.3 核对 HTML。
具体 Markdown 阅读器行为仍需调用方验证；记录不包含浏览器或原生 Office 界面验收。

## PDF 图片压缩

`jpeg_quality` 接受 0..100 整数，在 PDF 输出前拒绝布尔、非整数及越界值。
AUTO 即使请求较低 JPEG 质量，也保留含透明像素图片的原字节，避免把透明区域错误变为实色；
完全不透明的 alpha 图片仍可使用 JPEG。显式 JPEG 请求会将透明图片合成到白底并发出
`pdf.image_transparency_lost` 内容损失诊断，严格模式保留原输出并拒绝转换。
白底合成不承诺在其他背景上保持外观，见[复验记录](benchmarks/pdf-image-compression-contract.json)。

## PDF 排版契约

表格保留逐 Run 格式、图片及段落/子表格的交错顺序；单元格 `children` 有序视图与
`content_order` 序列化兼容旧 `paragraphs` / `tables` 构造。换页重复连续前导表头，
可容纳的行整体移动，超高行按内容行拆分；不可满足的行高、cantSplit 或单元格分页要求会报警。
过高的不可拆嵌套表格、无法容纳的表头或旋转内容明确失败。水平跨度保留，垂直合并仍为
网格/边线近似；Markdown 展平嵌套表格并报告损失。它们不保证 Word 的完整表格排版。

源上下标标志已用于正文、表格、页眉页脚、列表正文、标题、引用和代码块的富文本绘制，包含塑形路径。
缩放、基线偏移、装饰线及链接区域采用 fpdf2 度量；自动行高预留上标空间，固定行高保留源值。
代码块保留逐 run 字号、颜色、粗斜体、装饰和上下标，背景在换页/换栏后逐行绘制。
混合上下标旋转单元格仍会报告转换损失。度量不承诺与 Word 一致，见
[上下标验收](benchmarks/vertical-positions.json)与[代码块验收](benchmarks/code-runs.json)。
富文本正文和代码块的两端对齐会伸展自动换行行中的空格，链接与装饰同步伸展；
段末保持自然宽度。手动换行按文档级 `do_not_expand_shift_return`（默认 `False`）处理：
DOCX 读取并写出 `w:compat/w:doNotExpandShiftReturn`；设为 `True` 后手动换行不伸展。
普通及嵌套表格单元格复用相同两端对齐规则，列表编号保持自然宽度；
普通单元格段落应用左右及首行缩进，图片按可用宽度缩放。禁用自动换行仍保留显式换行，
超宽行会被固定单元格裁剪。负缩进限制在单元格内边界，旋转/文本框混排及 CJK 字符分布仍有限制。
见[两端对齐验收](benchmarks/justification.json)、[手动换行验收](benchmarks/shift-return.json)
及[表格对齐验收](benchmarks/cell-justification.json)。
LDM 的 `compatibility_mode` 默认 15；DOCX 按官方 URI 读取并写出兼容模式，缺失/格式无效时按 12。
旧模式小于 15 时，左对齐流式表格按首单元格边距外移；居中/右对齐忽略表格缩进。
这会改变此前忽略旧模式的 PDF 横坐标；设为 15 可使用现代位置。该字段不代表完整版本兼容，
浮动/嵌套定位和边框度量仍有限制，见[表格定位验收](benchmarks/table-compatibility.json)。

表格、单元格与表格样式的四侧边距以 pt 为单位：`None` 表示未设置/继承，`0` 表示显式零边距。
DOCX 读写保留零值，PDF 测量与绘制采用单元格覆盖、表格默认、渲染器回退的相同顺序；
DOCX 表格默认边距包括已支持的样式继承。表格上下边距是单元格内部间距，不再额外增加表外空白。
旧 LDM JSON 若用 `0` 表示默认间距，应改为 `null`；显式零值现在会改变布局。
条件表格样式及行级边距例外仍未实现，见[边距验收](benchmarks/cell-margins.json)。
`Table.bidi` 保留已支持的 DOCX RTL 表格方向，镜像视觉列、对齐及边框，同时保留逻辑内容顺序。
`start/end` 边距归一为 left/right 字段，DOCX 输出 left/right。普通 Markdown 报告方向丢失，
HTML 表格使用 `dir="rtl"`；表格方向不会自动启用文字塑形。
Word 与 LibreOffice 对照及边界见[方向验收](benchmarks/table-direction.json)。
小字号及短旋转单元格的最小行高已修正单位；受影响的行会变矮，页/栏断点可能改变，
见[行高验收](benchmarks/row-height-units.json)。
打包字体预存字形名称以减少初始化工作，保留全部字形与度量；wheel 增大约 0.86%，
实测收益依文档而异，见[字体基准](benchmarks/font-glyph-names.json)。

## 中文排版和字体

- 高亮文本、混合字体运行的居中/右对齐段落按实际字形宽度换行，保留显式换行与链接。
- 优先保留可放进一行的空格分隔单词；过长单词、无空格中文按字符拆行。
- 纯文字段落高度在隔离状态下试运行实际渲染器，包含首行/续行缩进、段前/段后间距和 Code/Quote 的有效字号；普通混合字号段落行高容纳最大字号。
- 左右缩进持续作用于续行及跨页/跨栏后的文字；首行缩进只作用于首行，可为负值（悬挂缩进）。
  `keep_together` 在下一完整页/栏能容纳时整体移动，超出完整区域的长段落仍允许拆分。图文绕排高度仍是近似估算。
- 四个完整字体从 TTF 无损压缩为 WOFF：当前四个 WOFF 资源合计约 **24.6 MiB**。
  字形映射、数量和宽度信息保留；PDF 内仍嵌入标准字体子集。只注册文档出现的字体样式。
  这是安装体积优化，不承诺下载体积同等下降（wheel 原本已 ZIP 压缩）。
- 复杂文字塑形可通过可选 `[shaping]` 依赖和 `PdfSaveOptions.text_shaping=True` 开启，复用 fpdf2/HarfBuzz；还需部署适合该语言的可信 fallback 字体。尚未实现语言标点禁则、源字体精确匹配或 Word 的完整分页规则。

### 缺字、降级和 fallback

```python
import warnings
import aspose.words_foss as aw
from aspose.words_foss.pdf_writer import PdfMissingGlyphWarning

warnings.simplefilter("error", PdfMissingGlyphWarning)  # 不接受缺字 PDF
options = aw.saving.PdfSaveOptions()
options.fallback_fonts = ["/trusted/fonts/extra-unicode.ttf"]
aw.Document("report.docx").save("report.pdf", options)
```

字体文件由应用部署和信任；库不会联网下载字体。fallback 按列表顺序尝试，使用 regular face，
可能无法保留相应字符的粗体/斜体；也不承诺彩色 emoji 支持。

诊断类别：

| 类别 | 触发情形 |
|---|---|
| `PdfMissingGlyphWarning` | 主字体和 fallback 都缺字；提示缺失码点 |
| `PdfFontSubstitutionWarning` | 源字体被 Document Sans SC 替代，可能改变布局 |
| `PdfUnsupportedOptionWarning` | 显式设置尚未实现的 PDF 选项，或请求 PDF/A、PDF/UA 合规 |
| `PdfContentLossWarning` | 未知模型节点、页眉页脚变体、超高不可拆行等已知内容或布局降级；未启用塑形的复杂文字也会提示 |
| `aw.loading.DocumentLoadWarning` | 脚注/尾注、评论、对象、altChunk、修订、简单字段、内容控件、公式以及多种页眉页脚不能完整保留 |
| `DocxWriterLossyWarning` | 未知模型节点、不同章节或不支持类型的页眉页脚不能完整写回 |
| `aw.ContentLossWarning` | 已知内容/布局损失的公共基类；用于统一过滤加载、DOCX、Markdown、PDF 的已知损失 |

默认值不会仅因选项尚未实现而报警；显式赋值即使等于默认值也会提示。
这些诊断只覆盖已检测到的情况，不是全面的 Word 兼容性检查。可用 Python `warnings` 分类过滤或升级为错误。
公开 `Document.diagnostics` 列表保存加载、`save()` 和 `to_bytes()` 的结构化记录，包含 `code`、`severity`、`location`、`message`；即使 Python warning 被忽略或升级为错误，已有诊断仍会记录。定位信息是已知部件/模型位置，不是假定页码。
记录不会自动清空，重复转换可能追加重复项；按次统计可在操作前调用 `doc.diagnostics.clear()`。
构造失败时处理抛出的异常，不能读取尚未返回的 Document。诊断上下文隔离不代表同一可变文档或 writer 可并发使用。

## 参考 DOCX 样式

设置 `OoxmlSaveOptions.reference_docx = "brand.docx"`，生成 DOCX 时应用参考文件中受支持的样式定义。
run 字体读取时先应用字符样式继承，再应用直接格式；写出时省略匹配的继承值，使参考字符样式生效，
同样覆盖链接和 PAGE 域。LDM 尚不能区分恰好等于原继承值的显式直接格式，换参考样式后该值可能随样式改变；
完整 Word toggle 语义仍未保证。
未显式指定段落样式的 DOCX 段落现在使用 XML 标记的默认段落样式，而非按 Normal 名称猜测；
正文、表格和页眉页脚共用此解析规则。LDM `Style.is_default` 保留默认标记，旧模型缺少字段仍兼容。
按样式显示名匹配，并重映射 ID 和 basedOn 引用；正文样式及超链接关系保持有效。已有直接格式优先。
不导入参考正文、图片、页眉页脚和页面设置。样式经过现有 LDM reader/writer，未支持的 OOXML 属性
不会原样保留；这不是原包模板复制或无损样式导入。

## 输入和输出边界

`Document` 的文件、字节和流入口统一有界读取：

| 限制 | 默认上限 |
|---|---:|
| 输入 | 64 MiB |
| DOCX ZIP 条目 | 10,000 |
| 单 ZIP 部件/图片数据 | 64 MiB |
| DOCX 总展开大小 | 256 MiB |
| 单张光栅图片 | 25,000,000 像素 |
| 单个 DOCX gridSpan / 输出表格网格列数 | 1,024 |

DOCX 禁止重复/加密 ZIP 条目、不安全条目路径和 XML 实体扩展。SVG 不允许外部图片引用，
只接受内部引用及内嵌光栅图片。超限会抛出错误；大文件业务需先评估资源需求再调整源码中的上限。

DOCX/Flat OPC 的字符单位缩进尚未映射到转换模型。加载时汇总 `leftChars/rightChars`、
`startChars/endChars`、`firstLineChars/hangingChars`，发出 `load.character_indents_ignored`
损失诊断；扫描正文、样式、编号和附属 story，包含显式零值，诊断不含正文或属性值。
`ContentLossWarning` 错误过滤及 CLI `--strict` 在输出前拒绝；原包 DOM 保存仍保留这些属性。
见[检测与拒绝复验](../tests/test_character_indent_diagnostics.py)。这是损失检测，不是字符缩进支持。

Markdown 默认不读取本地图片，仍支持内嵌 base64 图片。可信文件需要本地图片时显式开启：

```python
options = aw.MarkdownLoadOptions()
options.allow_local_images = True
aw.Document("documents/report.md", options).save("report.pdf")
```

开启后图片必须位于文档目录内，绝对路径、`..` 或符号链接不能逃出该目录；不会自动下载 HTTP 图片。
没有文档目录的字节/流输入不能加载本地图片。其他低层读写器不是完整的服务安全入口。

`Document.save()` 的 PDF、DOCX、TXT 和 Markdown 主文件在同目录暂存并原子替换；转换、写入或替换失败
不会覆盖原主文件，已有输出符号链接、非普通文件及不可写文件会被拒绝。新文件默认权限为 `0600`，替换已有文件时保留其权限位。
Markdown 导出的独立图片不是多文件事务；原子替换也不等于完整断电恢复保证。

## 有界单任务 CLI

```bash
python -m aspose.words_foss.convert report.docx report.pdf \
  --strict --timeout 60 --memory-mb 1024
```

每次调用只有一个隔离 worker；默认墙钟超时 60 秒、RSS 阈值 1024 MiB、单输出文件上限 256 MiB，
退出/超时/超内存时清理进程组与暂存目录。Linux 还设置 `RLIMIT_AS` 硬地址空间上限；macOS 使用
RSS watchdog，约每 0.1 秒检查，**可能短暂超过阈值，不是硬内存配额**。macOS 需要系统 `ps`。
`--memory-mb 0` 显式关闭内存检查。该 CLI 目前只支持 POSIX；Windows 仍可使用普通库 API。

`--strict` 将缺字及 `ContentLossWarning`（已知加载/写出损失）升级为错误，但不因正常字体替换而失败。它不是完整保真验证器。
`--fallback-font /trusted/font.ttf` 可以重复指定；它与 `--text-shaping` 都只适用于内置后端的 `.pdf` 输出，其他格式会拒绝这些参数。
`--text-shaping` 需安装 `[shaping]` 并部署可信字体：

```bash
python -m pip install "aspose-words-foss-enhanced[shaping] @ git+https://github.com/shenyankm/DocWeave.git@dev"
```

实际绘制和行宽/行高测量同时启用 shaping。实现使用 fpdf2 内部 bidi/断行 API，升级 fpdf2 后须复验。
普通 `Document.save()` 没有进程超时/内存限制；服务端请使用此入口或自己的任务隔离层。
CLI 不是 Web 服务、队列或全局并发控制器；应用需限制同时启动的作业数，并配置容器内存/磁盘/网络配额。

## 可选 LibreOffice 原文件路径

```python
from aspose.words_foss.libreoffice import convert_to_pdf

convert_to_pdf("report.docx", "report.pdf", timeout=60)
```

或者通过同一个受限 CLI：

```bash
python -m aspose.words_foss.convert report.docx report.pdf --backend libreoffice
```

需自行安装 LibreOffice，`soffice`/`libreoffice` 在 PATH 上；也识别 macOS 标准应用安装位置。
此路径传入**原始 DOC/DOCX/RTF 字节**，不经过可能丢信息的 LDM 重建。每个任务使用私有用户配置目录，
验证进程状态及 PDF 头/结束标记后才发布文件。内置 RTF reader 仍只支持 OLE2/DOC-backed 文件；现在会对标准文本 RTF 明确报错并提示此入口。

这是独立可选入口，不改变 `Document.save()` 默认后端。高保真仍受 LibreOffice 兼容性和部署的
**系统字体**影响；它不直接使用本库 WOFF 字体。`--strict`、`--fallback-font` 和 `--text-shaping` 不适用于该后端。
本仓库既测试协议/异常，也提供真实渲染测试。已在 macOS / LibreOffice 26.2.6 验证中文 DOCX、标准文本 RTF、
原始 DOCX 脚注和受限 CLI；脚注可提取到 `source_stories`，但轻量模型的转换输出仍会丢失，
原文件转换中可保留。另有新增 DOM 图片、超链接与水平合并的真实 LibreOffice 渲染用例。未安装 LibreOffice 的环境跳过真实测试。
这些用例证明入口可用和已测内容保留，不证明所有文档与 Microsoft Word 视觉一致。

**私有 profile 与资源限制不是安全沙箱。** 原生 Office 解析器、宏、外部链接与文件系统访问必须由
应用/容器沙箱约束；对不可信输入关闭网络、隔离可读写目录、以非特权账户运行并保持 LibreOffice 更新。
尤其不要将普通 `convert_to_pdf()` 当作具备硬资源/外部资源隔离的服务器入口。

## 保存选项逐字段审计

范围为 `saving.py` 的三个保存选项类及其 `OutlineOptions`：40 + 6 个初始化公开字段，其中 15 个字段使用枚举描述符。
下表记录源码行为与复验入口，不表示已完整验收每项；全矩阵效果、标量边界和跨格式组合仍需完整选项验收。
枚举拒绝与三项有效但未实现请求的诊断见[记录](benchmarks/save-option-contract.json)；
后续标量及内容专项按各行记录核对，不能升级为全矩阵已通过。
支持枚举值不等于对应功能已实现：有效但未实现值仍按表中限制处理。

### PdfSaveOptions

| 字段 | 当前行为/限制 | 证据或复验入口 |
|---|---|---|
| `compliance` | 只设置 PDF 版本；PDF/A、PDF/UA 请求警告，非标准认证 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `export_document_structure` | 结构输出开关；完整辅助技术/标准验收仍待决策 D02 | [test_pdf_structure_pages.py](../tests/test_pdf_structure_pages.py) |
| `image_compression` | AUTO 保留有意义的透明像素；显式 JPEG 合成白底并诊断透明度损失，严格模式可拒绝 | [test_pdf_image_compression_contract.py](../tests/test_pdf_image_compression_contract.py) |
| `jpeg_quality` | 0..100 整数，布尔/非整数/越界拒绝；作用于 JPEG 编码，AUTO 的透明图片保留原字节 | [test_pdf_image_compression_contract.py](../tests/test_pdf_image_compression_contract.py)、[复验](benchmarks/pdf-image-compression-contract.json) |
| `text_compression` | NONE/FLATE 页面内容流，字体/图片流独立 | [test_pdf_stream_options.py](../tests/test_pdf_stream_options.py) |
| `embed_full_fonts` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `use_core_fonts` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `font_embedding_mode` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `page_mode` | 写入 PDF PageMode；阅读器可以忽略 | [test_pdf_stream_options.py](../tests/test_pdf_stream_options.py) |
| `color_mode` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `export_bookmarks_outline` | 书签输出；显式 outline 层级优先 | [test_pdf_outline_options.py](../tests/test_pdf_outline_options.py) |
| `outline_options` | 下表六项；不是完整 Word TOC 契约 | [test_pdf_outline_options.py](../tests/test_pdf_outline_options.py) |
| `preserve_form_fields` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `memory_optimization` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `zoom_factor` | ZOOM_FACTOR 下使用：有限非负 int/float（拒绝 bool），0 保留默认打开行为；正值按百分比写成普通十进制，PDF 数字最长 47 字符。内存/文件/直接 writer、错误原文件保留及独立解析已验证；实际阅读器打开行为未验收 | [缩放契约](benchmarks/pdf-zoom-contract.json)、[writer](../aspose/words_foss/pdf_writer/writer.py) |
| `zoom_behavior` | PDF 打开动作；部分行为在序列化后改写 | [aspose/words_foss/pdf_writer/writer.py](../aspose/words_foss/pdf_writer/writer.py) |
| `display_doc_title` | PDF viewer preference；阅读器可忽略 | [aspose/words_foss/pdf_writer/writer.py](../aspose/words_foss/pdf_writer/writer.py) |
| `fallback_fonts` | 可信字体路径；覆盖、塑形及源字体替代有独立限制 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `text_shaping` | 可选 HarfBuzz 塑形；依赖/字体及格式边界有限制 | [test_pdf_shaping.py](../tests/test_pdf_shaping.py) |

### OoxmlSaveOptions

| 字段 | 当前行为/限制 | 证据或复验入口 |
|---|---|---|
| `save_format` | 只支持 DOCX；其他值在输出前拒绝 | [test_save_option_formats.py](../tests/test_save_option_formats.py) |
| `compliance` | Transitional/ECMA 路径；STRICT 明确拒绝 | [test_bounded_conversion.py](../tests/test_bounded_conversion.py) |
| `compression_level` | 四档 ZIP deflate；不是内容精简 | [ApiExamples/working_with_ooxml_save_options.py](../ApiExamples/working_with_ooxml_save_options.py) |
| `zip_64_mode` | NEVER/IF_NECESSARY；ALWAYS 仍等同按需，但现在警告 | [test_unsupported_save_requests.py](../tests/test_unsupported_save_requests.py) |
| `pretty_format` | XML 缩进；语义与全部部件仍待矩阵复核 | [ApiExamples/working_with_ooxml_save_options.py](../ApiExamples/working_with_ooxml_save_options.py) |
| `reference_docx` | 导入参考样式，非参考正文/资源/页面设置 | [test_reference_styles.py](../tests/test_reference_styles.py) |

### MarkdownSaveOptions

| 字段 | 当前行为/限制 | 证据或复验入口 |
|---|---|---|
| `table_content_alignment` | 表格对齐覆盖 | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py) |
| `list_export_mode` | Markdown 列表或纯文本路径 | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py) |
| `export_images_as_base64` | 无 images_folder 时始终内联；有目录时选择内联/外部 | [test_conversion_api.py](../tests/test_conversion_api.py) |
| `images_folder` | 同名不同内容分配序号，已有文件及目标符号链接保留；普通失败回滚新图片，占用标记保护待提交图片；非断电安全多文件事务；链接使用编码 URI 路径 | [test_markdown_image_rollback.py](../tests/test_markdown_image_rollback.py)、[test_markdown_image_collisions.py](../tests/test_markdown_image_collisions.py) |
| `images_folder_alias` | URL 前缀：编码文件名追加到 path，保留 query/fragment；激活时拒绝非字符串/原始 ASCII 控制字符及非法 URL | [图片 URI 契约](benchmarks/markdown-image-links.json)、[独立解析](../tests/test_markdown_image_links.py) |
| `export_underline_formatting` | 普通 Markdown 使用 `++` 扩展标记，HTML 表格使用 u 标签；不是所有 Markdown 阅读器都支持该扩展 | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py)、[实际输出](../tests/test_boolean_save_contract.py) |
| `link_export_mode` | 自动/内联/引用链接 | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py) |
| `export_as_html` | TABLES 走已有 HTML 路径；NON_COMPATIBLE_TABLES 仍等同 NONE，但现在警告 | [test_unsupported_save_requests.py](../tests/test_unsupported_save_requests.py) |
| `empty_paragraph_export_mode` | 空行、HTML br 或省略 | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py) |
| `image_resolution` | 未重采样；非默认 96 请求现在警告，源图片字节保留 | [test_unsupported_save_requests.py](../tests/test_unsupported_save_requests.py) |
| `save_format` | 只支持 MARKDOWN；其他值在输出前拒绝 | [test_save_option_formats.py](../tests/test_save_option_formats.py) |
| `encoding` | 文本编码；编码失败在主输出写入前发生 | [test_conversion_api.py](../tests/test_conversion_api.py) |
| `paragraph_break` | 顶层块分隔；内部换行仍为 LF | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py) |
| `style_map` | 精确样式名映射为支持的语义样式 | [test_ecosystem_semantics.py](../tests/test_ecosystem_semantics.py) |
| `export_notes` | 可见注锚点、定义与回链；缺失/不支持位置有诊断 | [test_markdown_notes.py](../tests/test_markdown_notes.py) |

### OutlineOptions

| 字段 | 当前行为/限制 | 证据或复验入口 |
|---|---|---|
| `headings_outline_levels` | 非负整数，拒绝 bool；标题深度大于六的值按现有规则夹紧 | [大纲契约](benchmarks/pdf-outline-contract.json)、[效果](../tests/test_pdf_outline_options.py) |
| `expanded_outline_levels` | 大纲展开范围 0..9 整数，拒绝 bool | [大纲契约](benchmarks/pdf-outline-contract.json)、[效果](../tests/test_pdf_outline_options.py) |
| `default_bookmarks_outline_level` | 0..9 整数，拒绝 bool；0 且空映射保留 export_bookmarks_outline 退回规则，非商业版完全相同契约 | [大纲契约](benchmarks/pdf-outline-contract.json)、[效果](../tests/test_pdf_outline_options.py) |
| `bookmarks_outline_levels` | 字符串名称到 0..9 整数的 Mapping，显式 0 省略该书签，隐藏名称仍跳过；每次保存重新校验 | [大纲契约](benchmarks/pdf-outline-contract.json)、[效果](../tests/test_pdf_outline_options.py) |
| `create_outlines_for_headings_in_tables` | bool；表格标题输出导航 | [大纲契约](benchmarks/pdf-outline-contract.json)、[效果](../tests/test_pdf_outline_options.py) |
| `create_missing_outline_levels` | bool；缺层导航补层，先验证请求层级防止无界分配 | [大纲契约](benchmarks/pdf-outline-contract.json)、[效果](../tests/test_pdf_outline_options.py) |

### 选项验收边界

逐字段契约不等于全矩阵验收：标量边界、显式默认值和跨选项组合仍须检查实际输出。
`models.ConversionOptions` 是直接 writer 的额外接口；共享的五个枚举和八个布尔字段
遵循上文校验规则，其余字段仍需单独核对，不能计入上述 46 项的完成结论。
ZIP64 强制输出、选择性 HTML 表格及图片重采样仍未实现；诊断说明退回路径，不计为功能完成。
