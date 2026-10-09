# 增强版转换、诊断与安全边界

本页描述本 fork 的 `26.7.0.post2`，不是官方 Aspose 的功能保证。
新增内存输出、结构化内容、原包定点替换和塑形的用法见 [升级说明](upgrade-notes.md)；
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
直接 `ConversionOptions` 构造的五个共享枚举使用相同校验，详见[完整字段审计入口](save-option-audit.md)。

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
本轮使用其解析 token 验证语法，另用 Pandoc 3.8.3 核对 HTML，未改写独立工具来制造一致结果。
具体 Markdown 阅读器行为仍需调用方验证；本轮未进行浏览器或原生 Office 界面验收。

## PDF 图片压缩

`jpeg_quality` 接受 0..100 整数，在 PDF 输出前拒绝布尔、非整数及越界值。
AUTO 即使请求较低 JPEG 质量，也保留含透明像素图片的原字节，避免把透明区域错误变为实色；
完全不透明的 alpha 图片仍可使用 JPEG。显式 JPEG 请求会将透明图片合成到白底并发出
`pdf.image_transparency_lost` 内容损失诊断，严格模式保留原输出并拒绝转换。
白底合成不承诺在其他背景上保持外观，见[复验记录](benchmarks/pdf-image-compression-contract.json)。

## 中文排版和字体

- 高亮文本、混合字体运行的居中/右对齐段落按实际字形宽度换行，保留显式换行与链接。
- 优先保留可放进一行的空格分隔单词；过长单词、无空格中文按字符拆行。
- 纯文字段落高度在隔离状态下试运行实际渲染器，包含首行/续行缩进、段前/段后间距和 Code/Quote 的有效字号；普通混合字号段落行高容纳最大字号。
- 左右缩进持续作用于续行及跨页/跨栏后的文字；首行缩进只作用于首行，可为负值（悬挂缩进）。
  `keep_together` 在下一完整页/栏能容纳时整体移动，超出完整区域的长段落仍允许拆分。图文绕排高度仍是近似估算。
- 四个完整字体从 TTF 无损压缩为 WOFF：资源约 **41.1 MiB → 25.0 MiB**，减少约 **39%**。
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
`--text-shaping` 需从本仓库安装 `[shaping]`，安装命令见 [升级说明](upgrade-notes.md#5-可选多语言塑形)。
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
