# DocKiln 转换、诊断与安全边界

本页描述 `dockiln` 0.1.0 的源码接口。迁移后的包名和导入名均为 `dockiln`，目录为
`src/dockiln/`；不再提供 `aspose.words_foss` 入口。运行库不依赖商业 Aspose 包。
上游来源见 [项目 README](../README.zh-CN.md)，原始版权声明保留于 [MIT 许可证](../LICENSE)。

内置读取 DOCX、旧版 DOC、Flat OPC 及 OLE2/DOC-backed RTF，写出 DOCX、Flat OPC 或 PDF。
Markdown/TXT 读写、`MarkdownLoadOptions`、`MarkdownSaveOptions` 和 `ConversionOptions`
已移除；纯文本提取使用 `Document.get_text()`，结构化输出使用 `to_dict()`。
输入按内容检测格式，包括无扩展名文件和 Flat OPC XML；也可用 `LoadOptions.load_format`
显式指定格式。Flat OPC XML 输出应显式指定对应 `SaveFormat`。
本页保留的专项记录和测试数量属于迁移前证据，不代表当前包或远端 CI 已重新验收。
内存输出、结构化内容和模板用法见 [项目 README](../README.zh-CN.md)，原包编辑见 [DOM 指南](docx-dom.md)；
同类项目的参考方向见 [生态对比](ecosystem-adoption.md)，固定商业基准、差距和复验材料见
[商业对齐计划](commercial-alignment-plan.md)。

## DOCX 设置部件的加载

转换模型通过文档的 settings 关系读取设置，支持包内自定义路径；仅名为
`word/settings.xml`、没有该关系的部件不改变默认制表位、兼容模式或字体设置。
源包编辑需要保留未关联部件时，使用 `DocxDocument`。
标为 `External` 但指向包内现存 settings 部件的关系按已观测的基准恢复行为读取；
真正的外部 URL 不读取。关联或孤立设置的损坏 XML 和 DTD 仍明确报错。
有效设置映射到 `Document.light_document_model`。`Document.default_tab_stop` 以点读取和设置默认制表位，按 twip 取整；负数报错且保持原值，非有限输入遵循已观测的基准哨兵值。此接口不表示完整商业设置 API。

## 保存选项与实际格式

DOCX／Flat OPC 转换中的下划线声明区分缺失属性和显式空值：
`<w:u/>` 保留样式继承，`w:val=""` 或 `w:val="none"` 取消下划线。
写出覆盖样式的无下划线格式时使用显式 `none`，避免保存后重新继承下划线。
该规则不表示完整 Font DOM 接口或 PDF 下划线绘制已对齐。

`OoxmlSaveOptions.save_format` 支持 DOCX 和四种 Flat OPC 变体，`PdfSaveOptions` 用于 PDF。
可使用 `SaveFormat`、对应整数或已有字符串别名（大小写不敏感，如 `docx`、`flat_opc`、`pdf`）。
不匹配或未知值在渲染前抛出 `ValueError`，文件输出保留原文件，内存输出不返回错误格式的字节。
`OoxmlSaveOptions(None)` 仍使用默认 DOCX；创建后显式改为 `None` 属于无效值。
扩展名不覆盖显式选项。此规则也适用于直接使用 `LdmDocxWriter`，以及创建 writer 后修改选项。
这是行为修正：此前不匹配的 `save_format` 被静默忽略；切换格式应传入对应的 `SaveFormat` 或选项类。

两个保存选项类的枚举字段接受已定义成员、对应整数和大小写不敏感的成员名；
未知字符串、越界整数及错误类型现在在赋值时抛出 `ValueError`，保留原属性值。
此前部分请求被静默忽略或在保存时才失败；调用方应在赋值与保存边界处理异常。
当前保存选项详见[逐字段审计](#保存选项逐字段审计)。

有效但未实现的请求不会因枚举校验而成为已实现功能：`Zip64Mode.ALWAYS` 仍等同
`IF_NECESSARY`，并产生 `ConversionWarning` 和 `docx.unsupported_option` 诊断。
需要拒绝退回行为时，可将 `ConversionWarning` 升级为错误；输出前失败会保留原主文件。
`--strict` 针对内容损失，并非拒绝所有选项退回的通用开关。

### 布尔保存开关

`PdfSaveOptions` 的布尔字段和 `OoxmlSaveOptions.pretty_format` 仅接受 `True` / `False`。
字符串（包括 `"false"` / `"true"`）、0/1、None 和其他非布尔值在赋值时抛出 `ValueError`，
不改变原属性，也不污染 PDF 显式请求诊断集合。
从配置读取开关时，调用方应按明确规则解析，不要依赖 Python 的一般真值判断。
尚未实现的 PDF 开关继续发出 unsupported_option 诊断，类型通过不意味着功能已支持。
历史独立 pypdf/ZIP/python-docx 回读记录见[布尔契约](benchmarks/boolean-save-contract.json)。

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

### RGB 与自动颜色的声明来源

`Font.color_explicit` 记录当前层是否声明 `w:color`：`True` 为显式声明，
`False` 为继承，`None` 为旧模型或新建模型未记录来源。
显式 `auto` 使用 `Color [Empty]`，可以覆盖上层 RGB；它不等于没有颜色声明。
赋值或 `model_copy(update={'color': ...})` 会记录直接声明，即使颜色值未改变。
JSON 保留来源，DOCX／Flat OPC 写出保留受测 RGB 和自动色声明。
旧 JSON 不含此标记时沿用原有生成规则，无法恢复已经丢失的声明来源。

[固定的 256 个组合及 1,024 个保存结果](benchmarks/font-color-categories-26.9.json)
覆盖默认、段落样式、字符样式及 Run 的 RGB／自动色继承。
主题色、背景对比、DOM 颜色编辑和完整颜色 API 仍未验收。

主题色修复进行中：LDM 的 `Font.color` 改为保留源文件的 RGB／自动色 getter，
`color_rendering` 保存主题解析后的绘制值，`render_color` 提供实际绘制颜色。
此前读取主题色后使用 `font.color` 取得绘制 RGB 的调用方应改用 `font.render_color`。
显式颜色赋值和 `model_copy(update={'color': ...})` 会清除旧的主题渲染值；模型 JSON 保留该状态。
这是转换模型的区分方式，不表示规范 Font API 或 DOM 颜色编辑已实现。
新增的不可变 `SourceColor`／`SourceTheme` 只保存来源快照，不是主题编辑 API。
主题 XML、颜色属性及可达关联资源现在可经 JSON、DOCX 和 Flat OPC 保存；直接修改颜色会
丢弃旧的颜色声明快照。旧 JSON 没有这些快照，不能恢复此前展开为 RGB 时丢失的主题引用。
来源快照包含原部件名、内容类型及可达资源；内部关系在独立资源目录中重映射，保留共享资源和循环引用。
外部关系仅保留链接，不访问网络；缺失资源、重复关系、越界路径和悬空引用在输出前拒绝。
[专项用例](../tests/test_theme_resource_snapshot.py)检查加载、JSON、DOCX／Flat OPC 往返、资源重名和原输出保护。
[官方冷读证据](benchmarks/font-theme-resources-26.9.json)保留输入、两种当前输出及官方再次保存结果；
受测主题图片在六个文件中的可达资源 SHA256 一致，尚未验证主题图片的页面绘制。
参考样式继续按原行为使用解析后的 RGB，避免目标主题改变引用样式的颜色，且不修改参考模型。
组合 tint/shade 在受测上下文按固定基准使用原始回退色；目标框像素复验的通道容差为 1，
不表示全页面或背景对比已验收。完整主题资源兼容、规范主题编辑 API 及平台验收尚未闭环。
此前 pytest 默认导入模式会把仓库根目录加回搜索路径；原称安装包的专项结果已改记为源码验证。
当前使用 `--import-mode=importlib`，在 pytest 同一进程内确认包路径，验证最终 wheel 的 Mac 全量及 Linux 集成。
资源名中空格、字面百分号和中文均按 OPC URI 编码保留，最终 wheel 资源专项 29 项通过。
Linux 安装包集成 5,505 项通过；Mac 安装包全量 22,197 项通过，2 项跳过。
本次来源／资源保存修复已通过本地验收，交付开发分支；远端跨平台 CI 另行复验。

### Boolean 字体直接格式来源

读取 DOCX 时，LDM 的 13 个 Boolean 字体字段各自保留 `*_explicit` 来源：
`True` 表示直接声明，`False` 表示继承，旧模型的 `None` 表示来源未知。
显式赋值（包括赋回原值）或 `model_copy(update=...)` 会记录直接声明；模型 JSON 保留该状态。
DOCX／Flat OPC 写出保留显式开关，即使其值与继承值相同，也不会把它省略。
[其余十个属性的 3,240 个往返输出](benchmarks/font-boolean-roundtrip-26.9.json)
经官方冷回读与独立 XML 核对，200 个 getter 差异降为零；直接声明和样式引用保留。
复验归档通过记录中的 `storage` 引用复用逐字节相同的输出；`output` 仍是逻辑样本名，
每个样本的原始文件哈希和官方观测保持不变。两个归档合计从约 20.3 MiB 降为 6.6 MiB，
去重前后的 6,480 个逻辑输出逐字节一致，生成器和复验器均支持这些引用。
最终安装包另经生成器重新生成全部 3,240 个输出：Flat OPC 字节及 DOCX 包内全部文件内容
与冻结输出一致；重新生成的 DOCX 容器时间戳不同，不要求整个 ZIP 字节一致。
这不证明动态样式修改后的自动重算、其他格式、完整 Font API 或字体效果的视觉等价。

### PDF 隐藏文字与读取值

DOCX 的 `Font.hidden` 读取值与最终 PDF 可见性并非总是相同。固定 26.9.0 的
[81 个官方输出](benchmarks/font-hidden-rendering-26.9.json)中，有两个默认字体、段落和
字符样式组合的 getter 为真，而 PDF 仍显示目标文字。LDM 用 `hidden_rendering`
保存这种继承差异，`render_hidden` 返回 PDF 使用的状态；
[32 个默认样式引用样本](benchmarks/hidden-style-contexts-26.9.json)还区分显式引用
和隐式使用默认段落样式，两者不能在解析时合并。缺少该元数据的旧模型使用
`hidden`。直接给 `hidden` 赋值（包括赋回同一值），或 `model_copy(update={"hidden": ...})`，
会清除继承的渲染差异。模型 JSON 往返保留该状态。

`hidden_explicit` 保留 Run 的直接隐藏声明；`ParagraphFormat.style_explicit` 保留
段落样式是否被显式引用。给段落的 `style_name` 或 `style_identifier` 赋值，或通过
`model_copy(update=...)` 修改它们，会清除隐式来源标记。
[452 个 DOCX／Flat OPC 往返输出](benchmarks/hidden-font-roundtrip-26.9.json)经官方冷回读与
PDF 转换，属性差异从 20 个、目标文字可见性差异从 52 个降为零；直接声明与
段落样式引用另经独立 XML 检查。官方冷回读的可见目标字形也经栅格化检查，
冷回读 PDF 未归档，不代表全页面视觉等价。

PDF 的正文、表格、旋转单元格、文本测量和诊断使用同一状态，隐藏的文字、字号和
分页符不参与可见文本布局。此专项尚不证明动态样式编辑、其他字体效果、全部格式的
来源保留或全页面视觉等价；含形状的旋转单元格仍明确拒绝。

### 字段可见性

共享可见性过滤按每层字段的指令/结果状态处理嵌套字段，PDF 不绘制字段指令区的
文字、图片或文本框；可见字段结果继续输出。图片测量与分页使用相同边界，
源模型和原包内容不会因此被删除。历史对照见[复验记录](benchmarks/field-shape-visibility.json)。

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
网格/边线近似，不保证 Word 的完整表格排版。

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
直接构造 LDM 段落时，用 `children` 提供有序子节点；`runs` 是读取视图，不是构造参数。
未知／非法子节点与错误容器会显式拒绝，不再静默移除；校验和旧字段迁移不会修改传入字典。
格式读取中的未知内容仍沿既有保留与诊断路径处理。
旧 LDM 的 `header_paragraphs` / `footer_paragraphs` 在没有 section 时建立第一节，
并迁入该节对应的默认页眉／页脚。JSON 中这些字段是所有节的段落镜像，完全一致时不重复迁入。
同时提供 section 页眉／页脚和不同的旧字段内容会明确报错：旧列表无法表达目标节，
调用者须选择一组结构或将内容明确放入目标节；不会覆盖或静默忽略任一组。
DOCX／Flat OPC 读取和保存分别保留每节显式页眉／页脚，缺少引用继续使用 OOXML 的前节继承规则。
`HeaderFooter.header_footer_type` 的 0/1 是默认页眉／页脚，2/3 是首页，4/5 是偶数页；
每份内容有独立图片、超链接关系，显式空故事可中断继承。首页开关按节保存，偶数页开关是文档级设置。
源文件的节引用决定页眉／页脚角色；关系类型或故事根元素不一致时，恢复后仍按节引用保存。
这一映射不表示 PDF 已支持每节／首页／偶数页排版；PDF 仍对相关简化发出诊断。

LDM 的 `compatibility_mode` 默认 15；DOCX 按官方 URI 读取并写出兼容模式，缺失/格式无效时按 12。

`Document.compatibility_options` 提供 71 个布尔偏好的读取与编辑，以及
`optimize_for(dk.settings.MsWordVersion.WORD2010)` 等八个 Word 版本配置。
配置会更新兼容声明，并联动现有手动换行对齐及表格兼容模式调用；`ul_trail_space` 保留原值。
新建 `Document()` 使用已有空白模板的兼容声明，缺省模式为 12；直接构造 LDM 的旧默认值不变。
原始 `SourceSettings` 快照保持不变，已识别的重复布尔声明在保存时取末项，false 声明省略并产生通用规范化诊断。
普通保存保留未知兼容记录，并按已有声明集合补齐现代缺省项；显式优化会重建 typed `compatSetting` 列表，
删除原列表中的未知名称、URI、重复项及其扩展内容，并给出不含正文的 `ContentLossWarning`。
原快照与失败保存保持不变，foreign sibling 节点仍保留；当前偏好及优化状态可通过 LDM JSON 往返。
`transparent_metafiles` 与四个现代偏好的显式编辑存在官方暖态/冷态区别；不会伪造不存在的 `compatibility_mode` 公共属性。
此闭环未证明全部选项的排版、图片插入、历史 DOC/RTF 行为或完整 `settings` 命名空间已对齐。
旧模式小于 15 时，左对齐流式表格按首单元格边距外移；居中/右对齐忽略表格缩进。
这会改变此前忽略旧模式的 PDF 横坐标；设为 15 可使用现代位置。该字段不代表完整版本兼容，
浮动/嵌套定位和边框度量仍有限制，见[表格定位验收](benchmarks/table-compatibility.json)。

表格、单元格与表格样式的四侧边距以 pt 为单位：`None` 表示未设置/继承，`0` 表示显式零边距。
DOCX 读写保留零值，PDF 测量与绘制采用单元格覆盖、表格默认、渲染器回退的相同顺序；
DOCX 表格默认边距包括已支持的样式继承。表格上下边距是单元格内部间距，不再额外增加表外空白。
旧 LDM JSON 若用 `0` 表示默认间距，应改为 `null`；显式零值现在会改变布局。
条件表格样式及行级边距例外仍未实现，见[边距验收](benchmarks/cell-margins.json)。
`Table.bidi` 保留已支持的 DOCX RTL 表格方向，镜像视觉列、对齐及边框，同时保留逻辑内容顺序。
`start/end` 边距归一为 left/right 字段，DOCX 输出 left/right；表格方向不会自动启用文字塑形。
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
import dockiln as dk
from dockiln.writers.pdf import PdfMissingGlyphWarning

warnings.simplefilter("error", PdfMissingGlyphWarning)  # 不接受缺字 PDF
options = dk.saving.PdfSaveOptions()
options.fallback_fonts = ["/trusted/fonts/extra-unicode.ttf"]
dk.Document("report.docx").save("report.pdf", options)
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
| `dk.loading.DocumentLoadWarning` | 缺失脚注/尾注正文、对象、altChunk、修订、简单字段、内容控件、公式等尚未完整保留的构造 |
| `DocxWriterLossyWarning` | 未知模型节点等已知写出损失；无对应正文、冲突 ID 或不支持位置的注释会在发布输出前明确拒绝 |
| `dk.ContentLossWarning` | 已知内容/布局损失的公共基类；用于统一过滤加载、DOCX/Flat OPC、PDF 的已知损失 |

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

DOCX/Flat OPC 中被读取的段落缩进、段间距/行间距、编号缩进和 dxa 表格/单元格边距
必须是可表示的 XML 整数 twip。缺失属性保持未设置；显式空值、小数、下划线数字、
非 ASCII 数字、指数形式及数值溢出抛出 `ValueError`，不会静默改成零或继承值。
错误消息及默认异常链隐藏原始属性值；同一缩进的后续有效别名不能掩盖损坏数值。
原包 DOM 保存仍可保留原始文件；转换失败保留已有输出。
这属于本项目的损坏输入拒绝规则，官方对同类文件的恢复行为尚未测量。

普通 `w:sz` 字号统一用于 DOM、样式/默认格式、转换 Run 和条件表格样式读取：
支持已测得的整数、小数、指数及 `pt/in/cm/mm/pc/pi`，截断为整半点后须为非负且可表示，
例如 `25.5` 半点为 12.5pt，`6.25pt` 为 6pt，`25.4mm` 为 72pt。
显式空值、非法数字/单位、负值和超出当前读取范围的值私密抛出 `ValueError`，
不再静默回退为默认字号；缺失属性仍可继承。
[固定官方观察](benchmarks/font-size-loading-26.9.json)记录了不同的损坏输入恢复结果，
本项目尚未实现这些恢复语义；复杂文字 `szCs`、完整字体继承和布局仍未验收。
66 个正值目标 Run 的转换输出已由官方包实际回读。此前缺失 `rPrDefault` 在转换后
成为空组、隐式样式字号从 11pt 变为 10pt 的问题已修复；
[87 个既有输入的 348 个输出](benchmarks/font-default-presence-26.9.json)覆盖 DOCX、
Flat OPC 和模型 JSON，独立 XML 及官方回读核对默认组、Run/样式字号与直接声明状态。

转换模型的 `Document.doc_defaults_rpr_present` 保留默认组是否存在；`False` 表示缺失，
`True` 表示存在（包括空组），`None` 表示旧 JSON 或新建模型没有来源信息，沿用原有
生成默认组的规则。缺失组的普通继承字号为 11pt，存在但没有普通字号的组为 10pt；
原始 `doc_defaults_font` 仍只保留已声明的属性。新设置默认字体时会生成默认组。
`Font.size_explicit` 的 `True`／`False` 分别记录当前层声明／继承的普通字号；`None`
沿用旧模型按值差异输出的规则。若编辑后要显式写入与继承值相同的字号，设置该标记为
`True`；改变字号值仍会写出。来源标记随模型 JSON 保存，旧 JSON 无法恢复丢失的来源信息。
普通字号 getter 会读取当前样式、条件表格样式和直接格式；修改这些声明后，
已有 Run 的字体视图随之更新，不把继承值写成直接格式。新建转换模型 Document 时，
未绑定的 typed 节点保持已有首次采用行为；传入图已有文档字体视图时会复制，
以免重绑定另一文档持有的字体视图。此时后续应在创建后的 Document 节点上编辑。
DOM 的 Run／Style `font.size = 0` 与文本范围 `apply_font(size=0)` 写出零字号；
Run 的 `font.size = None` 仍表示清除直接声明，继续继承样式或文档默认值。
`model_copy(deep=False)` 保持共享节点及原字体绑定；浅拷贝 update 不提供独立的有效字号视图。
需要独立编辑默认字号或样式时使用 `deep=True`，它会重建新图的绑定。转换模型非负有限数值字号
按最近的半点舍入（中点向上）；XML 半点数仍按上述读取规则处理。
显式声明或后续赋值的零字号保持为零，PDF 使用 0.5pt 绘制；旧模型未标记的
`Font(size=0)` 仍表示未设置。Code/Quote 及普通段落、表格共享受测字号绘制路径。
普通非表格段落的 PDF 字号读取当前继承声明；禁用 Normal 的默认样式标记仍保留其受测字号 fallback，直接与字符样式字号继续优先。
DOM 的普通有效字号采用相同的禁用 Normal fallback；这不启用其默认标记，也不改写段落样式或 Boolean 字体声明。
成功保存 DOCX/Flat OPC 后，字体视图采用实际写出的兼容声明，加载时的
`source_settings` 快照保持不变；`to_bytes()` 和失败的输出不提交此状态。
`optimize_for()` 的受测字号视图立即更新，且保留到后续成功保存；该 warm 状态
与实际保存的兼容 XML 视图分开，随 JSON 和深复制保留。基准产品保存 PDF 后改变部分表格字号 getter
的行为尚未实现；Quote 的既有缩进与完整字体、分页及字形保真仍未完成验收。
这两个标记只覆盖受测默认组与普通字号，不代表全部字体属性的设置来源已完整保留；
字体 JSON 通过 `_fields_set` 记录已设置字段，保持既有默认字段和值；新 JSON 不再把
字符样式中未声明的粗体/斜体写成显式 `False`。该元数据跟随复制和后续赋值，非法字段名、
非列表或引用不存在值的记录会拒绝。旧完整 JSON 不含该信息时仍按原来的字段设置规则读取。
`bold_explicit`／`italic_explicit` 与 `size_explicit` 一样记录当前层的普通属性声明来源，
显式开关即使与写出基线相同也会保留；需要显式写入同值编辑时将对应标记设为 `True`。
转换的普通粗体/斜体继承复用原包 DOM 的受测样式类别规则，
[745 个既有输入的 1,490 个模型 JSON 输出](benchmarks/font-json-origin-26.9.json)
经官方冷回读和独立 XML 核对；修复前的 194 个差异保留为对照。
这不代表全部字体属性的直接声明来源、表格上下文或完整编辑语义已验收。
21 个简单样本的 PDF 字号独立测量通过，完整分页、字形和商业渲染等价仍未验收。

DOCX/Flat OPC 的字符单位缩进保留到模型，但尚未换算为排版点值。加载时汇总 `leftChars/rightChars`、
`startChars/endChars`、`firstLineChars/hangingChars`，发出 `load.character_indents_ignored`
损失诊断；扫描正文、样式、编号和附属 story，包含显式零值，诊断不含正文或属性值。
`ContentLossWarning` 错误过滤及 CLI `--strict` 在输出前拒绝；原包 DOM 保存仍保留这些属性。
`ParagraphFormat` 和编号 `ListLevel` 保留 `character_unit_left_indent`、`character_unit_right_indent`
和 `character_unit_first_line_indent`，
以字符为单位；`None` 表示未设置，`0` 保留显式零值。样式继承、编号定义/覆盖、DOCX/Flat OPC
及模型 JSON 保留已支持路径中的这些值（各输出格式的排版限制仍适用），输出归一为 `leftChars/rightChars/firstLineChars` 或 `hangingChars`。
直接修改转换模型后保存 OOXML 会写入整数百分之一字符；非有限值、错误类型及超出有符号
32 位百分之一字符范围的值在输出前拒绝。整数换算不依赖调用方的 Decimal 精度；
损坏的字符属性错误不在异常链中暴露原始值。原包 DOM 已支持字符直接值和样式继承值的只读 getter；
字符 setter、与基准点值联动和渲染仍未实现或未验收。
PDF writer 对模型中的字符值另发 `pdf.character_indents_ignored`；错误过滤可拒绝并保留原输出。
见[检测与拒绝](../tests/test_character_indent_diagnostics.py)及[模型往返](../tests/test_character_indent_roundtrip.py)。
这些是单位值保留和损失检测，不能视为字符缩进排版验收。

`Document.save()` 的 PDF、DOCX 和 Flat OPC 文件在同目录暂存并原子替换；转换、写入或替换失败
不会覆盖原主文件，已有输出符号链接、非普通文件及不可写文件会被拒绝。新文件默认权限为 `0600`，替换已有文件时保留其权限位。
原子替换不等于完整断电恢复保证；流输出不具有文件原子替换契约。

## 有界单任务 CLI

```bash
python -m dockiln.convert report.docx report.pdf \
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
python -m pip install "dockiln[shaping] @ git+https://github.com/shenyankm/DocKiln.git@dev"
```

实际绘制和行宽/行高测量同时启用 shaping。实现使用 fpdf2 内部 bidi/断行 API，升级 fpdf2 后须复验。
普通 `Document.save()` 没有进程超时/内存限制；服务端请使用此入口或自己的任务隔离层。
CLI 不是 Web 服务、队列或全局并发控制器；应用需限制同时启动的作业数，并配置容器内存/磁盘/网络配额。

## 可选 LibreOffice 原文件路径

```python
from dockiln.libreoffice import convert_to_pdf

convert_to_pdf("report.docx", "report.pdf", timeout=60)
```

或者通过同一个受限 CLI：

```bash
python -m dockiln.convert report.docx report.pdf --backend libreoffice
```

需自行安装 LibreOffice，`soffice`/`libreoffice` 在 PATH 上；也识别 macOS 标准应用安装位置。
此路径传入**原始 DOC/DOCX/RTF 字节**，不经过可能丢信息的 LDM 重建。每个任务使用私有用户配置目录，
验证进程状态及 PDF 头/结束标记后才发布文件。内置 RTF reader 仍只支持 OLE2/DOC-backed 文件；现在会对标准文本 RTF 明确报错并提示此入口。

这是独立可选入口，不改变 `Document.save()` 默认后端。高保真仍受 LibreOffice 兼容性和部署的
**系统字体**影响；它不直接使用本库 WOFF 字体。`--strict`、`--fallback-font` 和 `--text-shaping` 不适用于该后端。
本仓库既测试协议/异常，也提供真实渲染测试。已在 macOS / LibreOffice 26.2.6 验证中文 DOCX、标准文本 RTF、
原始 DOCX 脚注和受限 CLI；原文件转换中可保留脚注，内置 PDF 路径仍会报告脚注/尾注省略。另有新增 DOM 图片、超链接与水平合并的真实 LibreOffice 渲染用例。未安装 LibreOffice 的环境跳过真实测试。
这些用例证明入口可用和已测内容保留，不证明所有文档与 Microsoft Word 视觉一致。

**私有 profile 与资源限制不是安全沙箱。** 原生 Office 解析器、宏、外部链接与文件系统访问必须由
应用/容器沙箱约束；对不可信输入关闭网络、隔离可读写目录、以非特权账户运行并保持 LibreOffice 更新。
尤其不要将普通 `convert_to_pdf()` 当作具备硬资源/外部资源隔离的服务器入口。

## 保存选项逐字段审计

范围为 `saving.py` 的 `PdfSaveOptions`、`OoxmlSaveOptions` 及其 `OutlineOptions`。
下表记录源码行为与复验入口，不表示已完整验收每项；全矩阵效果、标量边界和跨格式组合仍需完整选项验收。
历史枚举拒绝与未实现请求的诊断见[记录](benchmarks/save-option-contract.json)；其中 Markdown 专项已不适用。
后续标量及内容专项按各行记录核对，不能升级为全矩阵已通过。
支持枚举值不等于对应功能已实现：有效但未实现值仍按表中限制处理。

### PdfSaveOptions

| 字段 | 当前行为/限制 | 证据或复验入口 |
|---|---|---|
| `compliance` | 只设置 PDF 版本；PDF/A、PDF/UA 请求警告，非标准认证 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `export_document_structure` | 结构输出开关；完整辅助技术与 PDF/UA 标准验收未完成 | [test_pdf_structure_pages.py](../tests/test_pdf_structure_pages.py) |
| `image_compression` | AUTO 保留有意义的透明像素；显式 JPEG 合成白底并诊断透明度损失，严格模式可拒绝 | [test_pdf_image_compression_contract.py](../tests/test_pdf_image_compression_contract.py) |
| `jpeg_quality` | 0..100 整数，布尔/非整数/越界拒绝；作用于 JPEG 编码，AUTO 的透明图片保留原字节 | [test_pdf_image_compression_contract.py](../tests/test_pdf_image_compression_contract.py)、[复验](benchmarks/pdf-image-compression-contract.json) |
| `text_compression` | NONE/FLATE 页面内容流，字体/图片流独立 | [test_pdf_stream_options.py](../tests/test_pdf_stream_options.py) |
| `embed_full_fonts` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `use_core_fonts` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `font_embedding_mode` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `page_mode` | 写入 PDF PageMode，默认 USE_OUTLINES；不自动创建大纲/图层/附件，阅读器可以忽略 | [test_pdf_stream_options.py](../tests/test_pdf_stream_options.py) |
| `color_mode` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `export_bookmarks_outline` | 书签输出；显式 outline 层级优先 | [test_pdf_outline_options.py](../tests/test_pdf_outline_options.py) |
| `outline_options` | 下表六项；不是完整 Word TOC 契约 | [test_pdf_outline_options.py](../tests/test_pdf_outline_options.py) |
| `preserve_form_fields` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `memory_optimization` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `zoom_factor` | ZOOM_FACTOR 下使用：有限非负 int/float（拒绝 bool），0 保留默认打开行为；正值按百分比写成普通十进制，PDF 数字最长 47 字符。内存/文件/直接 writer、错误原文件保留及独立解析已验证；实际阅读器打开行为未验收 | [缩放契约](benchmarks/pdf-zoom-contract.json)、[writer](../src/dockiln/writers/pdf/writer.py) |
| `zoom_behavior` | PDF 打开动作；部分行为在序列化后改写 | [src/dockiln/writers/pdf/writer.py](../src/dockiln/writers/pdf/writer.py) |
| `display_doc_title` | PDF viewer preference；阅读器可忽略 | [src/dockiln/writers/pdf/writer.py](../src/dockiln/writers/pdf/writer.py) |
| `fallback_fonts` | 可信字体路径；覆盖、塑形及源字体替代有独立限制 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `text_shaping` | 可选 HarfBuzz 塑形；依赖/字体及格式边界有限制 | [test_pdf_shaping.py](../tests/test_pdf_shaping.py) |

### OoxmlSaveOptions

| 字段 | 当前行为/限制 | 证据或复验入口 |
|---|---|---|
| `save_format` | 支持 DOCX 和四种 Flat OPC 变体；其他值在输出前拒绝 | [test_save_option_formats.py](../tests/test_save_option_formats.py) |
| `compliance` | ECMA376_2006 与 ISO29500_2008_TRANSITIONAL 使用同一输出；Strict 抛出 NotImplementedError | [test_bounded_conversion.py](../tests/test_bounded_conversion.py) |
| `compression_level` | 四档 ZIP deflate；不是内容精简 | [examples/working_with_ooxml_save_options.py](../examples/working_with_ooxml_save_options.py) |
| `zip_64_mode` | NEVER/IF_NECESSARY；ALWAYS 仍等同按需，但现在警告 | [test_unsupported_save_requests.py](../tests/test_unsupported_save_requests.py) |
| `pretty_format` | XML 缩进；语义与全部部件仍待矩阵复核 | [examples/working_with_ooxml_save_options.py](../examples/working_with_ooxml_save_options.py) |
| `reference_docx` | 导入参考样式，非参考正文/资源/页面设置 | [test_reference_styles.py](../tests/test_reference_styles.py) |

### OutlineOptions

| 字段 | 当前行为/限制 | 证据或复验入口 |
|---|---|---|
| `headings_outline_levels` | 非负整数，拒绝 bool；标题深度大于六的值按现有规则夹紧 | [大纲契约](benchmarks/pdf-outline-contract.json)、[效果](../tests/test_pdf_outline_options.py) |
| `expanded_outline_levels` | 大纲展开范围 0..9 整数，拒绝 bool；0 全折叠，1 展开 PDF 树首层，依此类推；阅读器可覆盖初始状态 | [大纲契约](benchmarks/pdf-outline-contract.json)、[效果](../tests/test_pdf_outline_options.py) |
| `default_bookmarks_outline_level` | 0..9 整数，拒绝 bool；0 且空映射保留 export_bookmarks_outline 退回规则，非商业版完全相同契约 | [大纲契约](benchmarks/pdf-outline-contract.json)、[效果](../tests/test_pdf_outline_options.py) |
| `bookmarks_outline_levels` | 字符串名称到 0..9 整数的 Mapping，显式 0 省略该书签，隐藏名称仍跳过；每次保存重新校验 | [大纲契约](benchmarks/pdf-outline-contract.json)、[效果](../tests/test_pdf_outline_options.py) |
| `create_outlines_for_headings_in_tables` | bool；表格标题输出导航 | [大纲契约](benchmarks/pdf-outline-contract.json)、[效果](../tests/test_pdf_outline_options.py) |
| `create_missing_outline_levels` | bool；默认压缩缺失层级，True 插入空白中间项；先验证请求层级防止无界分配 | [大纲契约](benchmarks/pdf-outline-contract.json)、[效果](../tests/test_pdf_outline_options.py) |

### 选项验收边界

逐字段契约不等于全矩阵验收：标量边界、显式默认值和跨选项组合仍须检查实际输出。
直接 writer 使用对应的 `PdfSaveOptions` 或 `OoxmlSaveOptions`。
ZIP64 强制输出仍未实现；诊断说明退回路径，不计为功能完成。

设置关系选择先跳过空目标，再按跨类型重复 Id 的最后记录、UTF-16 顺序和准确 settings 类型选择首项。选中目标缺失、无法解析为包内路径或非 Word 设置根时使用默认值，与 `TargetMode` 无关，不回退到后续关系；相对 External 目标仅在包内解析，不执行网络或外部文件读取。图片、超链接及其它资源的既有验证不因此放宽。

加载的兼容性声明单独保存为不可变 `SourceSettings` 来源，区分缺少 `w:compat`、显式空块及原有未知扩展。JSON 往返和两种 OOXML 输出不再无条件添加模板兼容性开关；编辑已支持的模式/手动换行标志只替换其对应声明，新建模型仍沿用原模板基线。此来源只覆盖兼容性块，不代表完整 settings 资源保真。


## DOCX / Flat OPC 脚注与尾注

转换模型以有序 `NoteReference` 保存正文锚点，以 `SourceStory` 保存脚注/尾注正文；
`note_type` 区分普通注释、分隔线、延续分隔线和延续通知。普通负 ID 不再被当成分隔线丢弃。
段落与表格的顺序、受支持图片和超链接分别写入 footnotes/endnotes 部件及其关系侧车；
模型 JSON、DOCX 与 Flat OPC 连续保存保留这些值。`NoteMarker` 保留注释内的引用符号及分隔线符号，
`custom_mark_follows` / `reference_mark` 将定制标记写在锚点所在的同一个 run 中，保留其中的 CR、LF 与制表符。
`note_references` 是有序锚点的同对象视图；JSON 和深拷贝恢复共享对象，因此通过此列表编辑标记/格式会写回。
分别提供两组不一致的引用、替换成不同元数据的独立镜像或只清空一组时明确报错，须同步其位置和内容。

文档的 `note_properties` 和节的 `page_setup.note_properties` 分别保留 `footnotePr/endnotePr`
中的位置、编号格式（含 format）、起始编号、重启规则及分隔线 ID，按原顺序写回；
公开选项通过 `document.footnote_options/endnote_options` 与
`section.page_setup.footnote_options/endnote_options` 编辑，见下节；这不证明分页或脚注布局等价。
不能表示的属性或容器明确拒绝，避免悄悄舍弃设置。

注释 ID 按已校准的数值读取规则归一：ASCII 浮点值截断，超出有符号 64 位范围时饱和，
再按有符号 32 位值关联正文；错误词法和非 ASCII 数字映射到 0。
同角色归一后重复 ID、缺失正文、普通锚点指向特殊注释、未放入有序子节点的旧引用列表、
缺失的分隔线设置目标和循环正文都在输出前报错，已有输出文件保持原样。
这些明确拒绝与商业实现对部分损坏输入的恢复行为不同；不将拒绝计为成功保留。

内置 PDF 仍提示脚注/尾注省略；需要保留注释结构时使用 DOCX/Flat OPC 或原包 DOM 保存。
独立注释正文中的普通 Run、正文锚点和注释标记使用实时普通字号解析链，保留零字号、
段落/字符样式与直接声明的受测优先级；移动、JSON 和深复制后重新关联当前故事与表格上下文。
这不证明所有字体属性或注释 PDF 布局等价。
转换模型对未定义段落样式的普通字号回退到当前默认段落样式，并保留原样式引用；
包 DOM 的有效格式视图继续明确拒绝缺失样式。受测往返见
[test_conversion_notes.py](../tests/test_conversion_notes.py)；未覆盖的 Word 结构与视觉布局不在本阶段保真声明内。


## 可编辑脚注、尾注选项

`dk.notes.FootnoteOptions` 提供 `number_style`、`position`、`restart_rule`、
`start_number`、`columns`；`EndnoteOptions` 提供前四项。它们是文档或页面设置的关联视图，
不能直接构造。转换入口和 `DocxDocument` DOM 入口提供相同配置链路。

```python
import dockiln as dk

doc = dk.Document("input.docx")
doc.footnote_options.number_style = dk.NumberStyle.LOWERCASE_ROMAN
doc.footnote_options.restart_rule = dk.notes.FootnoteNumberingRule.RESTART_SECTION
doc.first_section.page_setup.footnote_options.start_number = 3
doc.endnote_options.position = dk.notes.EndnotePosition.END_OF_SECTION
doc.save("output.docx")
```

编号、位置、重启规则必须使用对应枚举；裸整数、布尔值和其他枚举会报 `TypeError`。
`start_number` 和脚注 `columns` 接受带符号 32 位整数，越界报 `OverflowError`。
这些 setter 在验证失败时保留原值和文件内容。起始编号的零与负值可以编辑，
不把 setter 接受范围误当成有意义的排版范围。

文档级 setter 更新所有节的同名配置，包括已单独设置的节。节级 setter 通常只更新该节；
尾注 `position` 的公开 setter 是例外，从任何节调用都会更新文档和全部节。
载入的原始节位置声明仍可与文档声明不同；读取不会把它们强制合并。

默认脚注为阿拉伯数字、页底、连续编号、起始编号 1、栏数 0；默认尾注为小写罗马数字、
文档末尾、连续编号、起始编号 1。转换模型 JSON 和文档的 `model_copy(deep=True)`
保留选项，并建立副本自己的关联关系。单独深复制 `PageSetup` 时只保留声明值，
不携带原文档的关联；加入目标转换模型后重新绑定。
`dk.NumberStyle` 现在是整数兼容枚举，旧 `KANJI_TRADITIONAL_2` 拼写仍保留，
同时提供 `KANJI_TRADITIONAL2`。

脚注 `columns` 在编辑状态下保留 setter 的值。Word 2013 及更新兼容模式
（`compatibility_mode >= 15`）只把正值写入各节的 `w15:footnoteColumns`。
重新载入后，文档级栏数为 0，节级栏数来自文件；较旧模式和非正值不写出。
如需使用此文件能力，在转换入口先调用
`doc.compatibility_options.optimize_for(dk.settings.MsWordVersion.WORD2013)`。
DOM 入口沿用原文件的兼容模式。

`NumberStyle.SIMP_CHIN_NUM4` 可以编辑，DOCX/Flat OPC 没有该值的编号格式投影，
保存时省略该格式，重载后回到对应文档或节的默认/继承值。其他公开编号样式写入对应格式。
这两种编辑值与文件值的区别遵循实际保存行为；不承诺它们在旧格式中全部持久化。

`get_type()` 属于商业包的 CLR 反射接口，本实现明确报 `NotImplementedError`；
没有伪造商业运行时类型。选项保存覆盖 DOCX 与 Flat OPC 的声明与关联，
不代表视觉分页、实际重启编号的排版、PDF 脚注布局或完整商业 API 已等价。
回归入口：[test_note_options.py](../tests/test_note_options.py)。
