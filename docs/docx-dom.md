# 原包保真的 DOCX DOM（资源关系与受限水平合并）

新增入口 `aspose.words_foss.DocxDocument`，与现有 `Document → LDM → writer` 转换路径并行。
它以原始 OOXML 为唯一权威状态，节点是 XML 的类型化视图，不维护另一份同步模型。
**这是可逐步扩展的基础 DOM，不是完整 Aspose.Words DOM，也不是 Word 排版引擎。**

`DocxDocument` 也可从路径或二进制流加载 Flat OPC XML 包。四种文档、宏文档、模板、宏模板
变体均还原 XML 与二进制部件，保留祖先命名空间及主部件 Content Type；沿用原包 DOM 的
编辑边界。`save()` / `to_bytes()` 输出 ZIP OOXML 包，需使用与主部件类型对应的
`.docx` / `.docm` / `.dotx` / `.dotm` 扩展名；不保留输入 XML 的字面序列化。
`to_flat_opc()` 返回 UTF-8 XML 字节，`save_flat_opc(path)` 原子保存 Flat OPC XML；
两者保留部件类型、XML 元素、注释和二进制资源，但 XML 声明、编码及序列化会规范化。
为避免官方 Flat OPC 读取器拒绝输出，XML 元素内部的处理指令会省略，并发出
`ContentLossWarning`（诊断码 `flat_opc.processing_instruction_omitted`）；诊断不包含指令内容。
根元素外的处理指令保留，原始 DOCX 与当前 DOM 不受此输出投影影响。
将该警告升级为错误可拒绝转换；原子保存失败时不会覆盖已有目标文件。
宏部件仅作为原始资源保留，不执行宏。签名包仍按原有边界拒绝修改。

现有 `Document` 转换入口也能识别 Flat OPC 字节/流，路径可使用 `.xml`，或显式指定
`LoadFormat.FLAT_OPC`、`FLAT_OPC_MACRO_ENABLED`、`FLAT_OPC_TEMPLATE`、
`FLAT_OPC_TEMPLATE_MACRO_ENABLED`。显式 `TEXT` 保留字面文本行为。
`Document.save(path, SaveFormat.FLAT_OPC)` 和 `to_bytes(SaveFormat.FLAT_OPC)`
以及其余三种变体可写出 Flat OPC；也接受对应的 `OoxmlSaveOptions`。必须显式指定格式，
不把 `.xml` 后缀推断为 Flat OPC（官方默认 `.xml` 实际选择 WordML）。
它仍经过 LDM，具有现有内容损失边界；VBA project 被省略时记录
`load.vba_project_omitted`。输出宏类型不意味着 VBA 被导出；需要保留未知部件或
宏资源的编辑应使用 `DocxDocument`。
四种原生试用样本与部件比较记录见 [加载验证](benchmarks/flat-opc-loading.json)，
导出与官方包回读见 [保存验证](benchmarks/flat-opc-saving.json)。
未以这些普通样本验收复杂宏、模板全部语义或视觉排版。

## 单文本节点字面替换

```python
from aspose.words_foss.docx_edit import replace_text

count = replace_text("original.docx", "edited.docx", {"客户公司": "新的客户公司"})
```

此窄接口同时替换 `word/` XML 部件内单个 `w:t` 节点中的字面文本，不跨 Run、段落，
不支持正则或格式编辑。所有 key 必须出现；重叠 key 优先较长值，新文本不会再次参与替换。
返回匹配总数。未修改 XML 字节及其他部件 payload 保留，ZIP 重新打包，不保证整个文件字节相同。
源和目标可相同，失败保留原目标。沿用 ZIP/大小检查，禁止 DTD，仅接受 UTF-8 XML，
拒绝签名包及包含 Tab、换行或非法 XML 字符的替换值。宏和外部关系仍保留，不能用于安全清洗。
跨普通 Run 替换及局部格式编辑使用下述 DOM 接口。

## 读取、格式与跨 Run 替换

```python
import aspose.words_foss as aw

editable = aw.DocxDocument("original.docx")
paragraph = editable.body.paragraphs[0]

print(paragraph.text)
print(paragraph.xml)  # 只读 XML 快照，不暴露可绕过修改检查的公共 XML 对象

# 在当前段落的直接文本 Run 之间匹配；返回匹配数量，找不到返回 0。
count = paragraph.replace_text("客户公司", "新的客户公司")
if count:
    paragraph.runs[0].font.bold = True
paragraph.paragraph_format.alignment = "center"  # OOXML 值，不是 aw.ParagraphAlignment 枚举

editable.save("edited.docx")
raw_bytes = editable.to_bytes()
```

`Run.text` 可直接修改。当前写入只接受普通 XML 文本，不接受 Tab、换行、控制字符或非法 Unicode。
Run 含图片、字段标记、换行等非普通文本节点时，文本设置明确拒绝；不会先删除其他内容。

`Paragraph.replace_text(old, new)`：

- 字面、非重叠匹配，替换后的文本不会再次参与本次匹配。
- 跨同一段落内的直接 Run 和 `w:t`，插入文本沿用匹配起点 Run 的格式。
- 原 Run 和文本节点保留；被覆盖的后续文本节点可能变为空，不自动合并 Run。
- 超链接、内容控件、字段标记、书签、修订、Tab、换行等复杂内联结构明确拒绝。
- 不跨段落、不支持正则、不自动计算字段，也不生成修订。

`Paragraph.text` 是段落内原始 `w:t` 的串联（Tab/换行映射为字符），不等于 Word 的最终可见文本：
包含隐藏文本和嵌套容器中的文本，不做修订视图过滤或字段计算。
`Run.text` 读取直接 `w:t`、Tab 和换行，不将字段指令当成普通文本。

## 直接格式与继承

```python
run = editable.body.paragraphs[0].runs[0]
run.font.bold = True   # 显式开启
run.font.italic = False  # 显式关闭
run.font.bold = None   # 移除直接设置，恢复继承

run.font.size = 12.5  # 点数，必须是正数且为 0.5 的整数倍；None 恢复继承

# ID 必须是主文档 styles 关系指向的部件中存在的 paragraph style。
# 没有这个 ID 时抛出 ValueError，不自动创建样式。
# paragraph.paragraph_format.style_id = "Title"
paragraph.paragraph_format.style_id = None
paragraph.paragraph_format.alignment = None
```

`run.font` 和 `paragraph.paragraph_format` 的 getter 返回**直接设置**，不存在时是 `None`。
现在支持字号、13 个字体布尔属性、段落对齐和段落/字符样式 ID。字体布尔属性为
`bold`、`italic`、`bold_bi`、`italic_bi`、`all_caps`、`small_caps`、`strike_through`、
`hidden`、`outline`、`shadow`、`emboss`、`engrave`、`no_proofing`；赋值只接受
`True`、`False`、`None`，数字和字符串会在修改前拒绝。`run.font.style_id` 只接受已存在的
character style，`paragraph.paragraph_format.style_id` 只接受 paragraph style；两者支持 `None` 移除引用。
不把样式继承结果物化为直接格式，也不自动创建样式。

内置字体和段落格式句柄只允许赋值已声明的可写属性。未实现属性（如当前的
`font.color`）、拼写错误及覆盖方法/内部排序常量会抛出 `AttributeError`，不会修改文档。
这些句柄不再允许附加任意 Python 属性；调用方的自定义信息应单独存储。
此保护不表示颜色等缺失属性已实现，它们仍须补齐加载、编辑、保存与渲染链路。

非布尔类型的字体开关与段落分页开关赋值现在抛出 `TypeError`；此前为 `ValueError`，
捕获旧异常的调用方需要调整。非法 OOXML 属性值仍抛出 `ValueError`。
[官方错误观测](benchmarks/dom-font-boolean-errors-26.9.json)覆盖 297 组输入，
其中 264 组错误类型差异已修正。Run 的直接字体接受 `None` 清除声明是本项目 DOM
接口的明确功能；商业 Font 拒绝 `None`，其完整接口契约尚未对齐。

### 样式集合与字体编辑

```python
style = editable.styles.get_by_name("Heading 1")
if style is not None:
    print(style.style_id, style.name, style.type)
    print(style.font.bold, style.font.italic, style.font.size)
    style.font.bold = True
    style.font.size = 17.5
    style.direct_font.italic = None  # 清除该层，恢复基样式/文档默认值
    style.paragraph_format.alignment = "center"
    style.paragraph_format.keep_with_next = True
    style.direct_paragraph_format.alignment = None  # 清除当前层的对齐
for style in editable.styles:
    print(style.name)
```

- 集合按实际 styles XML 顺序遍历，支持 `len()`、`get_by_id()` 和精确名称
  `get_by_name()`；未找到返回 `None`，同名歧义抛出 `ValueError`。不自动创建内建样式。
- `Style.font` 读取最近的样式层及文档默认字号、上述布尔属性，设置写入当前层；`direct_font`
  读取当前层，未设置返回 `None`，可赋 `None` 清除。`font` 的这些 setter 拒绝 `None`。
  字号 setter 仍限正的半点数；读取普通 `w:sz` 时支持已测得的整数、小数、指数及
  `pt/in/cm/mm/pc/pi` 单位，按官方观察截断不足半点的部分；结果须为正且可表示，原包 XML 不因读取改变。
  [126 个固定输入及 105 个官方保存输出](benchmarks/font-size-loading-26.9.json)中，
  66 个正值字号的 DOM、转换模型和保存重开已对照；零/负值及损坏值仍按本项目规则拒绝，
  与官方恢复行为不同，复杂文字字号和完整排版未验收。没有声明字号时使用受测普通字号默认值：有 `rPrDefault` 为 10pt，
  没有该组为 11pt。直接值 getter 仍返回 `None`，读取不会写入 XML。
  转换模型也保留默认组存在状态和普通字号直接声明来源；
  [348 个 DOCX/Flat OPC/模型 JSON 输出](benchmarks/font-default-presence-26.9.json)
  经独立 XML 与官方包实际回读，缺失组不再迁移为 10pt 空组，继承字号保持未直接设置。
  旧模型 JSON 没有来源标记时仍按旧生成规则处理，详见[转换说明](enhanced-conversion.md)。
- 段落/字符样式的这三项编辑已按 1482 次固定官方观察验证；尚无新增、删除、重命名或
  修改继承关系接口，不提供完整 Style/Font API。表格/列表字体编辑和嵌套 rStyle 修改拒绝。
- 段落样式的 `paragraph_format.alignment` 沿文档默认和 basedOn 链解析，未声明时为
  `left`；四种基本对齐的 108 次编辑/保存/重开已有官方固定观察。`direct_paragraph_format`
  保留当前层的 `None`，可用 `None` 清除；继承 getter 的 setter 拒绝 `None`。
  字符/列表样式不提供段落格式，表格样式仍拒绝，新增分页和数值属性见下；其余段落属性和嵌套 pStyle 编辑未实现。
- `keep_with_next`、`keep_together`、`page_break_before`、`widow_control` 支持段落直接值、
  样式继承值和有效值；前三项默认 `False`，孤行控制默认 `True`。直接值为 `None` 表示未设置，
  可用 `None` 清除；样式继承 setter 只接受 `bool`。324 个非首段输入、1296 次官方编辑已对照
  保存/重开和本项目 LDM 属性。另有 [10 个原生 PDF 对照](benchmarks/pagination-rendering-26.9.json)，
  覆盖四个开关的 False/True、两侧孤行和分页位置；普通字体 ASCII、固定 12pt 行距的非首页面
  原点/字宽容差为 0.02pt，144dpi 黑色像素差比例上限 1%，受测页面实测为 0。
  排除试用首页面和彩色水印；字体描述元数据、复杂文字、表格/图片和自动换行仍未验收。
  官方试用包会在加载时关闭受测首段的显式
  段前分页；18 个加载/新建对照观察表明，首段值变化与加入试用提示段落同时出现，
  已有前置段落时保留。持许可证的对照仍未执行，本项目保留原始值，差异继续记为未决。
- `left_indent`、`right_indent`、`first_line_indent`、`space_before`、`space_after` 以 pt 读写，
  继承值和有效值默认 0；直接值的 `None` 表示未设置，可清除。继承 setter 拒绝 `None`，
  布尔和字符串也拒绝；非有限值在修改前拒绝。段间距在量化前限定为 0–1584pt，
  越界抛出 `RuntimeError` 并保留原值。有限缩进以有符号 32 位 twip 边界夹紧；
  有限值乘以 20 后若发生浮点溢出，按官方观察得到 -2147483648 twip。
  值按 1/20pt 四舍六入五成双量化，如 12.375→12.4。负首行缩进写为 hanging，
  与 firstLine 同层出现时取 XML 中后出现的属性；编辑保留 ind/spacing 中未涉及的属性。
  点值读取拒绝空值、非 XML 整数和数值溢出，同一属性的损坏别名也拒绝；
  错误及默认异常链不显示原始值。转换模型沿用这一拒绝规则，不静默替换成零或继承值。
  [360 次官方编辑及 270 次错误观察](benchmarks/paragraph-dimensions-26.9.json)覆盖 45 个普通输入；
  本项目直接值接受 `None` 清除，与官方已解析值接口不同。字符缩进的字号上下文与加载/保存状态不同，见[原生生命周期观察](benchmarks/paragraph-character-indents-26.9.json)；当前 DOM 有效格式明确拒绝；
  逻辑字符缩进、相对行间距、编号/条件表格上下文和完整 ParagraphFormat API 仍未对齐。
  `start/left`、`end/right` 及 `firstLine/hanging` 点值在同层冲突时按官方观察取 XML 中后出现的属性，
  这属于固定版本行为，不将属性顺序敏感解释为格式标准要求。不同层按就近属性继承；修改或清除 left/right
  同时清除对应逻辑别名及字符别名，保留另一侧缩进和其他属性。
  [168 次缩进别名编辑](benchmarks/paragraph-logical-indents-26.9.json)覆盖 42 个继承/别名输入，
  包含 bidi 开/关 getter、保存重开及排版模型映射；不是双向段落排版验收。
  三个 `character_unit_*_indent` 属性现可只读访问字符值：段落直接值及
  `style.direct_paragraph_format` 未设置时返回 `None`，显式零返回 `0.0`；
  `style.paragraph_format` 按 docDefaults/basedOn 链逐属性继承，没有任何设置时返回 `0.0`。
  支持 `leftChars/startChars`、`rightChars/endChars`、`firstLineChars/hangingChars`，
  同层后出现的别名生效，悬挂字符值为负。重复属性组和损坏数值会报错，异常隐藏原始值。
  数值语法使用 [XML Schema 的整数形式](https://www.w3.org/TR/xmlschema-2/#integer)，
  拒绝 Python 专用的下划线数字、非 ASCII 数字和指数形式；转换模型采用相同检查。
  [24 组官方继承读取](benchmarks/paragraph-character-reads-26.9.json)及 205 个官方保存输出
  已用于读取验证；这是原包 DOM 的读取阶段，setter 尚未提供。
  字符值到点值仍需字体解析，段落有效格式读取明确拒绝，不以 0 冒充；
  编号/条件表格继承、复杂文字和渲染仍未验收。
  `to_light_document()` 及 `Document` 的 DOCX/Flat OPC 加载会汇总为
  `load.character_indents_ignored` 损失诊断；将 `ContentLossWarning` 升级为错误可拒绝快照转换。
  原包保存不触发转换诊断，也不移除原属性。
  [32 次间距边界编辑](benchmarks/paragraph-spacing-limits-26.9.json)覆盖段落/样式、合法边界、
  量化和保存重开；[极端 setter 观察](benchmarks/paragraph-dimension-extremes-26.9.json)发现官方缩进
  会夹紧大值，NaN/Infinity 可产生负边界哨兵。本项目仍拒绝非有限值，错误行为差异未解决；
  极端观察仅检查 live getter，不能作为保存或渲染验收。
  [54 次有限缩进边界编辑](benchmarks/paragraph-indent-limits-26.9.json)覆盖夹紧、浮点溢出、
  样式/段落编辑和保存重开。官方 10 次负首行边界输出含负数 `hanging`；本项目写合法正数
  `2147483648`（[Open XML 属性类型](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.indentation.hanging?view=openxml-3.0.1)），经独立解析和官方加载对照，解析值相同。原生异常属性及规范化对照分别保留，
  不声称文件字节相同，也未验收这些极端值的最终渲染。
  [10 个原生 PDF 对照](benchmarks/paragraph-dimensions-rendering-26.9.json)覆盖 0/12.375pt 样式编辑，
  原点/字宽容差 0.02pt、144dpi 黑色像素差上限 1%，实测最大约 0.412%。
  仅覆盖普通字体 ASCII、固定行距和非首页面，负缩进及复杂内容的最终渲染仍未验收。
- 读取不改变部件。样式句柄按 ID 读取当前相关部件，部件替换后不会继续编辑旧树。
  编辑启用保存投影，保留内存格式和未涉及的属性；保存重开后的 Run 格式可能按官方
  受测行为变化。重复直接属性在写入前拒绝，保存失败仍保留既有目标文件。

### 有效格式读取

```python
effective = run.effective_font
print(effective.bold, effective.italic, effective.size)
print(paragraph.effective_paragraph_format.alignment)
print(paragraph.effective_paragraph_format.widow_control)
```

- 返回不可变的值快照；重新访问属性会重新解析当前 XML，不缓存可能过期的结果。
- 通过主文档的 styles relationship 定位实际部件，支持相对路径、绝对包路径和 URI 转义；
  不把未被关系引用的 `word/styles.xml` 当成有效样式表。缺失/外部/非法关系明确报错。
- 字体按 `docDefaults → 默认或显式段落样式 basedOn 链 → 显式字符样式 basedOn 链 → 直接格式`
  解析；未引用的默认字符样式不应用到 Run。字体布尔属性在每条链内取最近的显式值；两类样式均
  定义时，其异或结果与文档默认开启值合并，只有一类定义时使用该值。直接格式明确设置开/关。
  此规则按官方 26.9.0 的固定语料校准，不能推广为所有 Word 版本、表格或复杂文字的行为。
- [字体布尔编辑观测](benchmarks/dom-font-booleans-26.9.json)复用 891 个自有输入，新增 11 个
  属性的 Run／段落样式／字符样式赋值共 5,346 组，当前 live getter 与官方完全一致。
  保存成 DOCX／Flat OPC 后，10,692 组官方冷读取无 getter 差异；原始输出、独立 XML
  声明复验及官方观测保存在同一压缩归档。该语料不覆盖 basedOn 链、表格/列表字体、
  跨文档样式翻译及渲染；`hidden` getter 尤其不能直接作为最终文字可见性判断。
- 段落对齐按 `docDefaults → 段落样式链 → 直接格式` 解析；段落标记字体不错误地应用到文字 Run。
- 样式循环、缺失父样式、跨类型继承、重复 ID 和非法已支持属性值不会静默忽略。
- 当前解析 `w:b`、`w:i`、`w:sz`、`w:jc`、上述分页标志及五个点数属性，不是完整字体或排版解析器。
  未定义对齐时返回 `left`；未定义粗斜体则返回 `False`。普通字号的隐式默认值按
  官方 26.9.0 固定观察解析为 10pt/11pt，不作为复杂文字脚本选择或最终字形尺寸。
- 编号段落、有显式或默认表格样式的单元格，以及样式定义内嵌套 `rStyle` 暂不提供有效值，
  抛出 `NotImplementedError`。直接格式修改仍可用。主题字体/颜色、条件表格样式和复杂文字的
  `bCs/iCs/szCs`、自动脚本选择尚未实现，不能把上述值理解成每个字形的最终显示效果。

## 显式文本范围与局部格式

```python
paragraph = editable.body.paragraphs[0]
selection = paragraph.range(1, 4)  # Python Unicode 码点，[start, end)，不含 end
print(selection.text)
selection.apply_font(bold=True, italic=False, size=13.5)
selection.replace("替换内容")
print(selection.start, selection.end, selection.text)

# 插入：段落中间沿用右侧 Run 的格式，段落末尾沿用最后一个有文本 Run 的格式。
paragraph.range(0, 0).replace("前缀")
```

- `range()` 默认选中整个普通段落；只接受整数边界，不接受 bool，也不跨段落。
  不按 UTF-16 单元或字素簇计算；调用方需避免主动切开组合字符。
- `replace()` 支持替换、删除（空字符串）和折叠范围插入，沿用匹配起点的格式。
  空段落可插入文本；原文本节点和其属性/注释保留，不自动合并 Run。
- `apply_font()` 支持 `size` 和上述 13 个字体布尔属性；不传某项则不改它，传 `None` 则移除该直接设置。
  自动拆分首尾边界 Run，只修改选中片段，其余格式保留。空范围不支持字体设置。
- 参数、整段普通文本结构及所选 Run 的可安全拆分性均在修改前检查。
  未支持的 XML 元数据（包括唯一 `xml:id`）不被拆分或复制；未知内容不被压平成纯文本。
- 范围记录创建时的段落 XML。其他操作改变这个段落的文字、结构或格式后，旧范围会报
  `ValueError`，要求重新选择；不会悄悄使用旧偏移。范围自身操作会更新自己的状态，其他段落的修改不使它失效。
- 复杂内联容器、超链接、字段、书签、修订、Tab、换行仍不进入可编辑范围。
  范围中的文本与普通替换一样包含隐藏 Run，不做可见性过滤。

也可直接使用 `run.split(offset)` 拆分普通直属 Run：返回右侧 Run；偏移为 0 时返回原 Run，
偏移为 Run 长度时返回 `None`，两者均不修改文档。只支持段落内、无复杂范围和未知元数据的 Run。
拆分保留属性、文本节点及注释；新增右侧 Run 继承原 Run 的格式。

## 顺序、节点归属与结构编辑

```python
paragraph = editable.create_paragraph("新增段落")
editable.body.append_child(paragraph)
run = editable.create_run("补充文字")
paragraph.append_child(run)

copied = paragraph.clone()
editable.body.insert_before(copied, paragraph)
copied.remove()

plain_table = editable.create_table(2, 3)
editable.body.append_child(plain_table)
cell = plain_table.rows[0].cells[0]
cell.paragraphs[0].append_child(editable.create_run("单元格内容"))
cell.append_child(editable.create_table(1, 1))  # 插在单元格最后一个段落前

for node in cell.child_nodes:
    print(node.node_type, node.part_name)

runs = editable.get_child_nodes(aw.NodeType.RUN, deep=True)
```

- `child_nodes` 是有序 tuple；`paragraphs`、`runs`、`tables`、`rows`、`cells` 是过滤视图。
  单元格的段落与嵌套表格不再分开存储，原始交错顺序得以保留。
- 节点提供 `parent_node`、`owner_document`、`part_name`、`clone()`、`remove()`；容器提供
  `append_child()`、`insert_before()`。插入已有节点会移动它，不会复制。
- `clone(deep=True)` 保留无参数深复制行为；`clone(False)` 清空段落、表格、行、单元格的
  内容子节点，保留直接格式及表格网格。Run 是内容叶节点，浅复制仍保留文本和字体格式。
  复制节点无父节点、仍归属原文档，修改副本不会修改来源。`tblGrid` 是表格格式元数据，
  可通过 `xml` 检查，不再作为 `child_nodes` 中的未知内容节点。`deep` 仅接受 bool。
- 单独删除节点后保存也会更新原 part；删除节点仍归属原文档，但不再有父节点。
  [商业基准与保存回读记录](benchmarks/commercial-26.9-dom-structure.json)保留了删除未持久化
  问题的修复前后证据。当前 `remove()` 返回自身、无父节点时抛出 `ValueError`，与商业基准的
  返回值/异常仍有差异；复杂内容的浅克隆与跨文档导入也尚未对齐。
- 检查合法父子类型、循环、跨文档/part 插入和单元格末尾段落。
  正文新增内容放在最后的 `sectPr` 前面，不重新解释分节。
- 只允许移动、删除、复制受支持的简单结构。含未知元素、复杂范围或关系属性的结构明确拒绝。
  目标或来源 story 包含书签、批注范围、字段、权限范围或修订时，结构操作保守拒绝。
- 支持整张普通表格的创建、插入、移动和删除，以及单元格内的段落编辑。
  现有表格仅开放下述受限水平合并；行/列增删、垂直合并和拆分暂不支持。
- 创建但未插入的节点不会使原包 part 被重新序列化。

`destination.import_node(source_node, True)` 返回归属目标、无父节点的副本，来源保持不变。
当前跨文档路径仅实现 USE_DESTINATION_STYLES：按样式名称/类型匹配目标样式，新增简单
样式时处理 ID 冲突；依赖校验完成后才提交 styles part。相同文档的普通节点支持三种模式
及两种复制深度，页眉普通节点可导入正文。`ImportFormatMode` 的三个枚举值与固定基准
一致，但枚举存在不表示跨文档 KEEP_SOURCE_FORMATTING / KEEP_DIFFERENT_STYLES 已实现。

新增段落/字符样式可迁移受测的粗体、斜体、显式字号及普通段落对齐，解析完整 basedOn
链和本次计划新增的依赖；段落样式去除与目标基样式相同的冗余属性，字符样式保留显式
继承值；默认开启的字符粗斜体按观测的导入语义转换，getter 本身不作该转换。
没有显式字符样式属性时仍按目标继承，不能笼统理解为始终保留源有效格式。

跨文档导入成功后，保存会在独立 XML 副本中处理正文及内部关联的页眉/页脚：当 Run
实际引用字符样式时，省略其使用的根段落/字符样式中与 docDefaults 相同的粗斜体属性。
未使用的样式保持原有定义；内存样式及 Run 的有效格式不因保存而改变。重新打开后的
Run 格式可能不同于保存前，与受测的官方行为一致。styles part 的外部 XML 注释和
处理指令亦保留；不支持的属性元数据在提交前拒绝，保存失败保留已有目标文件。

原 247 组、132 组段落默认值、108 组字符默认值现均可导入。新增 36 组使用/未使用
上下文包含正文 20 组、页眉和页脚各 8 组。[全部 523 组记录](benchmarks/style-import-projections.json)
保留安装包原始输出及官方回读；受测的内存/冷段落 getter 与冷 Style 粗斜体/字号匹配。
这不是完整公开 Font/Style API、格式或渲染验收；原有仅导入段落的检查也不能代替
既有目标内容的验证。

编号、图片/复杂关系、不同默认值或主题、段落标记的字符样式、隐式应用字号
需要覆盖目标字号、其他冲突属性或条件表格样式仍报 NotImplementedError，未视为范围
豁免。带未知元数据的待迁移属性也拒绝修改；拒绝发生在提交前，保留源与目标包。
冲突检查遍历完整祖先链，含多个条件表格样式区段；直接基样式 XML 相同不能证明
继承格式相同。不能用这个入口或 getter 结果宣称完整跨文档导入和渲染正确。
修改 part 后保存会省略仅含格式属性/网格的零行普通表格，与已观测的官方浅导入保存结果
一致；内存 DOM 不删除它。带注释、非空文本、修订或不受支持元数据的表格仍保留。

`get_child_nodes()` 默认遍历正文 story 的直属节点；`deep=True` 递归遍历，也能读到未知容器内的已知节点。
这些节点不一定允许修改，例如内容控件和修订容器下的 Run 只读。
`UnknownNode.xml` 可查看未建模内容；未知内容不会因遍历或保存而丢弃。

## 图片、超链接与水平合并

```python
paragraph = editable.body.paragraphs[0]
link = paragraph.add_hyperlink("网站", "https://example.com/report")
link.runs[0].font.bold = True
link.target = "mailto:team@example.com"  # 或 "#bookmark"
picture_run = paragraph.add_picture("logo.png", width=72, alternative_text="公司标识")

# 逻辑网格列范围 [0, 2)，保留各源单元格的段落/子表格顺序。
# merged = editable.body.tables[0].merge_cells(0, 0, 2)
```

- `add_hyperlink(text, target)` 返回 `dom.Hyperlink`，可读取 `text`、`runs` 和读写 `target`。
  写入仅允许非空 bookmark、HTTP(S)、mailto；不联网验证目的地址。更改共享链接中的一个目标
  不会篡改其他链接；旧的未使用关系保留，不做可能误删未知引用的资源清理。
- `add_picture(source, *, width=None, height=None, alternative_text="")` 接受 PNG/JPEG 的 bytes、
  二进制流或路径。尺寸以点为单位；只给一边时保持宽高比，未指定时按 96 DPI。
  像素/数据/包大小检查沿用现有安全限制，媒体按字节去重，更新 Content Types 和当前 part 的 relationships。
- relationship ID 在当前 part 内分配；drawing ID 扫描包内声明的 XML 部件，避免冲突。
  新资源插入前先验证和构建；验证失败不会改变文档。返回的图片 Run 不支持通过 `text` 设置删除图片。
- `Table.merge_cells(row_index, start_column, end_column)` 支持简单、完整网格的水平合并；
  已有 gridSpan 必须完整落在选区内，禁止切开已有单元格。首个单元格保留自身格式，
  旧单格首选宽度移除，由网格与新跨度决定宽度；内容按原顺序移动，不重复 ID。
  vMerge、旧式 hMerge、省略网格单元格、复杂范围/未知结构明确拒绝。
- 图片/链接可加到正文及可编辑页眉页脚。含字段、修订、书签范围等不安全 story 的插入仍拒绝；
  不开放带关系节点的任意复制、移动或跨 part/跨文档导入；普通节点的初步导入见上文。

## 页眉页脚与转换兼容

```python
# 按原包 part 名称访问，不推断该页眉属于哪一节或首页/奇偶页。
# header = editable.story("word/header1.xml")
# header.paragraphs[0].runs[0].text = "新的页眉"
# header.append_child(editable.create_paragraph("追加", part_name="word/header1.xml"))

print(editable.part_names)
styles_xml = editable.part_xml("word/styles.xml")

ldm_snapshot = editable.to_light_document()
```

页眉页脚支持与正文相同的基础操作和新增图片/链接，但不允许跨 part 移动或导入已有节点。
`part_xml()` 返回只读 XML 快照，其他部件（如批注、脚注、编号、主题）可检查，但没有语义编辑接口。
当前要求主文档是 `word/document.xml`；样式查询按它的内部 styles relationship 解析实际部件。
页眉页脚使用同一份文档样式表；图片/链接关系在自身 part 内管理，跨 part/跨文档导入仍未实现。

`to_light_document()` 调用现有 reader，从当前 DOM 生成独立快照。字符单位缩进保留在快照的
`character_unit_*_indent` 字段中；正文、表格、默认页眉页脚、样式和编号的受测路径可写回
DOCX/Flat OPC，尚不用于 PDF 点值和字号换算。快照可用于现有 Markdown/PDF writer；
**修改快照不会反写 DOM**，并且旧 reader 的内容损失警告和限制仍然适用。
也可以将 `editable.to_bytes()` 交给 `aw.Document(BytesIO(...))` 继续现有转换。

现有 `aw.Document`、`Document.light_document_model`、`Document.save()` 以及
内容顺序、提取和结构化输出契约见 [项目 README](../README.zh-CN.md#结构化内容与诊断)。
不能把转换路径的保存当作原包保真保存。

## 保存契约与安全边界

- 接受 DOCX 路径或二进制流，共用原有输入大小、ZIP 扩展大小、路径、重复条目和加密条目检查。
- XML 使用 `defusedxml.minidom`；DTD、实体和外部实体不被允许。
  XML 按需解析；未打开的未知 XML 部件仅作为原始字节保留，不宣称全部已验证。
- minidom 保留前缀、命名空间声明、未知属性、注释和处理指令，避免重写后
  `mc:Ignorable` 等前缀值失去绑定。未知元素与所有未修改部件保留。
- 没有修改时，每个 part 的解压后内容保持字节一致。修改后仅重新序列化被修改的 XML part，
  **该 part 不承诺原词法字节一致**（如引号、空元素形式和 XML 声明可能变化）。
- ZIP 会重新打包；不承诺整个文件字节一致。原 ZIP comment、条目时间和基本元数据保留。
- 实际输出再次检查部件大小、解压后总大小、ZIP 条目数和整个 ZIP 大小，沿用本项目的
  加载安全限制；编辑和样式投影不能绕过这些检查。恰好达到限制可保存并重新加载，
  超限抛出 `ValueError`，不返回无法被本项目加载的 ZIP。这些是本项目安全边界，
  不代表商业基准的文档容量或许可限制。
- `save()` 使用原子输出，允许源和目标相同；失败时不覆盖既有目标。
- 数字签名包和 Strict OOXML 主文档拒绝加载，避免无效签名或伪装格式支持。
- 宏、嵌入对象和外部关系照原样保留，不执行、不联网，**不是安全清洗或脱敏器**。
- 同一个可变文档不可并发编辑。关系仅随受支持的图片/链接操作更新；主题、编号、域结果和排版缓存不更新。

后续范围、依赖与验收条件统一见 [商业对齐计划](commercial-alignment-plan.md)。

测试位于 `tests/test_docx_dom.py`、`tests/test_docx_dom_ranges.py`、`tests/test_docx_dom_styles.py`、
`tests/test_docx_dom_resources.py`、`tests/test_content_integrity.py` 和 `tests/test_dom_libreoffice_integration.py`；
保真验证以 part payload、XML 内容和修改约束为准，
不替代 Microsoft Word 实际打开、修复提示检查及复杂业务文档的兼容性验证。

## 包内容保留报告

`doc.preservation_report()` 返回 `added`、`modified`、`removed`、`unchanged` 四组排序后的部件名。
它按解压后字节的 SHA-256 与最初加载的包比较，包含资源和关系部件；不以 ZIP 压缩字节
或 XML 语义等价判定。只读取节点不应产生修改。保存不会重置基线，重新打开保存结果才建立新基线。
该报告描述部件变化，不证明 Word 视觉保真或所有 OOXML 语义均受支持。

## 新建文档与默认页面尺寸

`Document()` 创建含一个空段落的新节。新建页面、TXT 和 Markdown 文档默认使用
Letter（612×792pt），JSON、DOCX/Flat OPC 往返和 PDF 页框保留尺寸；显式 A4 尺寸保持原值。
默认尺寸变化可能改变换行与分页。这不表示完整 PageSetup、PaperSize 公共枚举、
构造重载、全部 DOM 创建操作或多节排版已对齐。

## 缺失页面尺寸的加载默认值

DOCX/Flat OPC 加载及 DOM 转换快照中，未设置的页面宽度和高度分别使用 Letter 的
612pt 和 792pt。每个节独立应用这些默认值，不继承前一节的尺寸；显式尺寸和
`orient` 保留。这修正了缺失尺寸文档此前按 A4 输出 PDF 的行为，可能改变分页。
尺寸已验证读取、往返保存和 PDF 页框，不代表边距、完整 PageSetup、PaperSize 枚举值
或多节分页均已对齐。


## 页面边距的默认值与显式零值

转换模型的新建页面及 DOCX/Flat OPC 未设置的边距使用四边 70.85pt、
页眉页脚距离 35.4pt、gutter 0pt。部分设置只覆盖对应字段，各节独立应用默认值。
加载中的空或非法边距数字按固定基准观测规范化为零。JSON 和保存重开保留
显式零值及负值，保存时写出完整的七项 `pgMar` 属性；零边距不会因整组为零而丢失。

PDF 正文投影使用显式零边距，不再将其替换为 20mm。默认值变化可能影响
正文宽度、换行和分页。当前检查包括默认与零左边距的实际文字起点变化，
不代表整页排版等价；负边距渲染、页眉页脚零距离定位、gutter 布局及多节边距
尚未通过完整验收。

The lightweight model retains anchored DrawingML rectangle `solidFill`, `gradFill`,
`noFill`, and `fillRef` declarations, including colour transforms, rotation and
flip flags, through JSON, DOCX and Flat OPC saving. Direct fills do not require a
style reference. Assigning `Shape.fill_color` replaces the retained declaration.
This preserves source declarations; the PDF renderer currently warns that their
appearance is not resolved. Other DrawingML geometry and effects are not covered
by this declaration path.
## Font name channels

`run.font` exposes direct `name`, `name_ascii`, `name_other`, `name_bi`, and
`name_far_east` declarations. An unset direct channel returns `None`; an explicit
empty literal returns `""`. `run.effective_font` resolves document defaults,
paragraph and character style inheritance, then direct formatting. Style fonts
resolve their own `basedOn` chains and document defaults.

Assigning `name` sets all four OOXML font channels. Assigning one channel preserves
the others and removes only that channel's theme reference. Names require nonempty
strings; invalid values and duplicate property groups fail before mutation.
UTF-16 surrogate pairs are joined and isolated surrogates become U+FFFD. XML-invalid
characters may be omitted during serialization without mutating the live value.

Whole-table font name inheritance and numbered paragraph body fonts are supported
when their references resolve. Conditional or nested styled tables, numbering
overrides, label fonts, paragraph layout, and missing-theme save lifecycle remain
outside this stage. These name operations do not establish full Font API or
rendering equivalence.
