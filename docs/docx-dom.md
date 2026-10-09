# 原包保真的 DOCX DOM（资源关系与受限水平合并）

新增入口 `aspose.words_foss.DocxDocument`，与现有 `Document → LDM → writer` 转换路径并行。
它以原始 OOXML 为唯一权威状态，节点是 XML 的类型化视图，不维护另一份同步模型。
**这是可逐步扩展的基础 DOM，不是完整 Aspose.Words DOM，也不是 Word 排版引擎。**

`DocxDocument` 也可从路径或二进制流加载 Flat OPC XML 包。四种文档、宏文档、模板、宏模板
变体均还原 XML 与二进制部件，保留祖先命名空间及主部件 Content Type；沿用原包 DOM 的
编辑边界。`save()` / `to_bytes()` 输出 ZIP OOXML 包，需使用与主部件类型对应的
`.docx` / `.docm` / `.dotx` / `.dotm` 扩展名；不保留输入 XML 的字面序列化。
`to_flat_opc()` 返回 UTF-8 XML 字节，`save_flat_opc(path)` 原子保存 Flat OPC XML；
两者保留部件类型、XML 节点和二进制资源，但 XML 声明、编码及序列化会规范化。
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
现在支持粗体、斜体、字号、段落对齐和段落/字符样式 ID。`run.font.style_id` 只接受已存在的
character style，`paragraph.paragraph_format.style_id` 只接受 paragraph style；两者支持 `None` 移除引用。
不把样式继承结果物化为直接格式，也不自动创建样式。

### 有效格式读取

```python
effective = run.effective_font
print(effective.bold, effective.italic, effective.size)
print(paragraph.effective_paragraph_format.alignment)
```

- 返回不可变的值快照；重新访问属性会重新解析当前 XML，不缓存可能过期的结果。
- 通过主文档的 styles relationship 定位实际部件，支持相对路径、绝对包路径和 URI 转义；
  不把未被关系引用的 `word/styles.xml` 当成有效样式表。缺失/外部/非法关系明确报错。
- 字体按 `docDefaults → 默认或显式段落样式 basedOn 链 → 显式字符样式 basedOn 链 → 直接格式`
  解析；未引用的默认字符样式不应用到 Run。粗斜体在每条链内取最近的显式值；两类样式均
  定义时，其异或结果与文档默认开启值合并，只有一类定义时使用该值。直接格式明确设置开/关。
  此规则按官方 26.9.0 的固定语料校准，不能推广为所有 Word 版本、表格或复杂文字的行为。
- 段落对齐按 `docDefaults → 段落样式链 → 直接格式` 解析；段落标记字体不错误地应用到文字 Run。
- 样式循环、缺失父样式、跨类型继承、重复 ID 和非法已支持属性值不会静默忽略。
- 当前只解析 `w:b`、`w:i`、`w:sz` 与 `w:jc`，不是完整字体或排版解析器。
  缺少字号或对齐定义时返回 `None`，不猜测 Word 的应用默认值；未定义粗斜体则返回 `False`。
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
- `apply_font()` 支持 `bold`、`italic`、`size`；不传某项则不改它，传 `None` 则移除该直接设置。
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

编号、图片/复杂关系、不同默认值或主题的新样式、依赖冲突基样式的有效格式转换仍报
NotImplementedError，未视为范围豁免。不能用这个初步入口宣称完整跨文档导入。
新样式的冲突检查遍历完整 basedOn 祖先链，含多个条件表格样式区段；直接基样式 XML
相同不能证明继承格式相同。发现差异或复杂依赖时在提交前拒绝，保留源与目标包。
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

`to_light_document()` 调用现有 reader，从当前 DOM 生成独立快照。快照可用于现有 Markdown/PDF writer；
**修改快照不会反写 DOM**，并且旧 reader 的内容损失警告和限制仍然适用。
也可以将 `editable.to_bytes()` 交给 `aw.Document(BytesIO(...))` 继续现有转换。

现有 `aw.Document`、`Document.light_document_model`、`Document.save()` 以及
`docx_edit.replace_text()` 的调用方式保持兼容；内容顺序/提取改进见 [升级说明](upgrade-notes.md)。
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
- `save()` 使用原子输出，允许源和目标相同；失败时不覆盖既有目标。
- 数字签名包和 Strict OOXML 主文档拒绝加载，避免无效签名或伪装格式支持。
- 宏、嵌入对象和外部关系照原样保留，不执行、不联网，**不是安全清洗或脱敏器**。
- 同一个可变文档不可并发编辑。关系仅随受支持的图片/链接操作更新；主题、编号、域结果和排版缓存不更新。

## 后续实施顺序

1. 基础文本范围与有效格式已实现；后续补齐主题字体/颜色、编号和表格条件样式规则。
2. 图片/超链接关系管理已实现；跨文档/part 导入仍按需求评估。
3. 简单水平网格合并已实现；后续补齐垂直合并、行列编辑及分节/页眉页脚写回语义。
4. 按业务需求增加内容控件、批注、脚注、复杂域和修订操作。
5. 页面坐标、精确分页与字段计算单独评估渲染引擎，不将基础 DOM 承诺成完整排版引擎。

测试位于 `tests/test_docx_dom.py`、`tests/test_docx_dom_ranges.py`、`tests/test_docx_dom_styles.py`、
`tests/test_docx_dom_resources.py`、`tests/test_content_integrity.py` 和 `tests/test_dom_libreoffice_integration.py`；
保真验证以 part payload、XML 内容和修改约束为准，
不替代 Microsoft Word 实际打开、修复提示检查及复杂业务文档的兼容性验证。

## 包内容保留报告

`doc.preservation_report()` 返回 `added`、`modified`、`removed`、`unchanged` 四组排序后的部件名。
它按解压后字节的 SHA-256 与最初加载的包比较，包含资源和关系部件；不以 ZIP 压缩字节
或 XML 语义等价判定。只读取节点不应产生修改。保存不会重置基线，重新打开保存结果才建立新基线。
该报告描述部件变化，不证明 Word 视觉保真或所有 OOXML 语义均受支持。
