# 26.7.0.post2：可信转换与结构化输出

按“先可信，再好用，最后按需求扩展”的优先级实施。仍然是轻量文档模型驱动的转换库，
不是完整 Word 排版引擎、全功能 DOCX 编辑器或文档脱敏工具。已有路径输出保持原子替换；
没有静默切换到 LibreOffice。

## 1. 内容边界与诊断

- 内置 RTF reader 只接受 OLE2/DOC-backed 文件。标准文本 RTF（`{\\rtf...`）明确报错，
  提示使用原文件的 LibreOffice PDF 路径，不再把文本 RTF 当成 OLE2 解析。
- DOCX 加载检查脚注/尾注、评论引用、对象、altChunk、修订、简单字段、内容控件、公式。
  这些内容不能完整进入 LDM；简单字段的结果可读，不代表字段指令被保留。
- 不同章节、首页、奇偶页的页眉页脚不能完整保留；加载和写出都会对已知降级发出警告。
  默认页眉页脚及受支持的复杂字段标记可以写回，不能据此承诺所有 OOXML 保真。
- DOCX、Markdown、PDF 写出对未支持的模型节点补充诊断。PDF 对缺字、字体替换、
  不支持的选项、布局降级给出稳定分类。

```python
from dataclasses import asdict
import warnings
import aspose.words_foss as aw

with warnings.catch_warnings():
    warnings.simplefilter("error", aw.ContentLossWarning)
    doc = aw.Document("input.docx")
    doc.save("output.docx")

report = [asdict(item) for item in doc.diagnostics]
# code / severity / location / message
```

`Document.diagnostics` 累积加载、保存和内存转换的已知诊断，即使 Python warning 被忽略或升级为错误，
已有诊断仍被记录。列表不会在下一次转换前自动清空，重复保存可能追加重复记录；
需要按次统计时可在操作前调用 `doc.diagnostics.clear()`。构造函数因加载警告或错误失败时，
调用方无法通过尚未返回的 `doc` 读取诊断，应处理抛出的异常。
并发的诊断收集上下文隔离；这**不代表**同一个可变 Document/PDF writer 可并发使用。
诊断不是所有内容损失的枚举器；`--strict` 只将已知损失和缺字作为错误，普通字体替换不致失败。

修复了 Python 3.10 的另一个边界：其 Unicode 数据库可能不认识较新的 emoji，
但这些字符及私用区字符仍需字体覆盖，不能因为 `isprintable()` 为假就跳过缺字检查。

### 本次开发分支的内容完整性改进

- DOCX 按实际关系定位 styles/settings/numbering/theme/images 以及页眉页脚、脚注、尾注，
  支持非惯常路径与 URI 转义；同一 reader 重复加载不会泄漏前一文档的状态。
- `Document.to_dict()` 增加 `source_stories`：原始 part、note ID、有序 blocks，以及
  页眉页脚的 section/variant/inherited 引用。段落增加 `note_references`；
  单元格增加有序 `blocks`，原有字段保留，`schema_version` 仍为 1（加法扩展）。
- 提取注释内容**不等于渲染或完整写回**。DOCX/PDF/Markdown/TXT 输出均报告
  `*.notes_omitted`；页眉页脚输出仍有旧的聚合/变体限制。
- `get_text()` / TXT 包含正文嵌套表格的按序可见文本；过滤字段指令/隐藏 Run，
  展示链接标签而非内部 Markdown 编码，不附加 source stories。
- 标签括号与 URL 括号的编解码由转换路径共用；已有 DOCX 链接经过 Markdown/DOCX/PDF 时不截断目的地址。
  Markdown 的合并网格展开、嵌套表格展平会给出明确损失诊断。
- 原包 DOM 增加图片/超链接资源操作和受限水平合并，详见 [DOM 指南](docx-dom.md)。

## 2. 内存输出

```python
from io import BytesIO
import aspose.words_foss as aw

doc = aw.Document(BytesIO("中文报告".encode("utf-8")))
pdf_bytes = doc.to_bytes(aw.SaveFormat.PDF)
docx_bytes = doc.to_bytes("docx")
markdown_bytes = doc.to_bytes(aw.saving.MarkdownSaveOptions())
text_bytes = doc.to_bytes("text")
```

`to_bytes(format_or_options)` 必须显式指定 DOCX、PDF、Markdown 或 text；
PDF/DOCX 使用与文件输出相同的 writer，Markdown 保留编码设置。
它不会创建输出文件，不支持外置 Markdown 图片目录；需要多个文件时仍用 `save(path, options)`。
文件保存仍为原子发布，不必为了服务端响应先创建临时文件。

## 3. PDF 表格

- 在网格内保留逐 run 的混合字号、粗斜体、颜色、下划线、突出显示及超链接。
- 单元格图片和嵌套表格不再导致整表降级为顺序段落。
- 使用实际测量的行高；连续的前导表头在换页后重复，能完整容纳的行尽量整体移动，
  超高行在内容行边界拆分。表头无法容纳或旋转内容过高时明确失败。
- 支持横向 gridSpan / hMerge；vMerge 延续行的内部边线不重复绘制。
  DOCX 读写也保留 gridSpan，并写出匹配的 tblGrid。跨度/网格列数限制为 1,024。
- 修复剪裁区域退出后 fpdf2 字体缓存与实际 PDF 状态不一致的问题，
  防止邻近单元格或表后的文字消失。
- 精确行高、不能满足的最小行高、必须拆分的 cantSplit、单元格页/分节符产生已知损失警告。
  垂直合并仍是边线和网格层面的近似，不是 Word 的完整合并单元格排版。
  嵌套表格按整体布局；过高而不能拆分的内容会失败，不会偷偷裁掉再宣称成功。

保留标准旋转文本支持，其他未知方向明确拒绝。单元格新增 `children` 有序视图和序列化的
`content_order`，DOCX/PDF/结构化输出保留段落与嵌套表格的交错顺序；旧 `paragraphs`/`tables`
构造保持兼容。Markdown 按序展平嵌套表格并报警。图文绕排、分页和页眉页脚仍存在 Word 兼容性边界。

## 4. 保留原 OOXML 包的定点替换

```python
from aspose.words_foss.docx_edit import replace_text

count = replace_text("original.docx", "edited.docx", {"客户公司": "新的客户公司"})
```

这是与 `Document → LDM → 重建 DOCX` **分开的窄接口**：

- 同时进行字面替换，匹配范围是每个 Word `w:t` 文本节点，覆盖 Word XML 部件内的正文、
  页眉页脚、脚注等。不跨 run 合并匹配，不做正则替换、格式编辑或分页重排。
- 未修改的 XML 字节和其他部件 payload 保持原样，ZIP 容器重新打包，并非整个文件逐字节相同。
- 返回实际替换的匹配总数。所有替换 key 都必须至少出现一次；缺失或只跨 run 的匹配失败。
  重叠 key 优先匹配较长字面值，替换后的文本不会再次参与本次匹配。
  源和目标可以相同，写出失败不会覆盖旧文件。
- ZIP 大小/路径/重复条目等检查与现有入口共用，XML 禁止 DTD；
  当前仅接受 UTF-8 XML，拒绝数字签名包及不适合文本节点的替换字符（包括换行和 Tab）。

**不是 sanitizer**：宏、外部关系和未知部件会保留。只在适当隔离环境中处理文件；
不能将保留原包误当成安全清洗。新增的 `aw.DocxDocument` 提供 XML 绑定的基础 DOM，支持
同段落普通 Run 之间的替换、基础直接格式和简单结构修改，同时保留未知 XML 和原始部件。
后续升级增加 `Paragraph.range()`、局部字体格式与边界 Run 拆分，并按文档默认值和段落/字符样式链
解析粗斜体、字号及段落对齐。未支持的编号/表格样式上下文明确拒绝，不猜测完整有效格式。
详见 [DOCX DOM 使用与边界](docx-dom.md)。它不改变旧 `Document.save()` 的重建语义，
也不是全功能 Word 编辑器；复杂修改仍需专门的 OOXML 工具，原文档渲染仍需 LibreOffice/商业引擎。

## 5. 可选多语言塑形

```bash
# 从本 fork 安装；同名 PyPI 包不作为本仓库的发布来源
python -m pip install "aspose-words-foss-enhanced[shaping] @ git+https://github.com/shenyankm/Aspose.Words-FOSS-for-Python.git@dev"
# 已有本地 checkout 时也可使用：python -m pip install -e ".[shaping]"
python -m aspose.words_foss.convert input.docx output.pdf \
  --text-shaping --fallback-font /trusted/ArabicFont.ttf
```

```python
options = aw.saving.PdfSaveOptions()
options.text_shaping = True
options.fallback_fonts = ["/trusted/ArabicFont.ttf"]
pdf_bytes = doc.to_bytes(options)
```

复用 fpdf2/HarfBuzz 的 shaping 和 bidi 处理，实际渲染与行宽/行高测量一起开启，
避免只换 glyph、不换测量的排版错误。没有开启 shaping 的复杂文字会提示已知降级。
fallback 字体需要可信、覆盖目标语言且有合适的 OpenType 字形信息；
仅安装 HarfBuzz 不会提供字体，也不能保证所有语言或所有混合文本的 Word 相同布局。
这一实现使用 fpdf2 的内部 bidi/断行 API，升级 fpdf2 后须重跑塑形/测量回归。

## 6. 结构化内容与模板组合

```python
import json
payload = doc.to_dict()
json_text = json.dumps(payload, ensure_ascii=False)
```

返回 schema_version、source、正文顺序 blocks、页眉页脚和已有 diagnostics；
段落包含文本、标题级别、列表、样式及带链接/混合格式的 runs；
表格包含行、单元格、跨度/合并标记、段落与子表格；图片导出 alt/尺寸而不是二进制。

位置是 `sections[0].body.children[2]` 等**模型路径，不是 PDF 坐标/页码**。
字段指令和隐藏 runs 不进入段落文本；文本框、脚注/评论等并未成为完整可搜索对象。
导出可能含原始路径和文档文本，发布给其他系统前自行过滤。它不是通用脱敏器或无损文档备份，
也不提供 embedding、向量库、OCR 或 chunk 策略。

不新增模板引擎。可安装 `docxtpl` 并运行 `ApiExamples/template_report.py`，
用 Jinja StrictUndefined 与 autoescape 生成 DOCX，再用本库转 PDF：

```bash
python -m pip install docxtpl
python ApiExamples/template_report.py template.docx context.json report.docx report.pdf
```

模板必须可信；Jinja 不是不可信模板的安全沙箱。模板与两个输出的路径不能相同，PDF 输出显式指定格式而不靠扩展名猜测。DOCX 与 PDF 分别原子保存，
不是两个文件一起提交的事务；PDF 转换失败时已生成的 DOCX 仍可能存在。

## 7. 工程与性能验证

- 增加 `py.typed`；发行版本提升为 `26.7.0.post2`，支持 Python 3.10–3.14（<3.15）。
- GitHub Actions：Linux / Windows / macOS × Python 3.10–3.14。
  验证源码测试/示例、构建 wheel、卸载 editable 包，并从源码目录外验证已安装 wheel。
  Windows 不支持的 POSIX 进程边界测试会显式 skip，不宣称跨平台提供相同隔离级别。
- `scripts/check_wheel.py` 校验真正从 site-packages 导入、版本、typing 标记、
  四个字体样式与 OFL 许可资源。
- `tests/test_markdown_conformance.py` 固化选定的 CommonMark/GFM 子集：嵌套列表/引用、转义、
  强调、括号链接、代码块、表格及内存往返。原 parser 在这些用例上通过，
  因此没有为此增加 Markdown 依赖；不声称完整 CommonMark/GFM 合规。Tab 缩进仍保持继承的固定 +4
  行为，而非 CommonMark 的列位置 tab-stop 规则。全面合规有需求时再迁移成熟 parser。
- `scripts/benchmark.py` 对小文档、百页正文、图文表格、重复图片、混合样式执行冷子进程基准，
  记录 import / parse / layout / serialize / write / process 时间、RSS、PDF 大小和页数。
  Windows 的 RSS 返回 null。

```bash
python -m pytest tests -q
python -m pytest ApiExamples -q --rootdir=ApiExamples -c ApiExamples/pytest.ini
python scripts/benchmark.py --repeat 3 > benchmark.json
```

局部实现已避免原 XML 批量替换的反复全量拼接，并清理 PDF 序列化后的共享字体状态，
保证重复 writer 使用不会复用被子集化的字体。没有添加未经数据证明必要的全局字体缓存、
共享 PDF 对象或常驻 LibreOffice 池。

本轮实测及基准见 [优化核验报告](optimization-report.md)，其中第 7 节单独记录后续缺字诊断优化：
只检查文本中出现的字符，不再遍历整张字体 cmap；百页合成正文的布局中位耗时下降约 15%。
这是特定环境下的实测，不是所有文档的加速保证；该后续核验未重建 wheel，且跳过了未安装依赖的塑形测试。

历史 `verification-report.md` 仍是上一版的验证记录，不替代本轮运行结果。
