# DocWeave — Aspose.Words FOSS 非官方增强版

[English](README.md) | [简体中文](README.zh-CN.md)

[![非官方增强版](https://img.shields.io/badge/status-unofficial_enhanced_fork-orange.svg)](#分支定位与上游差异) [![代码许可：MIT](https://img.shields.io/badge/code_license-MIT-blue.svg)](LICENSE) [![字体许可：OFL-1.1](https://img.shields.io/badge/font_license-OFL--1.1-blue.svg)](aspose/words_foss/pdf_writer/fonts/OFL.txt)

**无需 Microsoft Word 的多格式文档转换、中文/Unicode PDF、结构化提取及有明确边界的原包 DOCX 编辑。**

DocWeave 是独立维护的 [Aspose.Words FOSS for Python](https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python)
增强分支，不是 Aspose 官方发布版，也不替代商业 Aspose.Words。
仓库已迁移至 [shenyankm/DocWeave](https://github.com/shenyankm/DocWeave)。
开发分支为 `dev`；发行包名和导入名**尚未变更**：

| 项目 | 当前值 |
|---|---|
| 发行包 | `aspose-words-foss-enhanced` |
| Python 导入 | `aspose.words_foss` |
| 版本 | `26.7.0.post2` |
| Python | 3.10–3.14（`>=3.10,<3.15`） |
| 代码 / 内置字体许可 | MIT / OFL-1.1 |

## 导航

- [分支定位与上游差异](#分支定位与上游差异)
- [功能概览](#功能概览)
- [安装](#安装)
- [快速开始](#快速开始)
- [原包 DOCX 编辑](#原包-docx-编辑)
- [结构化内容与诊断](#结构化内容与诊断)
- [PDF 字体与可选塑形](#pdf-字体与可选塑形)
- [有界 CLI 与 LibreOffice](#有界-cli-与-libreoffice)
- [模板与其他示例](#模板与其他示例)
- [API 入口](#api-入口)
- [支持范围与限制](#支持范围与限制)
- [开发与测试](#开发与测试)
- [文档与资源](#文档与资源)
- [许可证](#许可证)

## 分支定位与上游差异

本分支起源于上游提交
[`2d2efee2787cb9e56d071d17f8d7b740dce8b784`](https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python/commit/2d2efee2787cb9e56d071d17f8d7b740dce8b784)。
在此基础上增加 Unicode PDF 字体与排版修复、内容损失诊断、资源限制、原包 DOM 编辑、
结构化 source stories 及可选 LibreOffice 转换；这不是与最新上游发布版的比较。

本项目保留原作者及原始库的归属说明。熟悉的 `Document` / `SaveFormat` 名称不意味着与上游或商业版
具备相同功能、渲染结果或行为。请以本仓库的实现、测试和文档为准；上游文档仅供背景参考。

## 功能概览

### 格式支持

| 格式 | 内置读取 | 内置写出 | 边界 |
|---|---|---|---|
| DOCX | 支持 | 支持 | 受支持的 Transitional OOXML 内容；转换会重建文档包，不是无损往返 |
| DOC | 支持 | 不支持 | 旧版 Word 97–2003 OLE2 输入；`SaveFormat.DOC` 是预留常量，不代表存在 writer |
| RTF | 有限支持 | 不支持 | 仅 OLE2/DOC-backed 文件；内置 reader 拒绝标准文本 RTF |
| Markdown | 支持 | 支持 | 选定的 CommonMark/GFM 功能，并非完整规范兼容 |
| TXT | 支持 | 支持 | 仅正文文本；输出为 UTF-8，不保留图片或原始排版 |
| PDF | 不支持 | 支持 | 使用内置字体与渲染器，不承诺 Word 一致分页 |

独立的 LibreOffice 后端可将原始 DOC/DOCX/RTF（包括标准文本 RTF）**转换为 PDF，不提供其他输出格式**。

### 选择合适的处理路径

| 需求 | 入口 | 保留内容 |
|---|---|---|
| 格式转换或 PDF 渲染 | `aw.Document` | 先进入轻量文档模型（LDM），再重建受支持的输出 |
| 提取 JSON 安全内容 | `Document.to_dict()` | 有序 blocks、run 格式、模型位置、注引用及已提取 source stories |
| 编辑 DOCX 并保留未知部件 | `aw.DocxDocument` | 原始 OOXML 包，仅执行受支持且经过验证的编辑 |
| 窄范围字面替换 | `docx_edit.replace_text()` | 保留原包；匹配必须落在单个 `w:t` 节点内 |
| 原生 Office 风格 PDF 转换 | `libreoffice.convert_to_pdf()` | 原始输入字节交给已安装的 LibreOffice，不经过 LDM 重建 |

受支持的内联图片可嵌入 DOCX/Markdown 或渲染到 PDF；Markdown 未指定外置图片目录时使用 base64。
TXT 不保留图片。复杂定位、字段、脚注/尾注、修订以及按分节区分的页眉页脚，即使能提取内容，也仍有保真限制。

## 安装

建议使用干净的虚拟环境，从本仓库的 `dev` 分支安装：

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install "git+https://github.com/shenyankm/DocWeave.git@dev"
```

Windows PowerShell 使用 `.venv\Scripts\Activate.ps1`，不要执行 `source`。
需要可复现部署时，将 `@dev` 替换为已经验证的完整提交哈希。

**`pip install aspose-words-foss` 安装的是上游，不是本增强版。** 不要仅凭类似的 PyPI 包名认定它来自本仓库。
两者共享 `aspose.words_foss` 导入空间，不适合在同一环境共存：安装或卸载可能覆盖/删除共享模块，
pip 不会检测这种冲突。仓库更名不意味着存在 `docweave` 导入名或同名发行包。

检查实际安装的模块及发行元数据：

```bash
python -c "import aspose.words_foss as aw; from importlib.metadata import version; print(aw.__version__, version('aspose-words-foss-enhanced'), aw.__file__)"
```

### 依赖

运行时依赖会自动安装：

| 依赖 | 最低版本 | 用途 |
|---|---|---|
| `olefile` | 0.46 | 旧版 DOC/OLE2 输入 |
| `fpdf2` | 2.8.9 | PDF 渲染及 WOFF 字体支持 |
| `pydantic` | 2.0.0 | 类型化文档模型 |
| `defusedxml` | 0.7.1 | 加固 XML 解析 |
| `Pillow` | 10.0.0 | 光栅图片处理及尺寸检查 |

可选组件需另行安装：`[shaping]` 提供 `uharfbuzz>=0.39.0`；`docxtpl` 用于模板示例；
LibreOffice 是独立应用。内置 PDF 路径不依赖这些组件，也无需 Word、COM 或系统字体安装。

## 快速开始

### 转换文件

```python
import aspose.words_foss as aw

doc = aw.Document("report.docx")
doc.save("report.md", aw.SaveFormat.MARKDOWN)
doc.save("report.pdf", aw.SaveFormat.PDF)
doc.save("report.txt", aw.SaveFormat.TEXT)
```

路径输入按扩展名选择 reader；`save(path)` 可从目标扩展名推断输出格式。
实际仅支持 DOCX、Markdown、PDF 和 TXT 输出。

### Markdown 与内存输出

```python
from io import BytesIO
import aspose.words_foss as aw

markdown = "# Report\n\n客户公司\n\n中文 **粗体** and English.\n"
doc = aw.Document(BytesIO(markdown.encode("utf-8")), aw.MarkdownLoadOptions())
doc.save("report.docx", aw.SaveFormat.DOCX)
pdf_bytes = doc.to_bytes(aw.SaveFormat.PDF)
print(doc.get_text())
```

流输入可通过魔数识别 DOCX 及 OLE2/DOC。Markdown 没有可区分的魔数，需传入 `MarkdownLoadOptions`
或显式指定 `LoadOptions.load_format`。标准 RTF 魔数能够识别，但内置 reader 仍拒绝该格式。
字节输入请包装为 `BytesIO`；`stream=` / `data=` 参数已弃用。

`to_bytes(format_or_options)` 必须显式指定格式，不创建输出文件。TXT 为 UTF-8；Markdown 尊重编码选项。
需要外置 Markdown 图片文件时，应使用 `save()`。

### 保存选项

```python
import aspose.words_foss as aw

doc = aw.Document("report.docx")
md_options = aw.saving.MarkdownSaveOptions()
md_options.encoding = "utf-8-sig"
md_options.export_underline_formatting = True
doc.save("report.md", md_options)

docx_options = aw.saving.OoxmlSaveOptions()
docx_options.pretty_format = True
docx_options.compression_level = aw.saving.CompressionLevel.MAXIMUM
doc.save("report-rebuilt.docx", docx_options)
```

这些选项描述的是已实现行为，不代表完整商业 API 兼容。DOCX 保存选项作用于**重建式转换路径**，
不控制 `DocxDocument` 的原包保存。

## 原包 DOCX 编辑

需要保留不支持的部件时，应使用 `DocxDocument`，而非经 LDM 重建 DOCX。
下面的自包含示例创建简单源文档，修改文本，并添加一张水平合并表格：

```python
from io import BytesIO
import aspose.words_foss as aw

source = aw.Document(BytesIO("客户公司".encode("utf-8")))
editable = aw.DocxDocument(BytesIO(source.to_bytes("docx")))
paragraph = editable.body.paragraphs[0]
paragraph.replace_text("客户公司", "新的客户公司")
paragraph.range(0, 2).apply_font(bold=True, size=12)
paragraph.paragraph_format.alignment = "left"

table = editable.create_table(1, 2)
editable.body.append_child(table)
table.rows[0].cells[0].paragraphs[0].append_child(editable.create_run("left"))
table.rows[0].cells[1].paragraphs[0].append_child(editable.create_run("right"))
table.merge_cells(0, 0, 2)
editable.save("edited.docx")
```

继续向该文档插入内联 PNG 和超链接：

```python
from io import BytesIO
from PIL import Image
import aspose.words_foss as aw

editable = aw.DocxDocument("edited.docx")
paragraph = editable.body.paragraphs[0]
link = paragraph.add_hyperlink("Documentation", "https://example.com/docs")
link.target = "mailto:team@example.com"

image = BytesIO()
Image.new("RGB", (96, 48), "navy").save(image, format="PNG")
paragraph.add_picture(image.getvalue(), width=72, alternative_text="Company logo")
editable.save("with-resources.docx")
```

- 文本范围按段落内 Python Unicode 码点计算，区间为 `[start, end)`，不是字素簇。
  替换及范围编辑只接受受支持的普通内联结构；复杂范围明确拒绝。
- 直接格式支持粗体、斜体、字号、对齐及已存在的段落/字符样式 ID。
  `effective_font` / `effective_paragraph_format` 只解析文档默认值和样式继承的已支持子集，
  未完整处理编号、条件表格样式、主题字体/颜色以及复杂文字规则。
- 图片插入接受 PNG/JPEG 的 bytes、流或路径。尺寸单位为点；只指定一边时保持宽高比。
  媒体去重、关系和 Content Types 更新、ID 分配均由该接口管理。
- 新增超链接只允许 HTTP(S)、mailto 或非空书签片段。关系 ID 在各 part 内分配；
  重定向一个共享链接不会篡改其他链接，不联网获取或验证目标。
- 水平合并要求受支持的完整网格，不能切开已有单元格。垂直/旧式合并、省略网格单元格、
  行列编辑、拆分以及跨文档/跨 part 导入仍不支持。
- 未修改部件的 payload 保持字节一致；已修改 XML 重新序列化，ZIP 重新打包。
  **整个文件不保证字节一致。** 该 DOM 拒绝数字签名包及 Strict OOXML 主文档。
- 正文和页眉页脚 story 可在支持边界内编辑。`to_light_document()` 返回独立快照：
  修改快照不会反写 DOM，快照转换仍可能丢失不支持的内容。

更窄的替换接口 `docx_edit.replace_text()` 及完整边界见 [DOM 指南](docs/docx-dom.md)。
保留原包**不等于安全清洗**：宏、嵌入对象及外部关系都会保留。

## 结构化内容与诊断

```python
import json
from dataclasses import asdict
import aspose.words_foss as aw

doc = aw.Document("report.docx")
content = doc.to_dict()
print(json.dumps(content, ensure_ascii=False))
print([asdict(item) for item in doc.diagnostics])
```

`to_dict()` 当前使用 `schema_version: 1`，以加法方式增加字段：

| 字段 | 含义 |
|---|---|
| `source` | 以路径加载时的输入路径，可能含敏感信息 |
| `blocks` | 有序正文段落/表格，包括格式、链接、图片元数据及模型位置 |
| 单元格 `blocks` | 有序段落与嵌套表格；原 `paragraphs` / `tables` 字段继续保留 |
| 段落 `note_references` | 脚注/尾注类型及标识符 |
| `source_stories` | 已提取的注 ID/内容，以及页眉页脚 part 的分节、变体和继承引用 |
| `headers_footers` | 原有 LDM 页眉页脚表示，不是完整分节渲染契约 |
| `diagnostics` | 累积记录的快照，含 `code`、`severity`、`location`、`message` |

位置为模型路径，**不是 PDF 页码坐标**。`get_text()` / TXT 按序包含正文嵌套表格文本，
过滤隐藏 run 和字段指令，显示真实超链接的标签，不会把普通文本中的字面链接语法误当成链接。
它们不附加 source stories、页眉页脚或图片。

**能够提取脚注/尾注，不代表能够渲染或完整写回。** DOCX/PDF/Markdown/TXT 转换会报告
`*.notes_omitted`；Markdown 对合并单元格几何丢失和嵌套表格展平发出诊断。
`diagnostics` 在加载和各次转换之间累积，包括失败前已检测到的警告；不会自动清空，也不是完整损失审计。
可将已知内容损失和缺字作为错误：

```python
import warnings
import aspose.words_foss as aw
from aspose.words_foss.pdf_writer import PdfMissingGlyphWarning

warnings.simplefilter("error", aw.ContentLossWarning)
warnings.simplefilter("error", PdfMissingGlyphWarning)
aw.Document("report.docx").save("checked.pdf")
```

单纯字体替换警告不会因上述过滤器而失败。将 JSON 导出到信任边界之外前，应过滤源路径/文本；
该输出不是经过脱敏的文档归档。

## PDF 字体与可选塑形

内置渲染器对子集化的 Document Sans SC 字体进行嵌入，覆盖常用简体/繁体中文、拉丁文字和标点。
使用独立粗体字面与派生倾斜字面；四个 WOFF 资源合计约 **25 MiB**。
不保留源字体家族、等宽字体度量或精确分页。缺字会报警；源字体替换可能改变布局。

阿拉伯语等复杂文字可安装塑形支持，并部署适合语言的可信 fallback 字体：

```bash
python -m pip install "aspose-words-foss-enhanced[shaping] @ git+https://github.com/shenyankm/DocWeave.git@dev"
```

按 [转换与安全说明](docs/enhanced-conversion.md) 设置 `PdfSaveOptions.text_shaping = True`
及可信字体路径 `fallback_fonts`。不会自动下载字体；fallback 的样式与 emoji 覆盖仍有限。
开启塑形或结构化导出，都不能据此证明 PDF/UA 或 PDF/A 合规。

## 有界 CLI 与 LibreOffice

### 内置后端

```bash
python -m aspose.words_foss.convert report.docx report.pdf --strict --timeout 60 --memory-mb 1024
```

该单任务 CLI 需要 **POSIX**，Windows 请使用普通库 API。默认墙钟超时 60 秒、内存阈值 1024 MiB、
单输出文件上限 256 MiB。Linux 还设置地址空间限制；macOS 使用可能短暂超限的 RSS watchdog，
不是硬内存配额。`--memory-mb 0` 关闭内存检查。
`--strict` 拒绝检测到的损失/缺字，不保证检测每种保真问题。
普通 `Document.save()` 不提供进程超时/内存隔离。

### LibreOffice 原文件后端

```python
from aspose.words_foss.libreoffice import convert_to_pdf

convert_to_pdf("report.docx", "native.pdf", timeout=60)
```

```bash
python -m aspose.words_foss.convert report.docx native.pdf --backend libreoffice
```

需另行安装 LibreOffice，让 `soffice` / `libreoffice` 位于 PATH，或使用 macOS 标准应用安装位置。
后端将原始 DOC/DOCX/RTF 字节交给每个任务的私有 profile，检查 PDF 后才发布结果。
它使用**系统字体**，不使用内置 WOFF；保真度仍受 LibreOffice 兼容性与已安装字体影响。
不会静默替换内置后端。`--backend libreoffice` 不支持 `--strict`、`--fallback-font`、`--text-shaping`。
本项目不提供常驻转换服务或队列。

**资源检查和私有 profile 不是安全沙箱。** 不可信输入仍需限制网络/文件系统访问、隔离原生解析器、
使用非特权账户，并由应用限制并发。

## 模板与其他示例

复用现有可选 `docxtpl` 集成，不提供新的内置模板语言：

```bash
python -m pip install docxtpl
python ApiExamples/template_report.py template.docx context.json report.docx report.pdf
```

模板必须可信。示例使用 Jinja `StrictUndefined` 与 autoescape；DOCX、PDF 分别发布，不是双文件事务。
模板路径及两个输出路径必须互不相同。

| [ApiExamples/](ApiExamples/) 中的文件 | 用途 |
|---|---|
| `convert_document.py` | 使用随附样本进行受支持的格式转换 |
| `loading_document.py`、`loading_markdown.py` | 路径、流、显式格式及 Markdown 导入 |
| `working_with_markdown_save_options.py` | Markdown 编码、换行、下划线导出 |
| `working_with_ooxml_save_options.py` | DOCX XML 格式化与压缩 |
| `working_with_pdf_save_options.py` | 内置 PDF 选项与转换 |
| `working_with_txt_save_options.py` | 正文文本提取与 TXT 输出 |
| `working_with_images.py` | 受支持的图文文档；TXT 仍只输出文本 |
| `template_report.py` | 可信 DOCX 模板生成及后续 PDF 转换 |

## API 入口

以下列出主要已实现入口，不是内部类型的完整目录：

| API | 契约 |
|---|---|
| `aw.Document(source, load_options=None)` | 路径或二进制流，构造时立即生成 LDM |
| `Document.save(path, format_or_options=None)` | 按扩展名推断格式或接受格式/保存选项对象；原子发布主文件 |
| `Document.to_bytes(format_or_options)` | 显式 DOCX/Markdown/PDF/TXT 内存输出 |
| `Document.get_text()` / `to_dict()` | 正文文本 / 结构化内容提取 |
| `Document.light_document_model` | 可变解析模型，不是保留原 OOXML 的 DOM |
| `Document.page_count` | 模型估算值，不是精确 Word/PDF 页数 |
| `aw.LoadOptions`、`aw.MarkdownLoadOptions` | 显式加载格式/编码；Markdown 空行保留和本地图片 opt-in |
| `aw.saving.MarkdownSaveOptions` | 编码、段落分隔、表格/列表/链接/HTML 导出模式、下划线和图片处理 |
| `aw.saving.OoxmlSaveOptions` | Transitional DOCX 压缩/XML 格式化/ZIP64 设置 |
| `aw.saving.PdfSaveOptions` | 已实现的 PDF/阅读器/图片/大纲选项、fallback 与可选塑形；部分兼容字段未使用 |
| `aw.DocxDocument(source)` | 原包 DOM，支持 `save()`、`to_bytes()` 及独立 `to_light_document()` |
| `DocxDocument.body`、`story(part_name)`、`part_xml(name)` | 正文、可编辑页眉页脚 story、只读 part XML 快照 |
| `Paragraph.replace_text()`、`range()`、`add_picture()`、`add_hyperlink()` | 上文所述的受验证、有限 DOM 操作 |
| `DocxDocument.create_table()`、`Table.merge_cells()` | 简单表格创建和有限水平网格合并 |
| `docx_edit.replace_text(source, destination, replacements)` | 单个 `w:t` 内字面替换；所有 key 必须命中；返回替换次数 |
| `libreoffice.convert_to_pdf(source, output, timeout=60)` | 独立的原文件 LibreOffice PDF 转换 |

## 支持范围与限制

- 不提供内置 PDF 读取、DOC/RTF 写出、OCR 或完整 Word 排版/字段计算引擎。
- LDM 转换不是无损往返。脚注/尾注单独提取，但不写入转换输出；批注、修订、复杂字段、内容控件、
  数学公式、浮动内容及页眉页脚变体仍有限制。支持内联图片不等于任意形状/图片定位或完整 OOXML 保真。
- Markdown 的合并网格与嵌套表格会丢失几何；parser 仅实现选定 CommonMark/GFM 子集。
  Tab 缩进沿用固定 `+4`，不是按列位置计算的 tab stop。
- Strict OOXML 写出抛出 `NotImplementedError`。`ECMA376_2006` 与 `ISO29500_2008_TRANSITIONAL`
  使用同一输出；`Zip64Mode.ALWAYS` 等同于 `IF_NECESSARY`，不强制生成 ZIP64 记录。
- 未实现 PDF/A、PDF/UA 合规。显式设置以下未使用 PDF 字段会报警：
  `text_compression`、`embed_full_fonts`、`use_core_fonts`、`font_embedding_mode`、`page_mode`、
  `color_mode`、`preserve_form_fields`、`memory_optimization`。
- Markdown `image_resolution` 未使用；`export_as_html=NON_COMPATIBLE_TABLES` 等同于 `NONE`。
  属性存在不意味着它影响输出。
- 默认上限：输入/单 ZIP 部件/图片数据 **64 MiB**、DOCX 总展开大小 **256 MiB**、**10,000** 个 ZIP 条目、
  单光栅图 **25,000,000** 像素、表格网格/跨度宽度 **1,024** 列。
  拒绝不安全路径、重复/加密 ZIP 条目及 XML 实体扩展；SVG 外部资源受到限制。
- Markdown 不下载远程图片。本地图片默认禁用；启用 `MarkdownLoadOptions.allow_local_images`
  时需要路径输入，访问范围限定在文档目录内。
- 主文件同目录暂存后原子替换。新文件默认权限为 `0600`，覆盖已有文件时保留权限位。
  外置 Markdown 图片不是多文件事务，原子替换也不等于完整崩溃恢复保证。
- 不要并发编辑/使用同一个可变文档。接口不清洗宏内容，也不提供完整沙箱或所有内容损失诊断。

详细限制见 [转换与安全说明](docs/enhanced-conversion.md) / [DOM 指南](docs/docx-dom.md)。

## 开发与测试

```bash
git clone --branch dev https://github.com/shenyankm/DocWeave.git
cd DocWeave
python -m pip install -e ".[dev,shaping]" build docxtpl
python -m pytest tests -q
python -m pytest ApiExamples -q --rootdir=ApiExamples -c ApiExamples/pytest.ini
```

`[dev]` 包含 pytest、PyMuPDF、pypdf、python-docx、docx2python，用于回归、独立 PDF 检查和 DOCX 内容对照。
塑形与 docxtpl 是可选运行时集成；按上面命令安装，可运行相关测试而非跳过。
未安装 LibreOffice 时，真实原生渲染测试会跳过。

CI 精简为 **4 个组合**：Linux Python **3.10/3.14**、Windows **3.14**、macOS **3.14**。
每组构建并安装 wheel，检查导入/字体/许可证/typing 资源，再在**源码目录外**各跑一次完整回归与示例。
中间版本不在每次提交中运行；已配置任务不代表已通过跨平台验证。

已记录的本地验证环境为 macOS/Python 3.13.15：**370 项回归测试、28 项 API 示例通过**，
并验证安装后的 wheel、docxtpl 及已安装 LibreOffice 的真实渲染。
环境和边界见 [当前核验报告](docs/ecosystem-optimization-validation.md)。
这些结果不证明 Microsoft Word 视觉一致性或全部 OS/Python 组合兼容。

```bash
python -m build --wheel --outdir wheelhouse
python scripts/benchmark.py --repeat 3 > benchmark.json
```

wheel 核验应在独立环境安装后，离开源码目录运行 `scripts/check_wheel.py`，并使用
`--import-mode=importlib` 跑 pytest；参见 [升级说明](docs/upgrade-notes.md)。
基准在新进程中记录导入/解析/布局/序列化/写文件耗时、RSS、PDF 字节数和页数。
最新小样本对照**不能证明整体提速**，个别中位数增加约 10%。

## 文档与资源

| 资源 | 内容 |
|---|---|
| [英文 README](README.md) | 内容对应的英文使用指南 |
| [DOCX DOM 指南](docs/docx-dom.md) | 文本范围、格式继承、资源、合并与原包保留规则 |
| [转换与安全说明](docs/enhanced-conversion.md) | 字体、诊断、限制、CLI 与 LibreOffice |
| [升级说明](docs/upgrade-notes.md) | 结构化输出、API、模板与打包 |
| [当前核验报告](docs/ecosystem-optimization-validation.md) | 内容/资源重构、370/28 测试结果及基准数据 |
| [早期优化报告](docs/optimization-report.md) | 历史 post2 核验与性能测量 |
| [早期核验报告](docs/verification-report.md) | 历史 post1 验证，不代表当前覆盖范围 |
| [Issues](https://github.com/shenyankm/DocWeave/issues) | 增强分支的问题与需求 |

[上游源码](https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python)、
[上游 PyPI](https://pypi.org/project/aspose-words-foss/)、
[上游 API 文档](https://reference.aspose.org/words/python/) 描述原项目，不代表本增强版的功能保证。
[商业 Aspose.Words](https://products.aspose.com/words/python-net/) 是独立专有产品，不承诺可以直接替换本项目。

## 许可证

库代码使用 [MIT License](LICENSE)，需保留版权及许可声明。
内置字体单独采用 [SIL OFL-1.1](aspose/words_foss/pdf_writer/fonts/OFL.txt)，不是 MIT；
来源和修改说明见 [字体说明](aspose/words_foss/pdf_writer/fonts/README.md)。
软件不提供担保；可选集成组件有各自的许可与部署要求。
