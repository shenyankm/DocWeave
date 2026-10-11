# DocKiln

[项目主页](https://shenyankm.github.io/DocKiln/)

[English](README.md) | [简体中文](README.zh-CN.md)

**用于文档转换、中文/Unicode PDF 生成、结构化提取和 DOCX 编辑的 Python 工具库，无需 Microsoft Word。**

## 功能概览

- 读取 DOCX、旧版 DOC 和 Flat OPC，写出 DOCX、Flat OPC 或 PDF。
- 使用内置中文字体生成 PDF，可选配置备用字体与文字塑形。
- 提取正文、有序段落与表格、格式信息，以及脚注、尾注、页眉页脚等来源内容。
- 在保留原始文档包部件的同时，编辑受支持的 DOCX 文本、格式、表格、图片和超链接。
- 提供带资源限制的转换 CLI，以及可选的 LibreOffice PDF 后端。

| 格式 | 内置读取 | 内置写出 |
|---|---|---|
| DOCX | 支持 | 支持 |
| DOC（Word 97–2003） | 支持 | 不支持 |
| Flat OPC | 支持 | 支持 |
| PDF | 不支持 | 支持 |
| RTF | 仅 OLE2/DOC-backed 文件，不支持标准文本 RTF | 不支持 |

输入按内容检测格式，包括 Flat OPC XML 文件和无扩展名文件。Flat OPC XML 输出请显式指定保存格式。独立的 LibreOffice 后端可将原始 DOC/DOCX/RTF 文件转换为 PDF，包括标准文本 RTF。

## 安装

需要 **Python 3.10–3.14**。建议在干净的虚拟环境中，从仓库的 `dev` 分支安装：

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install "git+https://github.com/shenyankm/DocKiln.git@dev"
```

Windows PowerShell 使用 `.venv\Scripts\Activate.ps1` 激活。需要可复现部署时，将 `@dev` 替换为已验证的完整提交哈希。

当前发行包名和导入名均为 `dockiln`，使用独立命名空间。迁移后不再提供 `aspose.words_foss` 导入入口；请将旧代码改为 `import dockiln as dk`。Markdown/TXT 读写及其加载、保存选项已移除；提取纯文本请使用 `get_text()`。

运行时依赖会自动安装。可选文字塑形支持：

```bash
python -m pip install "dockiln[shaping] @ git+https://github.com/shenyankm/DocKiln.git@dev"
```

## 快速开始

### 转换文档

```python
import dockiln as dk

doc = dk.Document("report.docx")
doc.save("copy.docx", dk.SaveFormat.DOCX)
doc.save("report.pdf", dk.SaveFormat.PDF)
```

`save(path)` 可根据输出扩展名推断格式。转换会重建受支持的内容，不是 DOCX 无损往返。

### 内存输入与输出

```python
from io import BytesIO
from pathlib import Path
import dockiln as dk

doc = dk.Document(BytesIO(Path("report.docx").read_bytes()))
docx_bytes = doc.to_bytes(dk.SaveFormat.DOCX)
pdf_bytes = doc.to_bytes(dk.SaveFormat.PDF)
```

`to_bytes(format_or_options)` 返回内存中的输出，不创建文件。

### 不重建文档包的 DOCX 编辑

```python
import dockiln as dk

editable = dk.DocxDocument("report.docx")
paragraph = editable.body.paragraphs[0]
paragraph.replace_text("客户公司", "新的客户公司")
paragraph.range(0, 2).apply_font(bold=True, size=12)
editable.save("edited.docx")
```

需要保留未知部件时，使用 `DocxDocument`。未修改部件的内容保持字节一致，但已修改 XML 和 ZIP 会重新写出。超出支持范围的复杂编辑会被拒绝；不支持数字签名包及 Strict OOXML 主文档。详见 [DOCX 编辑指南](docs/docx-dom.md)。

### 结构化内容与诊断

```python
import json
from dataclasses import asdict
import dockiln as dk

doc = dk.Document("report.docx")
print(doc.get_text())
print(json.dumps(doc.to_dict(), ensure_ascii=False))
print([asdict(item) for item in doc.diagnostics])
```

`get_text()` 提取正文文本，包括嵌套表格。`to_dict()` 包含有序内容块、格式、来源位置，以及提取的脚注、尾注、页眉页脚内容。位置是模型/XML 路径，不是 PDF 坐标。诊断会跨操作累积，但不代表完整内容损失审计；结构化输出可能包含敏感路径和文档内容。

## 命令行与 LibreOffice

在 **POSIX** 系统上使用带资源限制的转换：

```bash
python -m dockiln.convert report.docx report.pdf --strict --timeout 60 --memory-mb 1024
```

`--strict` 拒绝已检测到的内容损失和缺字，不能发现所有保真问题。普通库调用不提供进程超时或内存隔离；Windows 用户可使用库 API。

使用另行安装的 LibreOffice：

```bash
python -m dockiln.convert report.docx native.pdf --backend libreoffice
```

该后端使用原始输入文件和系统字体，不使用内置 PDF 字体。仅支持输出 PDF，不能与 `--strict`、`--fallback-font` 或 `--text-shaping` 组合使用。

## 限制与安全

- 不提供内置 PDF 读取、DOC/RTF 写出、Markdown/TXT 读写、OCR 或完整 Word 排版引擎。
- 复杂字段、修订、浮动内容、脚注/尾注及分节页眉页脚有保真限制，请用代表性文档验证结果。
- 内置字体覆盖常用中文和拉丁文字，不覆盖所有 Unicode。字体替换可能改变排版，缺字会报警。未实现 PDF/A 或 PDF/UA 合规。
- 输入、压缩包、图片限制及加固 XML 解析可降低风险，但**不是安全沙箱**。不可信文档及原生解析器仍需隔离；保留原始 DOCX 部件也会保留宏、嵌入对象及外部关系。
- 主输出文件采用原子发布。不要在并发操作间共享可变文档。

保存选项、字体、资源上限及后端细节见 [转换与安全说明](docs/enhanced-conversion.md)。

## 开发与测试

```bash
git clone --branch dev https://github.com/shenyankm/DocKiln.git
cd DocKiln
python -m pip install -e ".[dev,shaping]" build docxtpl
python -m pytest tests -q
python -m pytest examples -q --rootdir=examples -c examples/pytest.ini
```

CI 仅在 Linux 上覆盖 Python 3.10–3.14。先构建 sdist，再从 sdist 构建 wheel 并校验发行元数据；各 Python 版本在仓库外测试同一份已安装 wheel。历史基准材料在独立任务中校验，发行包和 JUnit 报告保留七天。

未安装 LibreOffice 或可选生态依赖时，相关测试会跳过。当前结果见 [GitHub Actions](https://github.com/shenyankm/DocKiln/actions)；测试通过不保证与 Word 渲染一致，也不证明 Windows/macOS 兼容性。

## 文档

- [DOCX 编辑指南](docs/docx-dom.md)：文本范围、格式、表格及原包保留规则。
- [转换与安全说明](docs/enhanced-conversion.md)：保存选项、PDF 字体、诊断、CLI 与 LibreOffice。
- [示例](examples/)：加载、转换、图片及可选的 `docxtpl` 模板。
- [生态对比](docs/ecosystem-adoption.md)与 [API 对齐材料](docs/commercial-alignment-plan.md)。
- [Issues](https://github.com/shenyankm/DocKiln/issues)：问题反馈与功能需求。

## 许可证与致谢

代码采用 [MIT](LICENSE)，原始声明见 [第三方来源与许可](THIRD_PARTY_NOTICES.md)；内置字体采用 [SIL OFL-1.1](src/dockiln/writers/pdf/fonts/OFL.txt)，来源见 [字体说明](src/dockiln/writers/pdf/fonts/README.md)。可选集成组件有各自的许可证。

DocKiln 起源于 [Aspose.Words FOSS for Python](https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python)，保留原始版权声明。本项目独立维护，不是 Aspose 官方发布版，也不保证 API 或渲染兼容。
