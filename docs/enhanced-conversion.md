# 增强版转换、诊断与安全边界

本页描述本 fork 的 `26.7.0.post1`，不是官方 Aspose 的功能保证。

## 中文排版和字体

- 高亮文本、混合字体运行的居中/右对齐段落按实际字形宽度换行，保留显式换行与链接。
- 优先保留可放进一行的空格分隔单词；过长单词、无空格中文按字符拆行。
- 纯文字段落高度在隔离状态下试运行实际渲染器，包含首行/续行缩进、段前/段后间距和 Code/Quote 的有效字号；普通混合字号段落行高容纳最大字号。
- 左右缩进持续作用于续行及跨页/跨栏后的文字；首行缩进只作用于首行，可为负值（悬挂缩进）。
  `keep_together` 在下一完整页/栏能容纳时整体移动，超出完整区域的长段落仍允许拆分。图文绕排高度仍是近似估算。
- 四个完整字体从 TTF 无损压缩为 WOFF：资源约 **41.1 MiB → 25.0 MiB**，减少约 **39%**。
  字形映射、数量和宽度信息保留；PDF 内仍嵌入标准字体子集。只注册文档出现的字体样式。
  这是安装体积优化，不承诺下载体积同等下降（wheel 原本已 ZIP 压缩）。
- 尚未实现语言标点禁则、完整复杂文字塑形、源字体精确匹配或 Word 的完整分页规则。

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
| `PdfConversionWarning` | 其他可检测的 PDF 内容降级，如未知模型节点被丢弃 |
| `aw.loading.DocumentLoadWarning` | DOCX 的脚注/尾注引用、评论引用、OLE object、altChunk 不被轻量模型保留 |

默认值不会仅因选项尚未实现而报警；显式赋值即使等于默认值也会提示。
这些诊断只覆盖已检测到的情况，不是全面的 Word 兼容性检查。可用 Python `warnings` 分类过滤或升级为错误。

## 输入和输出边界

`Document` 的文件、字节和流入口统一有界读取：

| 限制 | 默认上限 |
|---|---:|
| 输入 | 64 MiB |
| DOCX ZIP 条目 | 10,000 |
| 单 ZIP 部件/图片数据 | 64 MiB |
| DOCX 总展开大小 | 256 MiB |
| 单张光栅图片 | 25,000,000 像素 |

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

`--strict` 将缺字和上述 DOCX 加载损失提示升级为错误，但不因正常字体替换而失败。
`--fallback-font /trusted/font.ttf` 可以重复指定，只用于内置 PDF 后端。
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
验证进程状态及 PDF 头/结束标记后才发布文件。支持一般文本 RTF 不代表内置 RTF reader 已修复。

这是独立可选入口，不改变 `Document.save()` 默认后端。高保真仍受 LibreOffice 兼容性和部署的
**系统字体**影响；它不直接使用本库 WOFF 字体。`--strict` 和 `--fallback-font` 不适用于该后端。
本仓库既测试协议/异常，也提供真实渲染测试。已在 macOS / LibreOffice 26.2.6 验证中文 DOCX、标准文本 RTF、
原始 DOCX 脚注和受限 CLI；脚注内容在轻量模型中丢失、在原文件转换中保留。未安装 LibreOffice 的环境跳过真实测试。
这些用例证明入口可用和已测内容保留，不证明所有文档与 Microsoft Word 视觉一致。

**私有 profile 与资源限制不是安全沙箱。** 原生 Office 解析器、宏、外部链接与文件系统访问必须由
应用/容器沙箱约束；对不可信输入关闭网络、隔离可读写目录、以非特权账户运行并保持 LibreOffice 更新。
尤其不要将普通 `convert_to_pdf()` 当作具备硬资源/外部资源隔离的服务器入口。
