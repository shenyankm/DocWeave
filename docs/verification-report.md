# 已落地项真实核验报告

## 当前结论（修复后）

**首次核验发现的五个失败用例已修复，原断言和容差未放宽。**
纯文字测量复用实际渲染代码，并隔离输出的大纲、结构标签和列表计数；续行缩进持续生效。
Code/Quote 按实际字号和边界测量；keep_together 正确选择下一页/栏，过长段落允许拆分。
跨栏分页探测不再提前改变页边距，也移除了会让文字越过栏底的行高补偿。
新页清理旧栏的游标；多页长内容采用顺序栏流，避免单页平衡算法跳过第一栏。

- macOS / Python 3.12.14 / 最低 fpdf2 2.8.9：**102 项回归 + 27 项 API 示例通过**，无失败/跳过。
- 修复后重建 wheel，在 checkout 外导入 site-packages 后：**129 通过**，无失败/跳过。
- 新增 `tests/test_pdf_indent_flow.py`：14 项，覆盖负/正首行缩进、左右边界、混合字号/高亮/对齐、
  实际跨页/跨栏、页眉页脚、下一栏 keep_together、恰好容纳/超出 0.01 mm、过长段落。
- LibreOffice **26.2.6.3** 已安装；安装包大小和 SHA-256 与官方一致。真实中文 DOCX、文本 RTF、
  原始 DOCX 脚注及受限 CLI 转换通过。脚注在 LDM 中丢失，但在原文件原生转换 PDF 中可提取并栅格化。

已测纯文字布局问题得到解决；图文绕排仍是近似高度，不能据此宣称完整 Word 排版兼容。
没有操作本机 Microsoft Word，也没有完成与 Word 的逐页视觉保真对比。

## 首次核验结论（保留历史）

首次仅添加测试和报告，未修改运行时代码，得到五个真实失败；以下初始计数及反例保留用于对照。

## 首次核验环境和计数

- macOS / Python 3.12.14 / fpdf2 2.8.9 / PyMuPDF 1.28.2。
- 重新创建隔离测试环境，不依赖上轮已失效的 `/tmp` 环境。
- 原测试：57 项回归 + 27 项 API 示例，全部通过。
- 新验收测试：`tests/test_claims_verification.py`，**25 通过、5 失败、1 跳过**。
- 新构建 wheel 在 checkout 外安装后重测：**82 项回归通过、5 失败、1 跳过；27 项 API 示例通过**。
- 安装后确认模块来自 site-packages、分发版本为 `26.7.0.post1`、四个 WOFF 资源与源码 SHA-256 一致。

整体为 **109 通过、5 失败、1 跳过**。现有测试全绿不足以证明补充验收也通过。

## 逐项结论

| 项目 | 状态 | 实际证据/边界 |
|---|---|---|
| 中文高亮、居中/右对齐换行 | 通过已测场景 | DOCX → `Document.save()` → PDF；独立文字边界、文本完整性及栅格化检查 |
| 跨页链接和粗体/斜体高亮 | 通过 | 长链接文本实际跨页，逐页 URI、背景填色、粗体/斜体标志和文本完整性通过 |
| 段落高度与分页 | 已修复已测反例 | 原五项失败转为通过；新增跨页/跨栏、页眉页脚、混合字号、边界容纳用例通过 |
| 缺字、字体替换及未支持选项警告 | 通过已测场景 | 从公开 `Document` 入口验证全部八个显式未支持选项；strict 缺字会拒绝输出并保留原文件 |
| 真实字体 fallback | 通过 | 使用系统 Arial 的希伯来 alef，验证真实字形嵌入、轮廓与原字体一致、文字与栅格输出存在；不是合成码点测试 |
| 输入/ZIP/XML/图片和本地资源边界 | 通过已测场景 | 实际 64 MiB+1 输入、实际超过 64 MiB 部件及 256 MiB 总展开量均拒绝；已有 XML 实体、重复条目、像素、SVG、目录逃逸测试通过 |
| 原子主文件输出 | 通过 | PDF/DOCX/TXT/MD 实际写入一部分后抛异常，原文件仍逐字节不变，暂存文件清理；已有替换失败/符号链接/权限测试通过 |
| 单任务受限 CLI | macOS 已测行为通过 | 正常中文转换、strict、超时、RSS 超限、进程组子进程清理通过；真实查询到 CPU=61 秒、单文件=256 MiB 的 worker 配置 |
| 字体压缩与加载 | 通过 | 四种 WOFF 对比 `7f6a745` 的 TTF，字形顺序、cmap、hmtx 及 glyf/cmap/hmtx/name/OS2/loca 原始表均一致；公开保存普通正文只注册 regular |
| 包名与版本 | 通过 | 隔离安装新 wheel，核验 `aspose-words-foss-enhanced`、版本及真实导入路径；仍共享官方 import 命名空间 |
| LibreOffice 入口 | 真实已测场景通过 | 中文 DOCX、文本 RTF、原始 DOCX 脚注及受限 CLI 通过；不是 Word 全面保真认证 |

## 首次确认失败的反例

相同的 76 字中文段落、14 pt 字号、窄页面，通过真正的 DOCX 读取和 PDF 输出对照：

| 场景 | 估算高度 mm | 实际高度 mm | 结果 |
|---|---:|---:|---|
| 普通正文 | 55.316 | 55.316 | 一致 |
| 高亮正文 | 55.316 | 55.316 | 一致 |
| 首行缩进 72 pt | 55.316 | 62.230 | 少估一行 |
| 左缩进 36 pt | 76.059 | 55.316 | 明显多估 |
| Code 样式 | 55.316 | 29.633 | 明显多估 |
| Quote 样式 | 55.316 | 43.462 | 多估 |

此外，构造剩余高度介于首行缩进段落估算与实际高度之间的页面：
`keep_together=True` 仍将该段落拆到两页，说明差异会影响真正的分页，不只是测量数值。

根因位置：

- `aspose/words_foss/pdf_writer/writer.py::_estimate_paragraph_height`：没有计入首行宽度变化；
  左缩进扣除用于全部测量行，却与实际渲染的缩进处理不一致。
- `aspose/words_foss/pdf_writer/paragraph_renderer.py::_render_styled_block`：Code/Quote 使用自己的
  字号、字体或宽度设置，但估算器未按同一渲染分支测量。
- `paragraph_renderer.py::_maybe_keep_together`：依赖不准确估算，导致实际跨页。

修复没有调整上述原始断言。修复后相同 76 字场景的估算/实际高度（mm）为：

| 场景 | 估算 | 实际 |
|---|---:|---:|
| 普通正文/高亮正文 | 55.316 | 55.316 |
| 首行缩进 72 pt | 62.230 | 62.230 |
| 左缩进 36 pt | 76.059 | 76.059 |
| Code | 29.633 | 29.633 |
| Quote | 43.462 | 43.462 |

左缩进实际高度也发生了变化，因为修复前续行丢失缩进、错误地占用了整页宽度；不是简单降低估算值来迁就旧渲染。

## 字体实测

资源 **43,124,928 → 26,226,992 bytes，减少 39.2%**。
本机五次字体注册的小基准：四样式中位数约 440 ms，纯正文单样式约 85 ms。
仅表示字体注册耗时，不代表整篇转换的加速倍数，也不承诺 wheel 下载体积同比下降。

## 重现和产物

```bash
python -m pytest tests/test_claims_verification.py -v
python -m pytest tests/ -q
python -m pytest ApiExamples/ -q --rootdir=ApiExamples -c ApiExamples/pytest.ini
```

安装 LibreOffice 的本次环境中上述命令全部通过；未安装的环境仅跳过真实原生渲染用例。

本地 git-ignored 证据：

- `output/chinese_pdf_test/verify_baseline.txt`、`verify_examples.txt`
- `output/chinese_pdf_test/verify_acceptance.txt`、`verify_wheel.txt`
- `output/chinese_pdf_test/claims_probe/`：高度对照 JSON 和 PDF
- `output/dist/verification/aspose_words_foss_enhanced-26.7.0.post1-py3-none-any.whl`：首次核验的旧 wheel
- `output/chinese_pdf_test/fix_full_source.txt`、`fix_api_examples.txt`、`fix_wheel.txt`：修复后通过记录
- `output/chinese_pdf_test/fix_libreoffice_real.txt`、`libreoffice_bounded.pdf`：真实原生转换记录/产物
- `output/chinese_pdf_test/fix_height_probe.txt`：修复后测量对照
- `output/dist/repair/aspose_words_foss_enhanced-26.7.0.post1-py3-none-any.whl`：修复后 wheel

未执行 Microsoft Word 逐页视觉对比、Windows、Linux 硬地址空间限制或其他 Python 版本测试。
macOS RSS watchdog 允许短暂超限；单文件 256 MiB 只验证了实际配置，未制造 256 MiB 写入来验证触发。
普通库 API、私有 profile 和进程组资源限制均不是完整安全沙箱；Markdown 图片侧文件不是多文件事务。
