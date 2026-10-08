# 26.7.0.post2 优化核验报告

本轮实际执行结果，不沿用上一版的通过数。改动范围及 API 用法见 [升级说明](upgrade-notes.md)。
所有实现保留输入/ZIP/XML/图片限制、原子输出与已存在的回归断言；未以放宽容差或 xfail 隐藏失败。

## 1. 按优先级落地

| 阶段 | 交付 |
|---|---|
| 可信转换 | 标准文本 RTF 明确拒绝；DOCX 加载/写出已知损失及页眉页脚边界诊断；稳定警告分类；严格模式和失败保留旧输出的回归 |
| 工程保障 | Python 3.10–3.14；py.typed；三平台 CI；源码目录外安装 wheel 的测试与字体/许可资源检查 |
| PDF 表格 | 混合 run 格式、链接、单元格图片、连续前导表头重复、超高行拆分、横向跨度、合并边线、嵌套表格；实际几何/提取/栅格化检查 |
| 服务端接口 | Document.to_bytes()；Document.diagnostics 的 code/severity/location/message；不改变路径输出的原子性 |
| Markdown | 选定 CommonMark/GFM 规范样例及公共 DOCX 往返；保留现有 parser，不引入新依赖或宣称完整规范合规 |
| 性能 | 五类冷子进程基准、分阶段耗时、峰值 RSS、页数/大小；原 XML 文本批量替换用分片拼接，避免反复复制整个部件 |
| 有边界的扩展 | 原包保留的单 w:t 字面替换；可选 HarfBuzz shaping 与一致的测量；JSON-safe 模型内容；docxtpl 组合示例 |

**条件扩展没有变成无限扩张**：原包接口不是任意 OOXML 编辑器，多语言不是全语言 Word 渲染，
结构化输出不包含向量库/embedding/OCR。Markdown 的固定 +4 Tab 行为仍与 CommonMark 不同；
需要完整规范合规时再迁移成熟 parser。没有添加全局字体缓存或常驻 LibreOffice 服务池。

## 2. 本地测试

平台：macOS 27.0.1 / arm64。全部安装了可选 shaping 和独立的 docxtpl 示例依赖。

| Python | tests/ | ApiExamples/ |
|---|---:|---:|
| 3.10.21 | 187 passed | 28 passed |
| 3.11.16 | 187 passed | 28 passed |
| 3.12.14 | 187 passed | 28 passed |
| 3.13.15 | 187 passed | 28 passed |
| 3.14.7 | 187 passed | 28 passed |

本地上述版本无失败/跳过。期望警告为字体替换、现有复杂文字样例未开启 shaping 的提示，
以及 PyMuPDF/SWIG 的弃用警告，不将这些误报为全面保真成功。

Python 3.12 核验依赖：fpdf2 2.8.9、fonttools 4.66.1、pydantic 2.13.5、uharfbuzz 0.56.3、
docxtpl 0.20.2、pytest 9.1.1、pypdf 6.19.0、PyMuPDF 1.28.2。
Python 3.10 使用其支持的 fonttools 4.65.0。

命令：

```bash
python -m pytest tests -q
python -m pytest ApiExamples -q --rootdir=ApiExamples -c ApiExamples/pytest.ini
```

本地原始日志保存在 git-ignored 的 `output/tests-{310,311,312,313,314}.log`
和 `output/examples-{310,311,312,313,314}.log`。

## 3. 新 wheel 而非 checkout

构建 `output/wheelhouse-post2/aspose_words_foss_enhanced-26.7.0.post2-py3-none-any.whl`，
安装到新建的 Python 3.12 虚拟环境，并从 checkout 外的临时目录运行：

- `scripts/check_wheel.py` 确认从 site-packages 导入，发行/模块版本一致，
  包含 py.typed、四种 WOFF 字体和 OFL.txt。
- 已安装 wheel：**187 tests + 28 API 示例通过，无失败/跳过**。
- 测试和示例都使用 `--import-mode=importlib`。仅改变工作目录还不够：
  默认 pytest 导入方式会因示例包结构将 checkout 加回 sys.path。
  本轮发现这个问题后修正 CI 命令并复测，警告路径也明确来自 site-packages。
- `output/wheel-tests-312.log` / `output/wheel-examples-312.log` 保留本地结果。

```bash
# 在源码目录外运行；REPO 为 checkout 的绝对路径，python 来自 wheel 环境
python "$REPO/scripts/check_wheel.py"
python -m pytest "$REPO/tests" -q --import-mode=importlib
python -m pytest "$REPO/ApiExamples" -q --rootdir="$REPO/ApiExamples" \
  -c "$REPO/ApiExamples/pytest.ini" --import-mode=importlib
```

GitHub Actions 已配置三平台 × 五个 Python 版本的同类检查，
**本轮未在 GitHub 执行该矩阵，也未本地实跑 Windows/Linux**；
上述本地结果不能替代跨平台 CI 的结果。

## 4. 回归覆盖与实际修复

- RTF 文件/stream/bytes、内存四格式输出、编码、外置 Markdown 图片拒绝、未知格式拒绝。
- 诊断被 warning filter 忽略/升级为错误时仍记录；并发收集隔离；严格模式已知写出损失不能覆盖旧 PDF。
- DOCX 脚注、字段、修订、内容控件、公式、多页眉页脚和未知模型节点的已知损失边界；
  复杂字段标记正常写回，避免把已经支持的内容误报为丢失。
- 新版本/未分配 Unicode 与私用区字符的缺字检查。多版本测试最初发现 Python 3.10 的五项失败：
  不能用 isprintable() 直接忽略其 Unicode 数据库未认识的新字符；修复后原断言通过。
- 表格字体/链接/图像、横纵合并、嵌套、重复表头、超高行、行高/cantSplit 警告、
  跨页、旋转、表后文字及越界检查。修复剪裁 q/Q 与 fpdf2 字体缓存不同步的真实显示问题。
- 拒绝恶意/非法 gridSpan、无法容纳的表头，并保留已有输出。
- Arabic 检查真实嵌入的上下文字形轮廓，不只检查 Unicode 提取；开启 shaping 的混合中文/Latin/Arabic
  测量与实际渲染一致。序列化后清除测量 PDF 状态，防止重复 writer 使用共享的已子集化字体。
- 原 OOXML 定点替换保留脚注、未知 XML、二进制及未改 XML 字节；测试同时替换、空白/转义、注释、
  同路径保存、跨 run/缺失 key/签名部件/签名关系/DTD/非法替换拒绝与旧文件保护。
- docxtpl 示例覆盖中文/XML 特殊字符、PDF 可提取文本、输出扩展名不改变 PDF 格式、模板/输出路径冲突拒绝，以及 StrictUndefined 失败不覆盖原 DOCX。

Python 3.12 compileall 与改动相关模块的 Ruff 关键错误检查（F/E9/B020）通过，`git diff --check` 通过。
没有进行全库风格整改或完整 mypy 审计。

## 5. 性能基准

Python 3.12.14 / macOS arm64，每种用例 3 个新子进程，取中位数。
这是当前实现的可复测基线，**不是与竞品同机竞速，也不是改动前后的提升百分比**。
“冷”指新 Python 进程，不代表清空系统文件缓存。布局阶段包含字体准备与实际 PDF 内容渲染。

| 合成文档 | 页数 | import / parse / layout / serialize（ms） | 进程总耗时（s） | 峰值 RSS（MiB） |
|---|---:|---|---:|---:|
| 小文档 | 1 | 136 / 2 / 79 / 25 | 0.288 | 125.6 |
| 百页正文 | 100 | 127 / 77 / 914 / 34 | 1.209 | 136.6 |
| 图文表格 | 3 | 129 / 10 / 182 / 40 | 0.412 | 169.0 |
| 重复图片 | 4 | 128 / 8 / 81 / 24 | 0.287 | 127.3 |
| 混合样式 | 5 | 132 / 11 / 397 / 61 | 0.656 | 241.8 |

文件发布阶段约 1 ms；外层 process 时间包含子进程启动/退出。
完整分阶段时间、PDF 字节数等保存在本地 `output/benchmark-full.json`。
Windows 无 RSS 数据时返回 null，不混淆为零内存。

```bash
python scripts/benchmark.py --repeat 3 > benchmark.json
python scripts/benchmark.py --case hundred_pages --repeat 5
```

混合样式的内存明显高于纯正文；这提供了下一轮剖析重点，但没有未经测量就加入全局缓存或并发共享字体对象。

## 6. 未验证与仍需知悉的边界

- 没有完整 Microsoft Word 逐页视觉对照；测试可打开/可提取/可栅格化不等于 Word 保真。
- LDM 不保留任意原始 OOXML；脚注、评论、修订和页眉页脚变体等仍可能降级。
  strict 只检测已知问题，不能保证无损。
- 原包替换只在单文本节点内工作，保留宏/外部关系/未知内容，不是安全清洗器。
- 多语言依赖适合的可信字体；双向/塑形回归不代表所有语言和混合样式完整支持。
- 模型位置不是页坐标；单元格段落/子表格源顺序不能恢复，结构化输出不是通用脱敏器。
- 版本和本地 wheel 已更新，但未提交 Git、推送或发布到 PyPI。

## 7. 后续性能优化：缺字诊断

在现有工作区上测量并修改 `pdf_writer/diagnostics.py`：每个 run 和列表标签的缺字检查
不再对整张字体 cmap 做集合差运算，而是仅查询文本中实际出现的字符。
保留缺字、回退字体样式损失、私用区及未分配 Unicode 的诊断行为；未增加缓存或依赖。

Python 3.12.14 / macOS arm64，改动前后各 3 个新子进程的中位数：

| 用例 | layout 前 → 后（ms） | process 前 → 后（ms） |
|---|---:|---:|
| 小文档 | 80 → 78 | 290 → 281 |
| 百页正文 | 918 → 782 | 1219 → 1071 |
| 图文表格 | 182 → 173 | 406 → 395 |
| 重复图片 | 79 → 81 | 284 → 284 |
| 混合样式 | 400 → 375 | 650 → 624 |

百页正文的布局耗时减少约 15%，进程总耗时减少约 12%；优化后另测 5 次，
中位数为 layout 784 ms / process 1069 ms。短用例的细小差异视为测量噪声，
不宣称内存改善。cProfile 下百页正文的 `warn_about_conversion` 累计耗时
从 267 ms 降至 48 ms；该数字含剖析开销，不与上表直接比较。
原始 JSON 保存在本机 `/tmp/aw-perf-{before,after,after-pages}.json`。

核验：

- 新增 4 个参数化回归用例，禁止遍历主字体 cmap，并断言文本/列表缺字及回退样式警告。
- `PYTHONPATH=. /tmp/words-foss-verify/bin/python -m pytest tests -q`：187 passed / 4 skipped。
  本轮环境未安装可选 `uharfbuzz`，4 个 shaping 测试因此跳过；不替代第 2 节历史环境结果。
- 使用修改前后的诊断函数分别渲染五类基准，固定 PDF 生成时间后，所有 PDF 字节完全相同。
- `git diff --check` 通过。未重建 wheel、运行 API 示例或跨平台矩阵。
