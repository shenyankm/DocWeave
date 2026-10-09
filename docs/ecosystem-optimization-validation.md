# 生态借鉴优化：实施与验证报告

本报告记录早期 370/28 项验收阶段，不代表当前测试规模，也不是上游或商业 Aspose 的功能保证。
最新阶段及边界见 [生态吸收与验收](ecosystem-adoption.md)。
该历史阶段保留了原有 `pyproject.toml` 警告过滤以及两个已有测试文件的修改；当时没有自动提交或发布。
后续阶段按用户要求，在验收后更新文档并提交推送。

## 按优先级实施

| 优先级 | 事项 | 本次结果与边界 |
|---|---|---|
| 高 | 原包 DOM 的资源关系 | 图片/超链接插入与链接重定向；part-local relationship ID、媒体去重、Content Types、包级 drawing ID；不删除未知引用，不做跨文档导入 |
| 高 | 表格网格安全 | 简单完整网格的水平合并，跨度边界与行网格验证；保留内容顺序；复杂结构、vMerge、旧式 hMerge 和行列编辑仍拒绝 |
| 高 | 内容完整性 | 单元格段落/嵌套表格顺序贯穿模型 JSON、DOCX、PDF、Markdown、TXT 和结构化输出；原 paragraphs/tables 构造兼容 |
| 高 | Notes 与分节信息 | source_stories 提取脚注/尾注 ID、页眉页脚原 part 和 section/variant/inheritance 引用；原包 DOM 保存仍保留原始 parts |
| 高 | 明确损失 | 四种转换输出报告 notes_omitted，Markdown 报告合并几何丢失/嵌套表格展平；不将“已提取”说成“已渲染/完整回写” |
| 中 | DOCX↔Markdown 对照 | python-docx 独立生成文档、docx2python 独立提取对照；中文、强调、业务标题、列表、合并表格、复杂链接标签/URL；literal link syntax 不被文本/JSON 提取误当成真实链接 |
| 保持组合 | 模板与渲染 | 沿用可选 docxtpl 和原文件 LibreOffice 入口；验证实际模板替换和真实 LibreOffice 渲染，不另造模板语言 |
| 条件性 | 常驻服务 | 只测冷启动/私有 profile 的当前入口，没有实施常驻服务；暂无业务并发、SLA、沙箱和故障恢复验收数据，不能据单机单文件时长认定服务瓶颈 |

共享关系路径/URI 解析位于 `_opc.py`，链接编解码位于 `_links.py`；
网格合并读取复用 `iter_grid_cells()`，没有另加转换引擎或运行时依赖。
`python-docx` 和 `docx2python` 仅加入开发测试依赖。DOM 细节见 [DOM 指南](docx-dom.md)，
结构化输出的加法字段和兼容性变化见 [升级说明](upgrade-notes.md)。

## 自动与真实引擎验证

环境：macOS，Python 3.13.15，LibreOffice 26.2.6.3；docxtpl 0.20.2 在临时验证 venv 中安装。
venv 复用已有系统依赖，不是独立验证所有最低依赖版本。

| 验证 | 结果 |
|---|---|
| 修改前 tests 基线 | 302 passed |
| 修改后源码 tests | **370 passed**，新增 68 项 |
| 构建并安装 wheel，离开仓库使用 importlib 模式跑同一 tests | **370 passed** |
| 安装 wheel 后跑 ApiExamples，含真实 docxtpl 示例 | **28 passed**，无跳过 |
| wheel 导入位置、版本、py.typed、WOFF 字体和授权文件 | 通过 scripts/check_wheel.py |
| 实际 LibreOffice | 原有中文 DOCX/RTF/脚注/CLI 测试及新增 DOM 图片、链接、水平合并渲染均通过 |
| 修改范围的 Ruff 错误/未用符号检查、compileall、git diff --check | 通过 |

全量 tests 中的一项嵌套表格展平警告、示例中的字体替换警告均为预期降级信息，不是异常被吞掉。
保留原有输出原子发布测试，实际半写入失败、超限、关系冲突和无效编辑均有回归覆盖。

主要新增用例：

- `tests/test_docx_dom_resources.py`：原包保留、图片类型/尺寸/alt text、关系复用/重定向、ID、去重、失败前不变更、python-docx 互操作。
- `tests/test_content_integrity.py`：模型序列化、交错内容顺序、PDF 文字坐标、Markdown 网格/降级、source stories、Note references、reader 重用、URI 安全、受限水平合并。
- `tests/test_docx_markdown_corpus.py`：独立生成/提取对照、DOCX→Markdown→DOCX 语义、链接 PDF annotation、普通文本中的 literal link syntax。
- `tests/test_dom_libreoffice_integration.py`：真实渲染器能够读取新增 DOM 图片/链接/合并结果；未安装 LibreOffice 时跳过。

## 内置 PDF 性能对照

使用 `scripts/benchmark.py`：各样本新进程；计时含导入、解析、布局、序列化、原子写文件。
基线为 HEAD `8e826a97bb5d367314627d04009409923754e8df` 的源码快照，
在同一环境交替运行基线/当前版本，每例各 3 次取中位数，并检查实际导入位置；
不是并发测试的总运行时间。原始数据：[before](benchmarks/ecosystem-before.json) /
[after](benchmarks/ecosystem-after.json)。

| 用例 | 修改前 total_ms | 修改后 total_ms | 变化（负数为更快） | 页数 |
|---|---:|---:|---:|---:|
| small | 306.4 | 316.8 | 3.4% | 1 |
| hundred_pages | 1073.0 | 1072.6 | 0.0% | 100 |
| image_table | 456.8 | 503.5 | 10.2% | 3 |
| repeated_images | 346.7 | 363.0 | 4.7% | 4 |
| mixed_styles | 755.0 | 831.3 | 10.1% | 5 |

五个用例的 PDF 页数和字节数均与基线一致，但这不等于视觉逐像素相同。
**不能宣称整体性能提升**：本次主要改善正确性/能力；个别中位数变慢约 10%，3 次采样不足以
区分回归与系统负载波动。大批量/延迟敏感业务还需在固定硬件、真实样本上增加重复数、RSS/吞吐验收。
没有新增全局缓存、默认常驻进程或隐藏并发。

### LibreOffice 冷入口测量

同一生成样本，每次新私有 profile，3 次取中位数：
单页约 **1667.9 ms**，100 页约 **2088.9 ms**；均产出正确页数并提取到中文。
原始数据：[LibreOffice](benchmarks/ecosystem-libreoffice.json)。

这反映当前入口的固定开销，不分离启动、字体初始化和渲染成本。
没有测常驻 UNO/unoserver、并发队列、Linux 容器或故障恢复，因此不据此引入服务基础设施。

## 可复验命令

```bash
python -m pip install -e ".[dev,shaping]" build docxtpl
python -m pytest tests -q
python -m pytest ApiExamples -q --rootdir=ApiExamples -c ApiExamples/pytest.ini
python -m build --wheel --outdir /tmp/wheelhouse
python scripts/benchmark.py --repeat 3
```

安装 wheel 的验证需另建 venv，在仓库之外执行 `scripts/check_wheel.py`，随后用
`--import-mode=importlib` 测试；不能把仍从 checkout 导入的结果当成 wheel 验证。
交替基准应为两份源码分别设置 cwd/PYTHONPATH，交替执行 `--repeat 1 --case <case>`，
检查 `aw.__file__` 确实指向各自源码，再合并样本取中位数。

## 未完成的边界

没有验证 Microsoft Word 实际打开/修复提示、真实业务文档的逐页视觉一致性，或全部 Python/OS CI 矩阵。
不支持跨文档资源/样式/编号导入、完整垂直合并编辑和所有 OOXML 语义；source stories 的内容并不会自动写回。
保留原包/原生转换都不是安全清洗；不可信文档仍需网络/文件系统/原生解析器沙箱。
