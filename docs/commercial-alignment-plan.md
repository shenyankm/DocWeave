# Aspose.Words 26.9.0 全能力对齐计划

本计划取代生态计划的 13 项范围上限。Python 生态台账继续作为实现参考；完整 Word
引擎、格式、标准、DOM 和布局能力不能因成本高而自动排除。这里只交付基准建立和清单
初稿，不表示全能力台账、实现或验收完成。只有固定版本所有能力逐项验收，或用户明确
批准某项豁免后，才能宣布目标完成。

## 已确定基准

- 用户选择：商业基准为 Aspose.Words for Python via .NET **26.9.0**，Mac／Python **3.13**；
  DocWeave 目标为 Mac／Python **3.14**；先使用官方试用包，没有应用许可证。
- 本机：macOS 27.0.1、arm64；隔离环境实际为 Python 3.13.15 和 3.14.7。
- PyPI 包要求 `>=3.6,<3.14`，不能把强制安装到 3.14 当作有效官方基准。
- 官方包导入、`DocumentBuilder.writeln`、DOCX 保存已实际运行；产物确有试用水印。
  试用文档大小限制、水印及其新增图像/页边内容必须单独记录；不能从试用小样本推断
  无限制的大文档表现，也不能把水印差异当作 DocWeave 的错误。
- 包摘要、文档抓取摘要、环境和当前 CI 记录见
  [基准记录](benchmarks/commercial-26.9-baseline.json)。网页是滚动文档；版本边界以
  26.9.0 安装包及实际行为为准。API 首页、许可说明、26.9 发布说明的 URL、抓取时间和
  摘要保存到[快照登记](benchmarks/commercial-26.9-doc-registry.json)。原始 HTML 快照 ZIP
  保留在本地，本次仓库材料不复制网页全文；其他类型/成员页面及完整文档归档仍需补齐。
- 两套运行器必须在仓库外执行，防止本地 `aspose` 命名空间遮蔽商业包。没有反编译商业
  二进制或读取其实现；能力发现来自公开声明、文档和合法取得的试用包实际输出。

## 清单结构与完整性门槛

[安装包声明清单](benchmarks/commercial-26.9-api.json)包含 **30 个产品模块、8,065 个声明项**，
排除了 `__nuitka` 等打包工具目录。包括类型、枚举、枚举值、属性访问器、方法与构造器；
同名方法的重载声明逐一保存；集合下标等公开特殊方法也纳入清单，不能只保留构造器。
清单生成器是
[`inventory_commercial_api.py`](../scripts/inventory_commercial_api.py)，输入为固定包随附的
`.pyi`，每个文件保存 SHA-256。公开声明中 chart 文档字符串存在不合法引号，解析时仅
去掉文档字符串，保留声明；没有修改安装包。

这不是完整行为验收台账：默认值、异常、回调、协议、可接受的参数组合及格式方向尚待展开。
[运行时接口与继承记录](benchmarks/commercial-26.9-runtime.json)实际遍历了 32 个可见模块、
799 个导出类型，按实际 MRO 记录成员来源。Document 可见的已筛选成员为 159 项；这包含
继承接口，不代表 Document 的行为已经验证。

`lowcode` 在随包声明中缺失，但 26.9.0 运行时实际导出 25 个类型，且这些名称与
[官方模块文档](https://reference.aspose.com/words/python-net/aspose.words.lowcode/)吻合，必须
补入范围。其内建方法没有可提取的 Python signature 或 docstring，重载需要文档和行为探针
继续补证。`xattr` 同样可导入，但其公开支持状态未确认；保存观察记录，不能擅自当成完整
产品接口，也不能把未确认项视为已经关闭。

[合并能力台账](benchmarks/commercial-26.9-capabilities.json)目前 **26,208 条记录**，将声明、
继承接口、运行时额外导出与[已安装 DocWeave 入口](benchmarks/commercial-26.9-current-runtime.json)
对照。21,790 条有公开声明依据，4,418 条仅有运行时观察、支持范围待确认。534 条匹配到
同位置名称；这不是兼容能力数量或完成率。未匹配只说明等价公共名称缺失，不能据此断言
内部功能不存在。台账保留逐成员行为、具体依赖和验收条件未完成状态。
[运行时枚举记录](benchmarks/commercial-26.9-enums.json)已取得 LoadFormat 28 项、SaveFormat
42 项、NodeType 39 项；枚举存在不等于每一种加载/保存方向在 Mac 上可用。

[完整运行时枚举快照](benchmarks/commercial-26.9-all-enums.json)通过 `Enum.__members__` 保存
3,548 个名称/数值及 40 个别名关系；包含声明中的 251 个枚举、文档与运行时均确认的两个
lowcode 枚举，以及一个没有公共值、支持状态未确认的 xattr 枚举。`dir(Enum)` 会遗漏别名，
只筛大写还会遗漏混合大小写名称，所以不能据早期大写名清单宣称全量。

[当前公共路径数值对照](benchmarks/commercial-26.9-enum-comparison.json)记录 227 个名称的
整数一致，0 个数值不同，3,321 个对应入口缺失或不是整数。缺失不表示内部完全没有相关
枚举；同名整数一致也不能证明对应功能可用。NodeType 当前为普通常量类，与可构造的商业
IntEnum 的协议不同，不能只比较 16 个现有常量就判断类型兼容。

每个声明项继承 JSON 的 `capability_defaults`；合并台账用 `record_defaults` 复用同一字段。
除已关联的 DOM 局部观察外，当前逐项行为为 `not_measured`，实现记录区分名称存在/缺失但行为未验证，交付为
`not_started`。后续逐项覆盖这些字段：

| 字段 | 必须记录的事实 |
|---|---|
| baseline_behavior | 版本、平台、输入、默认与显式参数、结果/异常、试用干扰 |
| implementation / gap | 实际实现入口、支持程度、与基准的具体差异；不能凭同名判定兼容 |
| dependencies / approach | DOM、格式、布局、字体、外部服务依赖；独立实现方案 |
| samples / validation | 固定合法样本及哈希、执行命令、原始产物、独立解析/渲染断言 |
| exit_criteria / limits | 每项可判断的完成条件、容差、环境限制、未验证范围 |
| delivery | 对应提交、安装包、实际 CI；未执行保留未完成 |

继承图现已按运行时 MRO 展开；下一步核对声明存在但运行时缺失、仅运行时暴露和外部基类
接口的公开支持状态。每个类型必须覆盖其最终可见接口，构造器/重载/属性写入分别验证。
格式与选项应增加行为行，不能只用类型行代替数千项参数组合。计数的用途是检查遗漏，
不作为功能完成百分比。

## 实施顺序

| 阶段 | 内容与依赖 | 阶段退出条件 | 当前状态 |
|---|---|---|---|
| B0 固定基准与完整台账 | 扩展公开声明、继承/运行时导出；核对版本文档；逐项映射当前实现与格式方向；保存可复验语料 | 每项有行为/差异/依赖/方法/样本/验证/退出条件/限制/交付字段，未知项明确登记 | 已固定环境并取得声明、运行时继承和名称对照台账，逐项行为与现状审计未完成 |
| B1 安全与内容完整性 | 加载资源边界、失败原输出保留、恶意输入、隐藏内容、文字/图像/关系丢失；与基准逐项比较 | 支持内容无静默损失，拒绝和诊断可验证；兼容行为与安全要求冲突单列决策 | Markdown 普通文本有双方输出；DOM 单独删除未持久化已修复并完成安装包局部验证，整体未验收 |
| B2 完整 DOM 与核心格式 | 所有节点、集合、生命周期、顺序、所有权、跨文档导入、克隆；完整格式继承；统一 load/DOM/package/layout 映射 | 实际编辑→保存→独立回读→渲染验证，复杂节点/关系不被丢弃；所有受影响接口验收 | 现有原包 DOM 为受限基础，完整对齐未完成 |
| B3 其余公开功能与布局 | 所有固定版本模块；字段、列表、图形/图表、数学、邮件合并、报表、比较/修订、签名、宏、AI/服务接口；全部格式/选项与分页/字体 | 按完整能力台账逐项通过；外部依赖或 Mac 限制保留缺口，不能自动豁免 | 未开始完整对照 |
| B4 测量后优化与发布 | 在行为固定后测量时间/峰值内存/包体积；保持内容、结构、几何和像素结果 | 同语料/同环境对照；安装包和 CI 全验收；无未经批准的剩余能力缺口 | 原生态优化证据仅作历史参考 |

阶段可为前序缺口补证，但不得为追求性能跳过安全、DOM 或布局差异。重大问题按严重性
修复共享根因和全部调用路径，不以新增测试数量或一处小修改定义交付阶段。

## DOM 首轮差距映射

依据 [当前 DOM 指南](docx-dom.md)与 `aspose/words_foss/dom`，保留已有原包编辑优势，
不删除限制检查来冒充完整支持：

| 主题 | 当前已存在 | 必须补齐与验证 |
|---|---|---|
| 节点与顺序 | Body/Paragraph/Run/Table/Row/Cell/HeaderFooter 等视图；未知 XML 保留；有序子节点、所属文档与 part | 对照全部 NodeType/公开节点类型、集合、访问者、查找与生命周期；读/写/保存保持混合顺序 |
| 结构操作 | 受限简单节点插入、移动、删除、复制；合法父子和循环检查 | 跨文档/part 导入；图片/链接等关系及样式/列表 ID 重映射；复杂范围和修订；非法操作在修改前拒绝 |
| 文本与范围 | 普通 Run 编辑、普通段落替换和范围格式；过期范围检查 | 字段/超链接/内容控件/书签/批注/修订/换行等结构的可编辑语义；真实范围与保存后的行为 |
| 格式继承 | 部分直接格式；bold/italic/size/alignment 的有效值与样式链校验 | 完整属性的未设置/继承/显式状态；字符/段落/列表/表格/主题映射；读写不能擅自物化有效值 |
| 模型一致性 | 原包 DOM 和转换 LDM 各有能力 | 建立明确同步/失效机制，避免编辑原包后转换仍使用旧 LDM；节点、关系、布局实体与来源之间可追踪 |
| 诊断与保留 | 复杂操作拒绝、原包保留报告、转换损失诊断 | 每个支持声明与真实输出对应；不支持项与错误无静默降级；不能用 ZIP 保留证明渲染或字段结果正确 |

这只是主题级差距图，必须在 B0/B2 展开到逐个公开成员，不作为 DOM 完成证据。

## 固定语料和验证政策

保留已有 C01–C12，加入每个格式方向、节点、重载/默认值、字段/修订/宏/签名/图表/数学等
必要覆盖，既有 12 类不是上限。样本须合法生成或取得，记录工具版本、原件哈希与许可来源。
固定保存输入文件，不能只留随时间变化的生成脚本。负面和恶意样本与普通样本分别建档。

每个测试保留商业原始输出和 DocWeave 输出。针对文字顺序、节点、关系、样式、字段值、
图像原件、格式选项做独立断言；布局比较页数、页/栏/区域、位置、尺寸、字体、链接区域和
可见像素。容差由具体样本和度量理由决定，不用统一宽松阈值掩盖错页或漏字。
需要原生工具时记录实际版本和产物，不把商业 Aspose、Word、LibreOffice 互相等同。

试用限制影响的长文档、布局、内容或功能保留“未验证”，待可用授权再验收；没有授权不是
能力豁免。外部服务或凭据不足时也保留缺口，不伪造输出或删除对应行。

## 保存与加载格式首轮实际行为

[格式行为记录](benchmarks/commercial-26.9-format-behavior.json)使用同一份独立生成的小型 DOCX，
按 42 个 SaveFormat 值逐项显式保存到 `.bin` 路径；商业包使用其枚举参数，当前实现使用
已公开支持的整数参数，不依赖文件扩展名推断。全部原始输出和附属资源已保存到
[格式产物 ZIP](benchmarks/corpus/commercial-26.9-format-outputs.zip)，记录逐文件与整体摘要。

- 商业包返回并生成文件的格式为 41 个；`UNKNOWN` 抛出 RuntimeError/ArgumentException，
  未生成文件。包括 EMF 等格式在该 Mac 环境中实际产出，不能凭其他平台经验擅自排除。
- 当前 dd52d3d 安装包返回并生成文件的格式为 DOCX/Markdown/PDF/Text 和四种 Flat OPC，
  共八个，其余明确拒绝。
  商业保存返回 SaveOutputParameters，当前返回 None，返回契约仍未对齐。
- 返回和产出文件不是有效格式/内容/视觉保真验收；一个简单文档不能覆盖复杂结构、选项、
  字体、加密、标准或所有格式方向，elapsed_seconds 也不是正式性能基准。

加载探针使用上述真实保存产物，再执行 get_text 与 DOCX 保存，并用 python-docx 独立回读。
28 个 LoadFormat 中，CHM、DocPreWord60、MSWorks 缺少合法输入，保留未验证；其余 25 个
案例在商业包返回。当前实现返回 18 个案例，但不能将它们记为 18 种格式兼容。

当前四种 Flat OPC 小样本回存后均保留正文标签和一个表格。HTML、MHTML、WordML 和
显式 XML 案例仍被当作普通文本，回存 DOCX 中出现字面 XML/HTML，源表格由 1 个变成 0 个。PDF/MOBI/AZW3 出现 UnicodeDecodeError，
EPUB/ODT/OTT 出现缺少 word/document.xml 的 KeyError；标准文本 RTF 被当前内建 reader
明确拒绝。这些具体差距已写入 42 个保存与 28 个加载方向的台账，不能以“接口成功”关闭。

`LoadFormat.UNKNOWN` 在商业包中对原生 DOCX 实际自动识别并加载，不等同于保存 UNKNOWN
的拒绝行为。指定 XML 加载 WordML 的商业结果未包含目标标签；这是指定格式的观测，
不是正确 XML 样本的验收。加载选项与实际输入不匹配的语义还需进一步校准，不能只凭
`original_load_format` 或标签出现就判断显式参数生效。整个加载/保存方向均保留未完整验收。

### Flat OPC 公开加载集成

后续实现已将四种变体接入 `Document` 和 `DocxDocument`，更新后的格式台账已重新执行四个加载方向，旧的字面 XML
假成功案例不再代表最新实现。固定输入移至 `tests/fixtures/flat-opc-26.9.zip`；
[安装包加载记录](benchmarks/flat-opc-loading.json)确认正文、表格、编辑回读与其他部件字节保留。
新增边界覆盖 URI 转义、大小/数量限制、命名空间和错误 payload。
该记录仅关闭这四个小样本的加载与局部编辑问题；Flat OPC XML 输出、完整宏/模板语义、
复杂结构、选项优先级和视觉排版尚未验收，不能据此关闭全格式能力。

### Flat OPC XML 保存集成

[保存观察](benchmarks/flat-opc-saving.json)记录安装后的 Python 3.14 wheel 生成的八份实际输出：
四种 LDM 转换与四种原包 DOM 编辑。固定官方 26.9.0 包回读全部输出，正确识别 24–27，
正文标签与一个表格保留。官方 `OoxmlSaveOptions` 也实际接受四种变体。
当前格式快照已重跑全部 42 个保存值与 28 个加载值，并保存八种当前输出及全部原始资源。

Flat OPC 序列化会规范化 XML 属性顺序与根节点外空白。四份原生样本的所有非编辑 XML
经节点、属性、命名空间绑定、文本比较保持一致，二进制部件字节一致；不声称未编辑 XML
字节原样保留。原包 DOM 保留宏资源；LDM 省略 VBA 时新增明确诊断。
复杂宏/模板、全部 OOXML 选项、视觉排版和全量格式保真仍未验收。

## 首轮普通文本观察

[双方原始输出记录](benchmarks/commercial-26.9-literal-text.json)来自同九个 python-docx 生成的
DOCX，官方运行在 3.13.15、已安装 DocWeave wheel 运行在 3.14.7。官方输出含试用页边和图片，
完整保留，没有把这些内容剥离后宣称输出等价。
[原始输出和图片](benchmarks/corpus/commercial-26.9-literal-outputs.zip)逐文件摘要与记录一致；
JSON 的可读 `markdown` 字段将 CRLF 归一为 LF，原始输出摘要始终按归档字节计算。
原始 DOCX 已固定保存到[语料 ZIP](benchmarks/corpus/commercial-26.9-literal-text.zip)，记录
整体及逐文件哈希；不能用重新生成时变化的 ZIP 时间戳替换固定输入。
[`run_literal_corpus.py`](probes/run_literal_corpus.py)可分别在两套环境、仓库外重放：

```sh
cd /tmp
/tmp/docweave-commercial-26.9/bin/python \
  /Users/sheny/Developer/code/Aspose.Words-FOSS-for-Python/docs/probes/run_literal_corpus.py \
  /Users/sheny/Developer/code/Aspose.Words-FOSS-for-Python/docs/benchmarks/corpus/commercial-26.9-literal-text.zip \
  /tmp/docweave-literal-replay-commercial --module aspose.words --label commercial
/tmp/docweave-target-3.14/bin/python \
  /Users/sheny/Developer/code/Aspose.Words-FOSS-for-Python/docs/probes/run_literal_corpus.py \
  /Users/sheny/Developer/code/Aspose.Words-FOSS-for-Python/docs/benchmarks/corpus/commercial-26.9-literal-text.zip \
  /tmp/docweave-literal-replay-current --module aspose.words_foss --label docweave
```

普通文本 `<script>`、`<img onerror>`、Markdown link/image、HTML entity、反引号均被商业
试用版原样写出；所以“所有普通文本都必须转义”的原型不是既定商业验收标准。两方在反斜杠、
下划线和 autolink 等转义仍有差异，安全语义还需独立解析并登记策略。
[`markdown_literal_text.py`](probes/markdown_literal_text.py)是安全期望诊断原型，Mac／Python
3.14 运行结果 **64 failed / 12 passed**；它不进入正常 `tests/` 自动发现，不表示这些失败已被
修复，也不以期望错误的快照代替兼容验收。普通文本转义的生产代码未修改。

## DOM 结构首轮闭环与紧急持久化修复

[结构观察记录](benchmarks/commercial-26.9-dom-structure.json)覆盖十个普通 DOCX 场景：移动、
引用前插入、深/浅克隆、删除、无父节点删除、跨文档插入/导入、非法父子关系、循环。
原生输入与修复前/后及商业基准共 30 个实际输出已归档，并保留水印；用 python-docx 独立
回读检查目标顺序、表格内容和直接粗体，不依赖 writer 自证。

独立回读发现：原 `Node.remove()` 在断开节点后才调用节点自身 `_changed()`，节点已失去
到 XML 文档根的路径，原 part 没有标记为修改。内存删除成功、保存却保留原节点，是实际
数据完整性缺陷，不能为等待全量台账而保留。共享路径改为通知原父节点；其余保护检查不变。

新增段落/Run/表格/单元格段落/页眉段落五种“首次操作就是删除”回归，修复前五项失败。
源代码及实际安装 wheel 的 DOM 相关检查均 **142 passed**；轮包资源检查通过。固定真实
DOCX 的修复前独立检查仅 `remove` 出现内存/保存不一致，修复后十项全部一致；失败操作的
原包 payload 全部未变。保存结果和原文档字节/哈希已留存。

未解决的差异：浅克隆参数与行为缺失；跨文档导入缺失；删除返回值及错误类型不同；复杂
节点/范围/关系操作受限。上述证据已关联到合并台账的五类接口，但限定普通段落及当前语料，
不能推广为所有继承节点/重载通过。最终渲染、完整格式继承、其他节点和全量接口仍待验证。

`4a74ec6` 的四项 CI 全通过。持久化修复与 DOM 复验材料已提交并推送为 `a1b5a60`，
[该提交 CI](https://github.com/shenyankm/DocWeave/actions/runs/37906411435)已核验全部四项成功；
全量清单与普通文本诊断材料的检查独立登记，不能归入上述历史 CI 结果。

本轮清单工具的专项检查在 Python 3.14.7 上 **11 passed**，覆盖声明重载/属性 setter/集合
特殊方法、缺失声明模块、模块循环、继承成员来源、枚举别名、常量类/枚举协议区别，以及
“同名存在不代表兼容”、缺少合法样本与字面 markup 伪成功的台账状态、篡改/缺失产物拒绝、Windows 文本换行及二进制损坏检查、保存拒绝/返回契约/加载异常的具体差距。
`python scripts/verify_commercial_baseline.py` 检查声明/继承/枚举计数、文档来源登记摘要、
固定输入、100 个格式产物及独立 DOCX 结构、九个双方 Markdown 原始字节。
CI 的四个平台均执行该校验器；登记 JSON 的摘要按 LF 规范化，归档和输出仍按原始字节。
Ruff 与 JSON/语料哈希检查通过；九个双方输出通过固定语料重放逐字节复现。
没有重跑生产全套或新增文件的远端 CI；这些检查不构成行为全面验收。

## DOM 复制深度闭环（本地验证通过，平台验证待 CI）

`Node.clone(deep=True)` 保留现有无参数深复制调用，补充 `clone(False)`：容器保留直接格式、
清空内容子节点，Run 保留文本与字体格式。表格网格保留为格式元数据，不计入内容子节点。
复制节点保持原归属但无父节点；未插入的副本不标记原包修改。已有范围与资源保护仍适用。

[双方重放记录](benchmarks/dom-clone-depth.json)与[20 个原始输出](benchmarks/corpus/dom-clone-depth.zip)
来自固定原生 DOCX、官方 26.9.0／Python 3.13.15 与实际安装候选 wheel／Python 3.14.7。
普通段落的浅复制结果（空文本、零内容子节点、原归属、无父节点）一致；十项保存结果均通过
python-docx 独立回读，无内存/保存不一致。官方 trial 产物未清除。五种非布尔参数双方均
抛出 TypeError，未宣称异常消息完全一致。
[五类节点重放脚本](probes/run_clone_depth.py)按语料 ALPHA 标记定位段落，保留官方试用内容；
段落、Run、表格、行、单元格两种深度的节点类型、内容子节点数量、归属和脱离状态均一致，
Run 文本与直接粗体也一致。这十项观测不代替全部格式属性与复杂节点验收。

段落、Run、表格、行、单元格的专项 fixture 检查覆盖直接格式、网格、归属、副本独立性、
无副作用与段落/Run 保存回读；源码与安装 wheel 的 DOM 专项均 **96 passed**，安装 wheel
示例 **31 passed**，Ruff 与 diff 检查通过。实际安装 wheel、checkout 外运行的全套检查为
**2455 passed / 2 skipped**（344.45 秒）。本地实现与上述有限样本验证通过；远端平台验证
尚待本批 CI，不能据此宣称完整 DOM 阶段验收或全量能力完成。
复杂节点复制、范围/资源、跨文档导入、其他节点与完整渲染仍有缺口。

## 跨文档导入基准校准（尚未实现）

[导入观测](benchmarks/import-node-26.9.json)以自行生成的源/目标 DOCX 覆盖样式冲突、
新派生样式、默认样式、编号、图片、表格、Run，共七个对象 × 两种深度 × 默认及三种显式
格式模式。官方 26.9.0 56 项均返回归属目标、无父节点的副本，来源父节点保持不变；
当前实际安装 wheel 56 项均因 `DocxDocument.import_node` 缺失失败。失败后目标包全部
payload 未改变。两个输入和双方 112 个输出已[归档](benchmarks/corpus/import-node-26.9.zip)，
每个输出经过 python-docx 独立回读与原始字节摘要校验，试用水印未清除。

观测到的实际差别：默认/USE_DESTINATION_STYLES 使用目标冲突样式；
KEEP_SOURCE_FORMATTING 在冲突段落中将源字体属性写成直接格式；KEEP_DIFFERENT_STYLES
创建重命名源样式。源中新派生样式被导入时，即便 basedOn 映射为不同目标格式，也保留显式
源有效字体值。单独 Run 导入不会简单复制原段落的全部继承格式。四个深度图片导入输出的
图片字节摘要与原始蓝色图片一致；表格深导入保留单元格正文。记录保留直接字体和样式链
原始层次，不以这次解析宣称完整有效格式、编号语义或渲染正确。

下一实现依赖样式标识/名称匹配、继承与默认值、编号映射、关系与资源复制、节点归属和
事务提交，不能用仅改 XML ownerDocument 的复制替代。相同文档导入、重复 ID 分配、
样式名与 ID 不一致、更多资源/范围、损坏输入和 ImportFormatOptions 尚待校准；完整公开
导入能力仍未验收。本轮仅新增生成/观测/独立检查工具与材料，Ruff、自检通过；未重跑
生产全套，因为生产实现没有变化。

跨文档导入现进入实现：共享入口处理归属、两种深度、目标同名样式映射、简单新样式 ID
冲突/复用，以及依赖校验后的提交。不同默认值/主题、冲突基样式、编号、复杂资源和另外
两种跨文档格式模式仍未完成，不能宣称导入阶段验收。官方实际同文档 import_node(False)
产生零子节点的副本，未按声明中“simply a deep clone”字面描述硬编码成深复制。

首版安装 wheel 的 DOM 专项 118 项通过，56 组重放有 16 项返回；独立对照发现其中两项
浅表格保存差异（官方省略零行表格）。已新增保存规范化与注释/修订元数据保留回归，
源码导入专项 25 项、DOM 综合 121 项通过。首版安装 wheel 全套为 2477 passed / 2 skipped，
该结果不包含后续空表格修复。含该修复的新版 wheel 已构建并安装，56 组重放的 16 个返回
组合均与官方独立观察一致，40 个未实现组合均未改变目标包 payload；新版全套检查正在
运行。代码尚未提交推送。编号、图片和有效格式转换继续作为本阶段未决实现，保留全量目标。

空表格规范化复核又补充未知 Word 元素、表格描述保留，源码 DOM 综合为 123 passed。
[最新安装轮包重放](benchmarks/import-node-current.json)保留 56 个当前原始输出；16 个有限
观察与官方独立解析一致，40 个明确报未实现且原目标 payload 不变。源/安装包的导入和
有界转换专项合计 48 项通过。此前轮包全套 2480 passed / 2 skipped，不包含这次元数据
保留与进程清理修复；最终候选轮包全套仍在运行，未合并这些不同版本的验收结果。

`34fd252` 远端 macOS CI 的 test_timeout_kills_descendants 失败源于重复清理进程组：
超时分支已成功 killpg 并 wait，finally 再 signal 导致 PermissionError 掩盖 TimeoutExpired。
旧安装包回归已复现，修复后有界转换 21 项通过；保留首次权限失败、正常退出后的后代清理。
独立提交 `03aab7b` 已推送，其平台 CI 运行中。这不是导入阶段或全量行为验收通过。

最终候选 wheel 已在 checkout 外完成全套：**2483 passed / 2 skipped**（372.36 秒），
示例 **31 passed**。导入与进程专项 **48 passed**，DOM 专项 **123 passed**；最终版本
覆盖未知元素/描述保留与进程重复清理修复。后续仅变更复验工具，工具专项 **12 passed**，
含 Windows 换行、伪造字体观察和损坏导入归档拒绝。CI 证据校验器现独立解析并校验双方
历史与当前共 **168 个导入输出**，复查 16 个观察匹配；behavioral_acceptance 仍为 false。
此批本地实现和有限样本测试通过，远端平台验证待本批 CI；完整导入阶段仍未验收。

导入实现与复验材料已提交并推送为 `c50b4a1`，
[本批平台 CI](https://github.com/shenyankm/DocWeave/actions/runs/37923060688)运行中。
此前独立进程清理修复 `03aab7b` 的四项平台 CI 已全部成功；不能把这一历史通过状态
当作新导入实现的平台验收。当前未实现的模式、编号、资源和完整格式/渲染继续保留在全量范围。

### 粗斜体继承的官方校准

为解决新派生样式导入所依赖的有效格式差异，先校准共享读取路径。官方 26.9.0 试用包
在 Mac / Python 3.13.15 读取了 [745 个自有输入](benchmarks/corpus/style-toggles-26.9.zip)：
729 个 docDefaults、两层段落样式、两层字符样式、直接格式三态组合，以及 16 个默认
字符样式显式/隐式引用组合。粗体和斜体分别实际读取，
[官方记录](benchmarks/style-toggles-26.9.json)保存版本、平台、输入摘要与 getter 值。
[生成与重放脚本](probes/style_toggles.py)不保存试用文档，也不移除试用限制或水印。

实际行为与此前逐层 XOR 假设不同：每条 basedOn 链取最近显式值；段落和字符两个层级
都有定义时，异或结果与文档默认开启值合并；仅一个层级定义时采用该值；直接格式最后覆盖。
未引用的默认字符样式不应用到 Run。两个已有测试的旧预期已由官方实际读取复核并修正。
共享 StyleResolver 的修改覆盖 Run.effective_font，保留原始直接格式与包内容不变。

[安装 wheel 前后对照](benchmarks/style-toggles-current.json)中，旧 `c50b4a1` wheel 有
106 个观察差异，修复 wheel 的 745 个观察全部一致；记录各自 wheel SHA-256。
相关源码测试 868 项、复验工具 13 项、安装包示例 31 项通过；安装包在 checkout 外
全套为 **3230 passed / 2 skipped**（363.23 秒）。后续只加强默认字符样式字号测试和
复验工具断言，安装包最终格式与证据专项 **788 passed**，不把这些工具变化描述成重新运行生产全套。
CI 校验语料、记录与前后差异清单，并拒绝缺失案例、错误摘要及伪造匹配。
能力台账只给 Font.bold / Font.italic 挂接有限证据和替代入口，仍明确官方 Font getter
接口的整体对齐未完成。这批不覆盖表格/列表格式、复杂文字、其他字体属性、保存与渲染，
也未完成基样式冲突的导入转换；这些能力继续保留在完整目标中。

上一批导入生产提交 `c50b4a1` 的
[平台 CI](https://github.com/shenyankm/DocWeave/actions/runs/37923060688)四项已全部成功。
该结果对应上一批代码，不作为本次继承修复的平台验收。

### 导入样式祖先冲突与新增基准观察

粗斜体继承提交 `e6817f7` 的
[四项平台 CI](https://github.com/shenyankm/DocWeave/actions/runs/37925782625)已全部成功。
本轮继续检查新样式导入：[247 组自有输入](benchmarks/corpus/style-import-conflicts-26.9.zip)
覆盖段落/字符样式的粗体、斜体、字号，段落对齐，默认值、缺省/显式值与祖先链冲突。
固定官方试用环境实测并保存 247 个原始输出；当前旧 wheel 和修复 wheel 也各保留
247 个输出，合计 [741 个原始文件](benchmarks/corpus/style-import-conflict-outputs.zip)。
[完整观察记录](benchmarks/style-import-conflicts-26.9.json)包含输入/输出摘要、实际 getter、
独立 python-docx 原始样式/Run 层次与官方回读，保留试用水印。
[生成/重放工具](probes/import_style_conflicts.py)与
[独立检查工具](probes/inspect_style_imports.py)可复验，不能用它们宣称渲染正确。

旧 wheel 返回 83 个场景，其中两个祖先冲突案例的保存输出经官方回读后粗体不一致：
源与目标直接 Base 的格式 XML 相同，上层 Ancestor 粗体不同，旧检查只看直接 Base。
共享导入路径现比较全部祖先的格式层，含所有 tblStylePr 条件区段；复杂依赖明确拒绝。
段落、字符、后续表格条件区段三个回归均在旧安装包失败，新源码导入专项 30 项通过。
这是消除错误成功：尚未实现冲突格式转换，不将改为拒绝视为兼容能力完成。

修复 wheel 的 81 个返回输出均由官方实际回读，其粗斜体、字号和对齐 getter 与对应
官方保存重开输出一致；166 个未实现案例的目标包不变。安装包语料/导入/证据专项
**291 passed**，示例 **31 passed**，生产 wheel 在 checkout 外全套
为 **3233 passed / 2 skipped**（392.21 秒）。后续新增的语料/工具测试已按专项复验，
未描述为重新运行生产全套。CI 还会验证输入与
741 个输出摘要、独立样式层、官方回读匹配，并拒绝伪造观察或损坏归档。

新增官方观察也限制了实现假设：7 个场景的导入 getter 不等于源 getter，不能把“所有
新派生样式均保留源有效字体”作为规则。另有两个官方场景保存重开后的粗体与导入时不同，
独立 XML 检查发现目标 Base 的冗余粗体属性被规范化；未预读源格式的复测仍出现同样变化。
保留这些数据作为尚待解释的基准行为；本轮不改写已有目标样式来猜测复刻此变化。
下一步据完整上下文推导段落/字符样式各属性的迁移语义，继续完成格式转换和未实现模式、
编号、资源、完整 Font 接口与渲染验收；全量范围保持不变。

### 新段落/字符样式的受测属性迁移

祖先安全修复 `d74c5aa` 的
[四项平台 CI](https://github.com/shenyankm/DocWeave/actions/runs/37928947874)已全部成功。
后续迁移复用同一 247 组输入，在官方 26.9.0 试用版／Mac Python 3.13.15 与
DocWeave 安装包／Mac Python 3.14.7 上继续验证。新增简单样式的粗体、斜体、显式字号
和普通段落对齐现在解析全部祖先，处理既有目标基样式冲突和本次新建的基样式。
段落迁移去除与目标基样式相同的冗余属性；字符迁移保留源样式显式继承值，缺省字符
属性继续继承目标。已有目标样式保持不变，失败仍在提交前拒绝。

成功场景从 81 增至 209，剩余 38 组默认开启粗体的冲突场景仍未实现。官方实际回读
全部 209 个新安装包输出，粗体、斜体、字号、对齐与对应官方保存重开的 getter 无差异。
[本批记录](benchmarks/style-import-translated.json)引用原始官方基准，另保留
[247 个本批原始输出](benchmarks/corpus/style-import-translated.zip)；原 741 个输出与历史
结果保持冻结。当前语料重放检查原始样式层、节点归属、源不变和拒绝时目标不变，
证据检查覆盖摘要、独立层次、官方回读、Windows 文本换行及伪造/损坏记录。

新增迁移保护拒绝其他冲突属性、嵌套 run 样式、重复属性及带未知属性/子节点的格式叶子，
避免覆盖元数据。条件表格样式、复杂对齐、不同主题/默认值与隐式应用字号覆盖目标字号
仍未实现。[五组默认字号观察](benchmarks/font-defaults-26.9.json)及
[自有输入](benchmarks/corpus/font-defaults-26.9.zip)实测：无 rPrDefault 的三个受测输入
正文 Run/Style 字号为 11pt；空或粗体 rPrDefault 的两个输入为 10pt。探针定位自有
IMPORT 段落，避免读到试用水印的 12pt；[生成与观察工具](probes/font_defaults.py)可复验。
这些少量观察不足以定义全量默认格式，当前保护不写入猜测字号。

本批生产安装包在 checkout 外全套为 **3488 passed / 2 skipped / 199 warnings**
（413.52 秒）；随后补充的证据检查包含在最终专项 **300 passed**（11.04 秒）中通过，
示例 **31 passed**（5.95 秒）。Ruff、wheel 资源/导入守卫及冻结证据检查通过；新提交
平台 CI 待推送后验证，历史绿色状态不作为本批验收。尚未将 getter 匹配作为渲染或
完整 Font/DocumentBase API 验收。KEEP_SOURCE_FORMATTING、KEEP_DIFFERENT_STYLES、
编号、资源、默认开启的格式上下文及完整默认属性语义继续在全量范围内推进。

### 双段落验收与默认开启的段落样式

`e8187c5` 的 [四平台 CI](https://github.com/shenyankm/DocWeave/actions/runs/37934246343)
全部成功。该阶段的 209 个官方回读仅检查导入的 IMPORT 段落，后续补验发现其中三个
输出的既有 DESTINATION 段落粗体与官方保存重开结果不同，不能以 209 个导入段落匹配
宣称目标原有内容格式也对齐。历史归档保留；[本批记录](benchmarks/style-import-default-on.json)
新增两个自有段落的官方 getter 与匹配检查，保留三项差异的实际旧输出。

段落样式迁移现以 docDefaults 的粗斜体作为缺省基值，无字符样式上下文的默认开启
冲突可迁移；新增检查在提交前拒绝尚未校准的字符上下文基样式规范化。原 247 组中
224 组返回，两段落的冷回读均匹配官方；23 组仍未实现且拒绝时源与目标包不变。
新增[132 组自有输入](benchmarks/corpus/paragraph-style-defaults-26.9.zip)覆盖默认斜体、
粗斜体及段落/字符组合。旧包返回 60 组，其中 12 组双段落格式不匹配；本批返回
114 组且双段落无差异，18 组字符上下文仍未实现。
[观察与原始输出](benchmarks/paragraph-style-defaults-26.9.json)保留官方、旧包和本批安装包
全部结果，包括 9 个官方导入前后与保存重开的变化；未预读源字体 getter 的复测仍
复现这些变化，原始复测输出亦归档。

[494 次官方仅加载/保存](benchmarks/style-save-roundtrips-26.9.json)分别验证原 247 组的
源与目标，19 次字体 getter 重开后变化，说明部分差异并非导入操作独有。
原始输出保留试用水印；CLI 的 commercial-read-stories、commercial-roundtrip 和
commercial-no-source-getter 可复验。CI 同时验证官方原始输出的结构/摘要、SDK 回读
已观测字段及伪造目标段落格式拒绝，不能据此宣称完整 Font/Style 或渲染保真。

本批生产安装包在 checkout 外全套 **3628 passed / 2 skipped / 199 warnings**
（435.50 秒）；后续补充的证据检查包含在最终专项 **441 passed**（28.08 秒）中通过，
示例 **31 passed**（8.05 秒），Ruff、wheel 资源/导入守卫和冻结证据检查通过。
六个新增字符基样式回归在旧安装包均失败，在本批通过；新提交 CI 待推送后验证。
默认开启的字符格式、实际保存规范化、
全部导入模式、编号/资源、完整规范 DOM 与公开 Font/Style API、排版和渲染继续属于
完整目标，均未以保护性拒绝作为完成或范围豁免。
