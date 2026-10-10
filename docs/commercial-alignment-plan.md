# Aspose.Words 26.9.0 全能力对齐计划

本计划覆盖固定基准的公开能力。Python 生态台账作为实现参考；完整 Word 引擎、格式、
标准、DOM 和布局能力不能因成本高而自动排除。公开清单与局部实现持续更新，
不表示全能力实现或验收完成。只有固定版本所有能力逐项验收，或用户明确
批准某项豁免后，才能宣布目标完成。

## 已确定基准

- 用户选择：商业基准为 Aspose.Words for Python via .NET **26.9.0**，Mac／Python **3.13**；
  DocWeave 目标为 Mac／Python **3.14**；先使用官方试用包，没有应用许可证。
- 本机：macOS 27.0.1、arm64；隔离环境实际为 Python 3.13.15 和 3.14.7。
- PyPI 包要求 `>=3.6,<3.14`，不能把强制安装到 3.14 当作有效官方基准。
- 官方包导入、`DocumentBuilder.writeln`、DOCX 保存已实际运行；产物确有试用水印。
  试用文档大小限制、水印及其新增图像/页边内容必须单独记录；不能从试用小样本推断
  无限制的大文档表现，也不能把水印差异当作 DocWeave 的错误。
- 包摘要、文档抓取摘要、环境和基准采集时的 CI 记录见
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
| B1 安全与内容完整性 | 加载资源边界、失败原输出保留、恶意输入、隐藏内容、文字/图像/关系丢失；与基准逐项比较 | 支持内容无静默损失，拒绝和诊断可验证；兼容行为与安全要求冲突单列决策 | Markdown 普通文本有双方输出；DOM 保存边界已有局部验证，整体未验收 |
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
| 格式继承 | 部分直接格式；13 个字体布尔属性、普通字号、段落对齐的有效值与样式链校验 | 完整属性的未设置/继承/显式状态；字符/段落/列表/表格/主题映射；读写不能擅自物化有效值 |
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

## 复验与证据入口

使用契约与限制统一见 [DOM 指南](docx-dom.md)和[转换与安全](enhanced-conversion.md)；
逐成员实现状态与缺口以能力台账为准，下表只列复验入口。

| 材料 | 用途 |
|---|---|
| [全量能力台账](benchmarks/commercial-26.9-capabilities.json) | 当前入口、具体差距、依赖及局部行为证据；未完成项保留 |
| [加载/保存观察](benchmarks/commercial-26.9-format-behavior.json) | 固定格式方向及原始输出记录；返回文件不等于格式验收 |
| [样式编辑观察](benchmarks/style-font-edits-26.9.json) | 1482 次编辑前、编辑后、保存后及重开状态，36 次 Run 状态变化 |
| [字号词法与单位读取](benchmarks/font-size-loading-26.9.json) | 42 个值 × 直接/样式/默认值，126 个输入及 105 个官方保存/冷回读输出；独立 XML 复验。66 个正值目标 Run 字号已实现共享读取和模型往返，并经官方实际回读；三个 PDF 字号独立测量。此记录中的 22 个隐式样式默认值迁移见下项修复后的独立证据；损坏/非正数恢复、复杂文字与完整布局仍未验收 |
| [默认组与字号来源往返](benchmarks/font-default-presence-26.9.json) | 复用 87 个输入，348 个 DOCX/Flat OPC/模型 JSON 输出经独立 XML 和官方冷回读核对；缺失默认组不再成为空组，Run/样式字号与直接声明状态保留。21 个 PDF 普通字号独立测量；旧 JSON 无法恢复未记录的来源信息，完整字体来源、复杂文字和渲染仍未验收 |
| [字体 JSON 与粗体/斜体转换](benchmarks/font-json-origin-26.9.json) | 复用 745 个输入；1,490 个 JSON→DOCX/Flat OPC 输出经官方和独立 XML 回读，修复前的 194 个差异降为零。保留稀疏字段设置、Run/字符样式 b/i 直接声明，读取和写出基线复用 DOM 的类别组合规则。最终包专项 900 项、Linux 911 项、31 个示例通过。全套首次 10,318 项通过、2 项失败、2 项跳过，失败来自归档更新中的旧样本，两项在稳定归档上复验通过；未重跑整套。远端 CI 待核对；完整字体、表格上下文及布局未验收 |
| [隐式字号观察](benchmarks/font-default-matrix-26.9.json) | 16 个段落/字符样式输入，区分缺少默认组、空组及显式/复杂字号 |
| [其余字体 Boolean 读取](benchmarks/font-boolean-contexts-26.9.json) / [隐藏文字原生 PDF](benchmarks/font-hidden-rendering-26.9.json) | 11 个属性、891 个自有输入；读取类别组合及隐藏渲染修复已通过本阶段专项验收。81 个官方试用 PDF 保存了目标文字可见性；其中 `hidden-1-0-1-n` 和 `hidden-1-1-0-n` 的 getter 为真而 PDF 显示文字，证明读取值不能直接充当渲染策略。独立复验入口为 `verify_font_boolean_contexts`，生成器为 `probes/font_boolean_contexts.py`；属性、81 个 PDF 可见性及模型 JSON 往返专项已通过；新增 [32 个默认样式引用样本](benchmarks/hidden-style-contexts-26.9.json)修复显式/隐式引用的四个渲染差异。隐式引用补丁全套 11,460 项通过、2 项跳过；后续 [452 个隐藏格式往返输出](benchmarks/hidden-font-roundtrip-26.9.json)将 20 个 getter、52 个可见性差异降为零，保留直接声明和隐式引用。最终来源修复安装包专项 2,608 项、Mac 全套（含示例）11,916 项通过、2 项跳过，Linux 专项 2,504 项通过。全套收集后新增的 Windows ZIP 元数据回归与 CLI 平台断言修正另经 830 项专项复验；远端 CI 待完成。完整字体、动态样式编辑、全部格式及全页面视觉等价仍未验收 |
| [样式段落格式观察](benchmarks/style-paragraph-format-26.9.json) | 27 个继承输入、108 次基本对齐编辑及官方原始输出 |
| [段落分页属性](benchmarks/paragraph-pagination-26.9.json) | 324 个非首段输入、1296 次样式/直接值编辑，24 次非法 setter 观察 |
| [分页渲染对照](benchmarks/pagination-rendering-26.9.json) | 五个自有输入、10 个原生/当前 PDF；非首页面位置、字宽及黑色像素，完整排版仍未验收 |
| [段落数值格式](benchmarks/paragraph-dimensions-26.9.json) / [渲染](benchmarks/paragraph-dimensions-rendering-26.9.json) | 五个 pt 属性的层次继承、360 次编辑和 270 次错误观察；10 个普通字体 PDF 对照，完整格式仍未验收 |
| [段落间距边界](benchmarks/paragraph-spacing-limits-26.9.json) / [极端 setter 未决项](benchmarks/paragraph-dimension-extremes-26.9.json) | 32 次间距边界编辑及保存重开；35 次极端 live getter 观察，非有限值仍不一致；有限缩进见下项 |
| [有限缩进边界](benchmarks/paragraph-indent-limits-26.9.json) | 54 次段落/样式编辑；54 个原生输出、10 个合法 hanging 规范化对照，经官方和独立解析重开；极端渲染未验收 |
| [逻辑缩进](benchmarks/paragraph-logical-indents-26.9.json) | 42 个逻辑/首行缩进别名与继承输入、168 次样式/段落编辑和重开；bidi getter 对照不等于双向排版验收 |
| [字符缩进生命周期](benchmarks/paragraph-character-indents-26.9.json) | 8 种字号上下文分别执行直接保存和排版后保存，另有 45 次字符 setter、27 次类型错误；原生 DOCX 独立读取；只读字符 getter 已实现，setter、点值联动与渲染仍未验收 |
| [字符 setter 字体与样式上下文](benchmarks/paragraph-character-setters-26.9.json) | 8 种字号 × 段落/样式 × 3 个属性 × 零/正/负值，144 次编辑及 144 次 TypeError；保留加载、编辑、保存后的活动对象和冷回读，独立 XML 核对点值及字符值。样式 setter 的点值在保存时才更新；DOM setter 和渲染仍未验收 |
| [字符缩进继承读取](benchmarks/paragraph-character-reads-26.9.json) | 24 组 docDefaults/basedOn/直接值输入，包含显式零；原包 DOM 支持直接字符 getter、样式继承 getter，205 个官方保存输出可读取。setter、点值联动、编号/条件表格继承及渲染未验收 |
| [字符缩进继承编辑](benchmarks/paragraph-character-inheritance-edits-26.9.json) | 24 个相同输入在 Base/Derived/段落上执行零/正/负值 setter，216 次编辑、216 次 TypeError；六个字符/点值 getter 的加载、编辑、保存活动对象和冷回读均保留，冷回读的继承链用独立 XML 验证。仅为官方行为证据，SDK setter 和布局仍未验收 |
| [损坏 twip 输入拒绝](../tests/test_twip_input_validation.py) | 段落缩进/间距、编号定义/覆盖和 dxa 表格/单元格边距拒绝非法整数及溢出，异常隐藏原值，保留原包和已有 CLI 输出；正常值及整数差值运算复验。官方损坏 twip 恢复行为未测；字号读取及不同的损坏值恢复见上述字号观察 |
| [字符单位模型往返](benchmarks/paragraph-character-indents-current.json) | 已安装 wheel 经模型 JSON 写出 61 组 DOCX/Flat OPC，122 个输出按原生字符值独立读取；样式、编号及显式零保留，点值联动与渲染仍未验收 |
| [首段分页未决差异](benchmarks/first-paragraph-page-break-26.9.json) | 4 次编辑输出及 [18 个试用对照](benchmarks/first-paragraph-trial-26.9.json)；首段变化伴随提示插入，持许可证对照仍缺失 |
| [保存状态观察](benchmarks/style-save-state-26.9.json) | 商业试用包的内存/重开状态，区分水印内容 |
| [样式导入投影](benchmarks/style-import-projections.json) | 当前导入的基准、输入输出 SHA 与独立格式读取 |
| [Flat OPC 加载](benchmarks/flat-opc-loading.json) / [保存](benchmarks/flat-opc-saving.json) | 四种受测变体；复杂宏、模板与布局仍未全面验收 |
| [保存选项审计](enhanced-conversion.md#保存选项逐字段审计) | 公开字段的实现路径、限制及验证入口 |
| [生态借鉴](ecosystem-adoption.md) | Python 同类项目的差别与吸收方向 |

执行 `python scripts/verify_commercial_baseline.py` 检查记录和产物摘要；执行
`tests/test_commercial_api_inventory.py` 检查台账与伪造证据拒绝行为。固定输入和输出保留在
`benchmarks/corpus/`，行为生成器在 `probes/`，被引用的专项记录保留在 `benchmarks/`。
回归测试、复验器和指南所需证据继续保留；无引用的旧阶段日志与一次性实验记录可从 Git 历史恢复。

安装包测试需在源码目录外运行，并检查实际导入位置、字体、许可和 typing 资源；
方法见 [开发与安装核验](../README.zh-CN.md#开发与测试)。具体提交的跨平台结果以
[GitHub Actions](https://github.com/shenyankm/DocWeave/actions) 为准，不能用旧 SHA 的成功代替。

其余十个 Boolean 属性的保存来源进入当前阶段：复用 810 个自有输入，
[3,240 个直接／模型 JSON→DOCX／Flat OPC 输出](benchmarks/font-boolean-roundtrip-26.9.json)
经官方 26.9.0 冷回读，getter 差异从 200 个降为零。前后原始输出与生成器已冻结；
独立 XML 复验检查 getter、直接声明和段落样式引用。读取、赋值、模型复制及写出
共享来源规则，源代码专项 4,083 项通过。最终安装包 Mac 全套（含示例）15,212 项通过、
2 项跳过，Linux 专项 5,798 项通过。全套结束后，归档以 `storage` 复用逐字节相同的
输出：总量从 21,256,900 字节降为 6,873,096 字节，6,480 个逻辑输出的原始字节不变；
归档与新增别名检查另经最终安装包 3,411 项专项复验通过。本阶段本地验收通过，远端 CI 待完成。
完整 Font API、动态样式编辑、其他格式与实际字体效果仍未验收。

此前 Windows CI 的 ZIP 元数据和 CLI 平台断言已修正；提交 `2bcecca` 的
[四个平台检查](https://github.com/shenyankm/DocWeave/actions/runs/38005138101)均通过。
该结果不能代替后续字体阶段提交的远端检查。

可编辑 DOM 的字体布尔阶段补齐 Run 直接属性、样式读取/赋值、有效字体与文本范围格式入口，
复用已有验证和结构编辑保护。新增 11 个属性的[固定编辑观测](benchmarks/dom-font-booleans-26.9.json)
覆盖 891 个输入、5,346 组官方 live 读取/赋值，以及 10,692 组当前输出的官方
DOCX／Flat OPC 冷读取，getter 差异均为零。原始输出和完整记录压缩在同一 2.7 MB 归档；
独立复验还检查输入/输出哈希、直接声明及段落引用。源代码与既有样式/范围集成测试
7,686 项通过；属性实现安装包（错误类型修正前）Mac 全套含示例 20,600 项通过、
2 项跳过，Linux amd64 专项 7,802 项通过。最终错误类型修正后的安装包仅变更
共享验证入口的异常类，Mac／Linux 各 9,472 项受影响调用路径与集成测试通过；
最终包生产模块与源代码逐字节一致，资源检查、全量证据复验、114 个本地文档链接通过。
本阶段受测 DOCX／Flat OPC getter、编辑、声明保存及非法类型错误的本地验收通过；远端 CI 待完成。
basedOn 链、表格/列表字体、跨文档样式翻译、
渲染与完整 Font API 仍为必需的未验收范围，不因当前 getter 一致而视为通过。

新增[297 组非法赋值观测](benchmarks/dom-font-boolean-errors-26.9.json)确认官方对
数字、字符串、容器及 `None` 均抛出 `TypeError`。此前 DOM 对非布尔值抛出
`ValueError`，264 组错误类型差异已在共享验证入口修正；Run 的直接字体允许 `None` 清除声明，另有
11 组接口语义差异。后者是当前直接格式接口的明确功能，不能把它描述为商业 Font
接口已对齐。完整 Font 接口的继承/直接格式契约仍待统一。

下一项数据安全修复：最终安装包已复现 `font.color = "FF0000"` 在缺少 descriptor 时
只写入临时 Python 对象，DOCX 字节完全不变。未实现的格式属性必须明确拒绝赋值，
不能静默制造编辑成功的假象；随后再逐项补齐颜色、字体选择等公开属性及保存/渲染验证。

格式 getter/setter 的固定观察见上表，不能替代最终渲染验收。首段复验生成器为
[`first_paragraph_trial.py`](probes/first_paragraph_trial.py)：15 个加载输入及 3 个新建/保存对照，
覆盖提示段落出现前后的值。官方[许可说明](https://docs.aspose.com/words/python-net/licensing/)确认
试用水印在加载和保存时加入，但未规定本次观察到的属性变化；持许可证结果仍未确认。
历史测试成绩和修复过程从 Git 历史查询，不作为当前工作树或后续提交的通过证明。
