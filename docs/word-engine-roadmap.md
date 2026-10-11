# DocKiln：完整 Word DOM 与文档渲染实施路线图

> 此文档是**待核验的能力清单**：未勾选不代表当前没有实现，开发前必须核对现有代码和测试。优先复用当前 dockiln 的 DOCX 编辑、提取与 PDF 功能。

## 目标

建立 OPC/OOXML → 可编辑 DOM → 样式系统 → 独立 Layout Tree → PDF/PNG/SVG Renderer 的渐进式架构。保留既有 Document 和 DocxDocument API，避免破坏兼容性。

优先级：P0（基础）、P1（增强）、P2（高级兼容）。每项完成后记录 PR、测试及样本证据。

## 完整任务清单

### 1. OPC / OOXML

- [ ] **[P0]** ZIP、Content Types、Relationships 安全解析
- [ ] **[P0]** Document、Styles、Numbering、Settings、Theme 与媒体关系
- [ ] **[P0]** 页眉页脚、脚注尾注、批注及未知部件无损保留
- [ ] **[P1]** Strict OOXML 和旧 DOC 兼容

### 2. 完整 DOM

- [ ] **[P0]** Node/CompositeNode/NodeType 统一抽象
- [ ] **[P0]** Document、Section、Body、Paragraph、Run、Text
- [ ] **[P0]** Table/Row/Cell、嵌套表格
- [ ] **[P1]** 父子兄弟导航、递归查询、增删改查与复制
- [ ] **[P1]** Hyperlink、Bookmark、Field、Comment、Footnote、Endnote
- [ ] **[P2]** Revision、Office Math、Shape、Drawing

### 3. 编辑和保存

- [ ] **[P0]** DOM 与 OOXML 双向映射及序列化
- [ ] **[P0]** 跨 Run 文本替换与样式编辑
- [ ] **[P0]** 表格合并、插入与删除
- [ ] **[P1]** 图片、书签与超链接关系管理
- [ ] **[P1]** 跨文档节点导入、合并与拆分
- [ ] **[P2]** 未改动 Part 保真和 round-trip 测试

### 4. 样式和字体

- [ ] **[P0]** 样式继承、直接格式与默认样式优先级
- [ ] **[P0]** 段落、字符、表格样式
- [ ] **[P0]** 列表/多级编号、主题颜色字体
- [ ] **[P1]** 中文 Unicode 字体回退与复杂脚本塑形
- [ ] **[P1]** 行距、制表位、边框、底纹和兼容模式

### 5. 布局引擎

- [ ] **[P0]** 独立 Page/Line/Box 布局树
- [ ] **[P0]** 纸张、页边距、分节和分页
- [ ] **[P0]** 文字度量、断行、段落行盒
- [ ] **[P1]** 表格列宽、行高和跨页拆分
- [ ] **[P1]** 孤行寡行与 keep-with-next
- [ ] **[P2]** 页眉页脚、动态页码和脚注排版
- [ ] **[P2]** 图片浮动、文字环绕、分栏和双向文本
- [ ] **[P2]** 增量布局缓存

### 6. 渲染

- [ ] **[P0]** 统一 Canvas 绘图接口
- [ ] **[P0]** PDF 与 PNG 页面渲染
- [ ] **[P0]** SVG 页面输出
- [ ] **[P1]** 文本/表格/图片/边框绘制与字体嵌入
- [ ] **[P1]** PDF 链接、书签及页级坐标
- [ ] **[P2]** DrawingML、SmartArt 和图表支持

### 7. 提取与查询

- [ ] **[P0]** 正文段落和表格自然阅读顺序
- [ ] **[P0]** 结构化 JSON 与来源 XML 路径
- [ ] **[P0]** 图片、公式、批注及脚注提取
- [ ] **[P1]** Markdown/HTML 语义导出和丢失诊断
- [ ] **[P1]** 节点页面坐标和 bounding box 查询

### 8. 测试/安全/性能

- [ ] **[P0]** 单测、集成、往返、属性测试
- [ ] **[P0]** 未知部件与关系的保真测试
- [ ] **[P0]** ZIP Bomb、XXE、路径穿越和外链防护
- [ ] **[P1]** Word/LibreOffice 参考布局及视觉回归
- [ ] **[P1]** 跨平台字体缺失测试
- [ ] **[P2]** 10/100/1000 页耗时和内存基准
- [ ] **[P2]** API 文档及兼容性矩阵

## 里程碑

| 阶段 | 交付 | 验收 |
|---|---|---|
| M1 | 已有能力审计和缺口矩阵 | 已实现/部分实现/待实现，有代码证据 |
| M2 | DOM 抽象与编辑能力 | 节点导航、修改、克隆、导入测试 |
| M3 | 样式与基础布局树 | 基础段落及表格分页样本 |
| M4 | PDF/PNG/SVG 统一渲染 | 视觉回归与页级输出 |
| M5 | 高级 Word 兼容 | 复杂表格、浮动对象、性能与安全基线 |

## 建议模块边界

```text
src/dockiln/
  opc/        # Part/Relationship
  ooxml/      # XML reader/writer
  dom/        # Word 节点与编辑 API
  styles/     # 样式及字体
  layout/     # 页面、行盒、分页
  render/     # PDF/PNG/SVG
  extraction/ # 内容提取及来源映射
```

该目录是建议，不要求立即重构现有项目。

## 验收要求

- 明确区分结构保真、包部件保真、视觉保真，不能仅凭文档可打开判定完成。
- 指定 Word/LibreOffice 版本、字体及系统环境建立视觉基线。
- 对不可信 ZIP/XML、资源耗尽、外部关系、路径穿越做安全测试。
- 建立 10/100/1000 页耗时和峰值内存基准。

## 现有项目资料

- [DOM 编辑说明](docx-dom.md)
- [转换与安全](enhanced-conversion.md)
- [商业功能对齐计划](commercial-alignment-plan.md)
- [生态对比](ecosystem-adoption.md)
- [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
