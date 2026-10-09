# 公开保存选项审计入口

范围为 `saving.py` 的三个保存选项类及其 `OutlineOptions`：40 + 6 个初始化公开字段，其中 15 个字段使用枚举描述符。
下表记录源码行为与复验入口，不表示本轮已完整验收每项；全矩阵效果、标量边界和跨格式组合归于 A1/A4。
本轮验收未知枚举拒绝与三项有效但未实现请求的诊断，见[记录](benchmarks/save-option-contract.json)。
支持枚举值不等于对应功能已实现：有效但未实现值仍按表中限制处理。

## PdfSaveOptions

| 字段 | 当前行为/限制 | 证据或复验入口 |
|---|---|---|
| `compliance` | 只设置 PDF 版本；PDF/A、PDF/UA 请求警告，非标准认证 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `export_document_structure` | 结构输出开关；完整辅助技术/标准验收仍待决策 D02 | [test_pdf_structure_pages.py](../tests/test_pdf_structure_pages.py) |
| `image_compression` | AUTO/JPEG 图片转换路径；实际损失与透明图像需按样本验收 | [ApiExamples/working_with_pdf_save_options.py](../ApiExamples/working_with_pdf_save_options.py) |
| `jpeg_quality` | JPEG 质量参数；全部边界值尚未系统验收 | [ApiExamples/working_with_pdf_save_options.py](../ApiExamples/working_with_pdf_save_options.py) |
| `text_compression` | NONE/FLATE 页面内容流，字体/图片流独立 | [test_pdf_stream_options.py](../tests/test_pdf_stream_options.py) |
| `embed_full_fonts` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `use_core_fonts` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `font_embedding_mode` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `page_mode` | 写入 PDF PageMode；阅读器可以忽略 | [test_pdf_stream_options.py](../tests/test_pdf_stream_options.py) |
| `color_mode` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `export_bookmarks_outline` | 书签输出；显式 outline 层级优先 | [test_pdf_outline_options.py](../tests/test_pdf_outline_options.py) |
| `outline_options` | 下表六项；不是完整 Word TOC 契约 | [test_pdf_outline_options.py](../tests/test_pdf_outline_options.py) |
| `preserve_form_fields` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `memory_optimization` | 未实现；显式赋值有 PDF unsupported_option 警告 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `zoom_factor` | ZOOM_FACTOR 下使用；阅读器行为与全部边界尚待复核 | [aspose/words_foss/pdf_writer/writer.py](../aspose/words_foss/pdf_writer/writer.py) |
| `zoom_behavior` | PDF 打开动作；部分行为在序列化后改写 | [aspose/words_foss/pdf_writer/writer.py](../aspose/words_foss/pdf_writer/writer.py) |
| `display_doc_title` | PDF viewer preference；阅读器可忽略 | [aspose/words_foss/pdf_writer/writer.py](../aspose/words_foss/pdf_writer/writer.py) |
| `fallback_fonts` | 可信字体路径；覆盖、塑形及源字体替代有独立限制 | [test_pdf_diagnostics.py](../tests/test_pdf_diagnostics.py) |
| `text_shaping` | 可选 HarfBuzz 塑形；依赖/字体及格式边界有限制 | [test_pdf_shaping.py](../tests/test_pdf_shaping.py) |

## OoxmlSaveOptions

| 字段 | 当前行为/限制 | 证据或复验入口 |
|---|---|---|
| `save_format` | 只支持 DOCX；其他值在输出前拒绝 | [test_save_option_formats.py](../tests/test_save_option_formats.py) |
| `compliance` | Transitional/ECMA 路径；STRICT 明确拒绝 | [test_bounded_conversion.py](../tests/test_bounded_conversion.py) |
| `compression_level` | 四档 ZIP deflate；不是内容精简 | [ApiExamples/working_with_ooxml_save_options.py](../ApiExamples/working_with_ooxml_save_options.py) |
| `zip_64_mode` | NEVER/IF_NECESSARY；ALWAYS 仍等同按需，但现在警告 | [test_unsupported_save_requests.py](../tests/test_unsupported_save_requests.py) |
| `pretty_format` | XML 缩进；语义与全部部件仍待矩阵复核 | [ApiExamples/working_with_ooxml_save_options.py](../ApiExamples/working_with_ooxml_save_options.py) |
| `reference_docx` | 导入参考样式，非参考正文/资源/页面设置 | [test_reference_styles.py](../tests/test_reference_styles.py) |

## MarkdownSaveOptions

| 字段 | 当前行为/限制 | 证据或复验入口 |
|---|---|---|
| `table_content_alignment` | 表格对齐覆盖 | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py) |
| `list_export_mode` | Markdown 列表或纯文本路径 | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py) |
| `export_images_as_base64` | 无 images_folder 时始终内联；有目录时选择内联/外部 | [test_conversion_api.py](../tests/test_conversion_api.py) |
| `images_folder` | 同名不同内容分配序号，已有文件及目标符号链接保留；普通失败回滚新图片，占用标记保护待提交图片；非断电安全多文件事务 | [test_markdown_image_rollback.py](../tests/test_markdown_image_rollback.py)、[test_markdown_image_collisions.py](../tests/test_markdown_image_collisions.py) |
| `images_folder_alias` | 外部图片链接目录别名 | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py) |
| `export_underline_formatting` | 下划线 HTML 标记 | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py) |
| `link_export_mode` | 自动/内联/引用链接 | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py) |
| `export_as_html` | TABLES 走已有 HTML 路径；NON_COMPATIBLE_TABLES 仍等同 NONE，但现在警告 | [test_unsupported_save_requests.py](../tests/test_unsupported_save_requests.py) |
| `empty_paragraph_export_mode` | 空行、HTML br 或省略 | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py) |
| `image_resolution` | 未重采样；非默认 96 请求现在警告，源图片字节保留 | [test_unsupported_save_requests.py](../tests/test_unsupported_save_requests.py) |
| `save_format` | 只支持 MARKDOWN；其他值在输出前拒绝 | [test_save_option_formats.py](../tests/test_save_option_formats.py) |
| `encoding` | 文本编码；编码失败在主输出写入前发生 | [test_conversion_api.py](../tests/test_conversion_api.py) |
| `paragraph_break` | 顶层块分隔；内部换行仍为 LF | [ApiExamples/working_with_markdown_save_options.py](../ApiExamples/working_with_markdown_save_options.py) |
| `style_map` | 精确样式名映射为支持的语义样式 | [test_ecosystem_semantics.py](../tests/test_ecosystem_semantics.py) |
| `export_notes` | 可见注锚点、定义与回链；缺失/不支持位置有诊断 | [test_markdown_notes.py](../tests/test_markdown_notes.py) |

## OutlineOptions

| 字段 | 当前行为/限制 | 证据或复验入口 |
|---|---|---|
| `headings_outline_levels` | 标题深度；大于六的值按现有规则夹紧 | [test_pdf_outline_options.py](../tests/test_pdf_outline_options.py) |
| `expanded_outline_levels` | 大纲展开范围 0..9 | [test_pdf_outline_options.py](../tests/test_pdf_outline_options.py) |
| `default_bookmarks_outline_level` | 默认书签层级；名称过滤与补层规则有限制 | [test_pdf_outline_options.py](../tests/test_pdf_outline_options.py) |
| `bookmarks_outline_levels` | 按名称指定书签层级 | [test_pdf_outline_options.py](../tests/test_pdf_outline_options.py) |
| `create_outlines_for_headings_in_tables` | 表格标题输出导航 | [test_pdf_outline_options.py](../tests/test_pdf_outline_options.py) |
| `create_missing_outline_levels` | 缺层导航补层 | [test_pdf_outline_options.py](../tests/test_pdf_outline_options.py) |

## 尚需复核

- 标量参数范围、显式默认值和跨选项组合必须验证实际输出，不能仅以属性可以赋值作为通过。
- HTML 合并跨度、格式标签与嵌套顺序已修复，见[复验记录](benchmarks/html-table-integrity.json)；字段动作/目标、语义注锚点、非图片形状和 Word 完整样式/布局仍未支持，相关损失明确诊断。不能把这些局部样本或 TABLES 枚举有效视为全语义验收通过。
- `models.ConversionOptions` 是直接 writer 的额外接口，五个共享枚举已验证；其余字段不属于上述 46 项，仍需单独核对，不能宣称全部选项已通过。
- 图片副产物普通异常回滚、已有文件保留和受控线程并发已验证，见[记录](benchmarks/markdown-image-rollback.json)；本机独立进程占用、失败与强制终止后的再次导出见[进程复验](benchmarks/markdown-image-processes.json)。中断仍有残留；自动恢复、断电及网络文件系统整体发布/恢复未验收。
- ZIP64 强制输出、选择性 HTML 表格及图片重采样本轮未实现；新诊断说明现有退回路径，不计为功能完成。
