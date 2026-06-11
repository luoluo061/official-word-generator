# Word Feature Catalog

This catalog is the shared capability pool for `official-word-generator`.

Use it before creating or changing a document profile. A profile should enable, disable, or mark features as overridable by referencing `feature_id`; do not hard-code feature decisions directly in `SKILL.md`.

Status values:

- `implemented`: supported by the current scripts/templates.
- `partial`: partly supported, but needs manual profile rules, template support, Word COM, or user verification.
- `planned`: design target only; do not claim support until implemented and validated.

## 1. 页面设置类

| feature_id | 中文名称 | 功能说明 | 默认启用 | 适用文档类型 | 依赖 | 可覆盖 | 验证方式 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `page.a4_portrait` | A4 纵向页面 | 设置或校验 A4 纵向纸张。 | 是 | 全部正式文档 | 无 | 是 | 读取 section page width/height | implemented |
| `page.official_margins` | 公文版心页边距 | 上 37mm、下 35mm、左 27mm、右 27mm。 | 一般公文默认 | 一般公文、工作汇报、会议纪要 | `page.a4_portrait` | 是 | 校验 section margins | implemented |
| `page.profile_margins` | 文种自定义页边距 | 由 profile 定义页面边距，例如合同或软著说明书可不同。 | 否 | 软著说明书、合同、项目材料 | `page.a4_portrait` | 是 | profile validation rules + section 检查 | planned |
| `page.section_breaks` | 分节控制 | 封面、目录、正文、附件使用不同节。 | 否 | 软著说明书、合同、项目申报书 | 无 | 是 | 检查 section 数量和 section 属性 | planned |

## 2. 字体段落类

| feature_id | 中文名称 | 功能说明 | 默认启用 | 适用文档类型 | 依赖 | 可覆盖 | 验证方式 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `paragraph.body_style` | 正文样式 | 正文使用 profile 的字体、字号、对齐、行距。 | 是 | 全部正式文档 | 无 | 是 | 检查 Normal 或 profile body style | implemented |
| `paragraph.first_line_chars` | 字符单位首行缩进 | 使用 `w:firstLineChars=200`，避免厘米缩进漂移。 | 一般公文默认 | 一般公文、工作汇报、会议纪要 | `paragraph.body_style` | 是 | OOXML 检查 firstLineChars | implemented |
| `paragraph.fixed_line_spacing` | 固定行距 | 正文/标题按 profile 配置固定行距。 | 是 | 公文、报告、说明书 | `paragraph.body_style` | 是 | 检查 spacing lineRule/line | implemented |
| `font.east_asia_ascii_split` | 中西文字体分离 | 中文字体与数字/英文 Times New Roman 分离设置。 | 是 | 中文正式文档 | 无 | 是 | 检查 rFonts eastAsia/ascii | implemented |
| `font.profile_style_rules` | 文种样式规则 | 从 profile format rules 定义不同文种的样式。 | 否 | 全部 profile | `paragraph.body_style` | 是 | 生成后样式检查 | partial |

## 3. 标题层级类

| feature_id | 中文名称 | 功能说明 | 默认启用 | 适用文档类型 | 依赖 | 可覆盖 | 验证方式 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `heading.level_1` | 一级标题 | Markdown `##` 映射到一级标题样式和大纲 1 级。 | 是 | 全部正式文档 | 无 | 是 | 样式名 + outline level | implemented |
| `heading.level_2` | 二级标题 | Markdown `###` 映射到二级标题样式和大纲 2 级。 | 是 | 全部正式文档 | `heading.level_1` | 是 | 样式名 + outline level | implemented |
| `heading.level_3` | 三级标题 | 支持三级标题样式和目录级别。 | 否 | 软著说明书、项目申报书、可研报告 | `heading.level_2` | 是 | 样式名 + outline level | planned |
| `heading.level_4` | 四级标题 | 支持四级标题样式。 | 否 | 技术说明书、长报告 | `heading.level_3` | 是 | 样式名 + outline level | planned |
| `heading.profile_numbering` | 文种标题编号规则 | 一、（一）、1.、（1）等结构识别或规范化。 | 否 | 软著说明书、合同、项目材料 | 标题层级 | 是 | Markdown 结构检查 | partial |

## 4. 目录类

| feature_id | 中文名称 | 功能说明 | 默认启用 | 适用文档类型 | 依赖 | 可覆盖 | 验证方式 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `toc.insert_field` | 插入 Word 目录域 | Markdown `[TOC]` 插入 Word TOC 域。 | 否 | 软著说明书、长报告、项目材料 | `heading.level_1` | 是 | OOXML 检查 TOC field | implemented |
| `toc.level_1_2` | 目录到二级标题 | TOC 使用 `\o "1-2"`。 | 否 | 当前基础模板、一般长文 | `toc.insert_field` | 是 | 检查 instrText | implemented |
| `toc.level_profile` | 文种目录层级 | 由 profile 指定目录到 2、3 或 4 级。 | 否 | 软著说明书、项目申报书 | `toc.insert_field` | 是 | 检查 instrText | planned |
| `toc.update_with_word` | Word 更新域 | 使用 Microsoft Word/pywin32 刷新目录页码。 | 否 | 有目录的文档 | `toc.insert_field` | 是 | Word COM 成功返回 + 重新校验 | partial |
| `toc.page_breaks` | 封面目录正文分页 | TOC 前后插入分页符。 | 有 `[TOC]` 时启用 | 软著说明书、长报告 | `toc.insert_field` | 是 | 检查分页符位置 | implemented |

## 5. 页眉页脚页码类

| feature_id | 中文名称 | 功能说明 | 默认启用 | 适用文档类型 | 依赖 | 可覆盖 | 验证方式 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `page_number.odd_even_official` | 公文奇偶页页码 | 奇数页右侧、偶数页左侧，格式类似 `- 1 -`。 | 一般公文默认 | 一般公文、工作汇报、会议纪要 | `page.a4_portrait` | 是 | footer XML 检查 left/right/page field | implemented |
| `page_number.centered` | 居中页码 | 页脚居中页码。 | 否 | 合同、说明书、普通报告 | 无 | 是 | footer XML 检查 center/page field | planned |
| `header.profile_text` | 页眉文字 | 按 profile 插入固定页眉或文档名。 | 否 | 合同、软著说明书、报告 | 无 | 是 | header XML 检查 | planned |
| `footer.profile_text` | 页脚文字 | 按 profile 插入公司名、密级、版本等。 | 否 | 合同、项目材料 | 无 | 是 | footer XML 检查 | planned |

## 6. 表格类

| feature_id | 中文名称 | 功能说明 | 默认启用 | 适用文档类型 | 依赖 | 可覆盖 | 验证方式 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `table.markdown_to_word` | Markdown 表格转 Word | 将 Markdown 表格生成 Word table。 | 是 | 全部文档 | 无 | 是 | 表格数量对比 | implemented |
| `table.three_line` | 三线表 | 顶线、表头底线、底线，不使用完整网格线。 | 是 | 报告、说明书、测算文档 | `table.markdown_to_word` | 是 | tblBorders 检查 | implemented |
| `table.header_repeat` | 表头跨页重复 | 首行设置 `w:tblHeader`。 | 是 | 长表格文档 | `table.markdown_to_word` | 是 | OOXML 检查 tblHeader | implemented |
| `table.cell_style` | 表格文字样式 | 表头和表体使用独立样式，不继承正文首行缩进。 | 是 | 全部含表文档 | `table.markdown_to_word` | 是 | 段落样式检查 | implemented |
| `table.autofit_content` | 根据内容自适应列宽 | 表格宽度和列宽根据内容/页面自动调整。 | 是 | 全部含表文档 | `table.markdown_to_word` | 是 | 表格布局检查 + 肉眼复核 | partial |
| `table.profile_borders` | 文种表格边框规则 | 合同/清单/报价等可使用网格表或自定义边框。 | 否 | 合同、报价、清单 | `table.markdown_to_word` | 是 | tblBorders 检查 | planned |

## 7. 图片与图题类

| feature_id | 中文名称 | 功能说明 | 默认启用 | 适用文档类型 | 依赖 | 可覆盖 | 验证方式 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `image.insert_inline` | 插入嵌入型图片 | Markdown 图片插入为 Word 内嵌图片。 | 是 | 说明书、报告 | 无 | 是 | 检查 drawing 元素和缺图警告 | implemented |
| `image.center` | 图片居中 | 图片段落居中、单倍行距。 | 是 | 说明书、报告 | `image.insert_inline` | 是 | 段落对齐检查 | implemented |
| `caption.figure` | 图题 | 图片下方生成图题样式。 | 是 | 说明书、报告 | `image.insert_inline` | 是 | 图题样式检查 | implemented |
| `image.profile_size` | 图片尺寸规则 | profile 定义图片最大宽度、缩放规则。 | 否 | 软著说明书、验收报告 | `image.insert_inline` | 是 | drawing 尺寸检查 | partial |
| `caption.numbering` | 图表自动编号 | 自动生成“图 1”“表 1”等编号。 | 否 | 说明书、报告 | caption features | 是 | 字段/文本编号检查 | planned |

## 8. 文种结构类

| feature_id | 中文名称 | 功能说明 | 默认启用 | 适用文档类型 | 依赖 | 可覆盖 | 验证方式 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `structure.cover` | 封面 | 文档封面页结构。 | 否 | 软著说明书、项目申报书、报告 | `page.section_breaks` 可选 | 是 | 检查封面字段/分页 | planned |
| `structure.toc_page` | 目录页 | 目录单独成页。 | 否 | 软著说明书、长报告 | `toc.insert_field` | 是 | 检查分页符/TOC | partial |
| `structure.body_start_page` | 正文另起页 | 封面/目录后正文从新页开始。 | 否 | 软著说明书、长报告 | `toc.page_breaks` 或 `page.section_breaks` | 是 | 检查分页符 | implemented |
| `structure.signature_block` | 签署区 | 合同、报告结尾签章/日期区域。 | 否 | 合同、汇报、会议纪要 | 无 | 是 | 检查固定结构段落 | planned |
| `structure.appendix` | 附件 | 支持附件标题、附件正文和分页。 | 否 | 公文、合同、项目材料 | 无 | 是 | 检查附件样式/分页 | planned |
| `structure.red_head` | 红头文件 | 红头、发文字号、签发人、版记等。 | 否 | 红头文件 | 页眉/版记/分隔线 | 是 | 模板字段 + 肉眼复核 | planned |

## 9. 内容整理类

| feature_id | 中文名称 | 功能说明 | 默认启用 | 适用文档类型 | 依赖 | 可覆盖 | 验证方式 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `content.normalize_text` | 文本编码和换行清洗 | 处理 pasted text、txt、md、content docx，输出 UTF-8 Markdown。 | 是 | 全部文档 | 无 | 否 | `prepare_content.py` 输出检查 | implemented |
| `content.docx_to_markdown` | 内容 Word 转 Markdown | 从内容型 docx 抽取标题、段落、表格为 Markdown。 | 否 | 迁移旧文档 | `content.normalize_text` | 是 | Markdown 结构复核 | partial |
| `content.markdown_structure` | Markdown 结构规范 | 要求用 `#`、`##`、`###`、表格、图片表达结构。 | 是 | 全部文档 | 无 | 是 | parser block 检查 | implemented |
| `content.profile_outline` | 文种结构大纲 | 根据文种提供默认章节结构。 | 否 | 软著说明书、合同、工作汇报 | profile 配置 | 是 | Markdown heading 检查 | planned |
| `content.writing_style` | 文风规则 | 按 profile 或用户文风生成/改写内容。 | 否 | 可选写作流程 | profile 文风规则 | 是 | 人工审阅 | planned |

## 10. 校验类

| feature_id | 中文名称 | 功能说明 | 默认启用 | 适用文档类型 | 依赖 | 可覆盖 | 验证方式 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `validation.basic_open` | 基础打开校验 | `.docx` 可被 python-docx 打开。 | 是 | 全部文档 | 无 | 否 | `validate_docx.py` | implemented |
| `validation.required_styles` | 必要样式校验 | 检查 profile/template 必需样式是否存在。 | 是 | 全部文档 | profile rules | 是 | 样式集合对比 | implemented |
| `validation.page_setup` | 页面设置校验 | 检查纸张、页边距、页眉页脚距离。 | 是 | 一般公文默认 | profile rules | 是 | section 属性 | implemented |
| `validation.indent_chars` | 字符缩进校验 | 禁止非 TOC 段落使用厘米/twip 首行缩进。 | 是 | 公文、正式中文文档 | `paragraph.first_line_chars` | 是 | OOXML 检查 | implemented |
| `validation.footer_page_number` | 页码位置校验 | 检查奇偶页页码或 profile 页码规则。 | 是 | 一般公文默认 | page number features | 是 | footer XML | implemented |
| `validation.table_rules` | 表格规则校验 | 检查三线表、跨页表头。 | 是 | 含表文档 | table features | 是 | OOXML 表格检查 | implemented |
| `validation.profile_rules` | profile 专属校验 | 根据 profile 的 validation_rules.json 动态校验。 | 否 | 多文种 | profile 配置 | 是 | profile validation engine | planned |
| `validation.visual_review` | 人工/截图复核 | 对 Word 版式进行肉眼检查。 | 否 | 重要交付文档 | Word/WPS | 否 | 打开文档或导出 PDF | partial |
