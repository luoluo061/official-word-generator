# Feature Selection: ????

- profile_id: `general_official`
- purpose: review and edit this checklist before finalizing `features.json`.
- note: implemented features may be enabled; partial features require review; planned features are listed for future planning and should not be enabled for production.

## 1. 页面设置类

- [x] `page.a4_portrait`：A4 纵向页面（implemented）
  - 说明：设置或校验 A4 纵向纸张。
  - 默认启用：是；可覆盖：是；适用：全部正式文档；依赖：无
- [x] `page.official_margins`：公文版心页边距（implemented）
  - 说明：上 37mm、下 35mm、左 27mm、右 27mm。
  - 默认启用：一般公文默认；可覆盖：是；适用：一般公文、工作汇报、会议纪要；依赖：page.a4_portrait
- [ ] `page.profile_margins`：文种自定义页边距（planned） - planned，不可作为 production 能力
  - 说明：由 profile 定义页面边距，例如合同或软著说明书可不同。
  - 默认启用：否；可覆盖：是；适用：软著说明书、合同、项目材料；依赖：page.a4_portrait
- [ ] `page.section_breaks`：分节控制（planned） - planned，不可作为 production 能力
  - 说明：封面、目录、正文、附件使用不同节。
  - 默认启用：否；可覆盖：是；适用：软著说明书、合同、项目申报书；依赖：无
## 2. 字体段落类

- [x] `paragraph.body_style`：正文样式（implemented）
  - 说明：正文使用 profile 的字体、字号、对齐、行距。
  - 默认启用：是；可覆盖：是；适用：全部正式文档；依赖：无
- [x] `paragraph.first_line_chars`：字符单位首行缩进（implemented）
  - 说明：使用 w:firstLineChars=200，避免厘米缩进漂移。
  - 默认启用：一般公文默认；可覆盖：是；适用：一般公文、工作汇报、会议纪要；依赖：paragraph.body_style
- [x] `paragraph.fixed_line_spacing`：固定行距（implemented）
  - 说明：正文/标题按 profile 配置固定行距。
  - 默认启用：是；可覆盖：是；适用：公文、报告、说明书；依赖：paragraph.body_style
- [x] `font.east_asia_ascii_split`：中西文字体分离（implemented）
  - 说明：中文字体与数字/英文 Times New Roman 分离设置。
  - 默认启用：是；可覆盖：是；适用：中文正式文档；依赖：无
- [ ] `font.profile_style_rules`：文种样式规则（partial） - 可启用但需复核
  - 说明：从 profile format rules 定义不同文种的样式。
  - 默认启用：否；可覆盖：是；适用：全部 profile；依赖：paragraph.body_style
## 3. 标题层级类

- [x] `heading.level_1`：一级标题（implemented）
  - 说明：Markdown ## 映射到一级标题样式和大纲 1 级。
  - 默认启用：是；可覆盖：是；适用：全部正式文档；依赖：无
- [x] `heading.level_2`：二级标题（implemented）
  - 说明：Markdown ### 映射到二级标题样式和大纲 2 级。
  - 默认启用：是；可覆盖：是；适用：全部正式文档；依赖：heading.level_1
- [ ] `heading.level_3`：三级标题（planned） - planned，不可作为 production 能力
  - 说明：支持三级标题样式和目录级别。
  - 默认启用：否；可覆盖：是；适用：软著说明书、项目申报书、可研报告；依赖：heading.level_2
- [ ] `heading.level_4`：四级标题（planned） - planned，不可作为 production 能力
  - 说明：支持四级标题样式。
  - 默认启用：否；可覆盖：是；适用：技术说明书、长报告；依赖：heading.level_3
- [ ] `heading.profile_numbering`：文种标题编号规则（partial） - 可启用但需复核
  - 说明：一、（一）、1.、（1）等结构识别或规范化。
  - 默认启用：否；可覆盖：是；适用：软著说明书、合同、项目材料；依赖：标题层级
## 4. 目录类

- [ ] `toc.insert_field`：插入 Word 目录域（implemented）
  - 说明：Markdown [TOC] 插入 Word TOC 域。
  - 默认启用：否；可覆盖：是；适用：软著说明书、长报告、项目材料；依赖：heading.level_1
- [ ] `toc.level_1_2`：目录到二级标题（implemented）
  - 说明：TOC 使用 \o "1-2"。
  - 默认启用：否；可覆盖：是；适用：当前基础模板、一般长文；依赖：toc.insert_field
- [ ] `toc.level_profile`：文种目录层级（planned） - planned，不可作为 production 能力
  - 说明：由 profile 指定目录到 2、3 或 4 级。
  - 默认启用：否；可覆盖：是；适用：软著说明书、项目申报书；依赖：toc.insert_field
- [ ] `toc.update_with_word`：Word 更新域（partial） - 可启用但需复核
  - 说明：使用 Microsoft Word/pywin32 刷新目录页码。
  - 默认启用：否；可覆盖：是；适用：有目录的文档；依赖：toc.insert_field
- [ ] `toc.page_breaks`：封面目录正文分页（implemented）
  - 说明：TOC 前后插入分页符。
  - 默认启用：有 [TOC] 时启用；可覆盖：是；适用：软著说明书、长报告；依赖：toc.insert_field
## 5. 页眉页脚页码类

- [x] `page_number.odd_even_official`：公文奇偶页页码（implemented）
  - 说明：奇数页右侧、偶数页左侧，格式类似 - 1 -。
  - 默认启用：一般公文默认；可覆盖：是；适用：一般公文、工作汇报、会议纪要；依赖：page.a4_portrait
- [ ] `page_number.centered`：居中页码（planned） - planned，不可作为 production 能力
  - 说明：页脚居中页码。
  - 默认启用：否；可覆盖：是；适用：合同、说明书、普通报告；依赖：无
- [ ] `header.profile_text`：页眉文字（planned） - planned，不可作为 production 能力
  - 说明：按 profile 插入固定页眉或文档名。
  - 默认启用：否；可覆盖：是；适用：合同、软著说明书、报告；依赖：无
- [ ] `footer.profile_text`：页脚文字（planned） - planned，不可作为 production 能力
  - 说明：按 profile 插入公司名、密级、版本等。
  - 默认启用：否；可覆盖：是；适用：合同、项目材料；依赖：无
## 6. 表格类

- [x] `table.markdown_to_word`：Markdown 表格转 Word（implemented）
  - 说明：将 Markdown 表格生成 Word table。
  - 默认启用：是；可覆盖：是；适用：全部文档；依赖：无
- [x] `table.three_line`：三线表（implemented）
  - 说明：顶线、表头底线、底线，不使用完整网格线。
  - 默认启用：是；可覆盖：是；适用：报告、说明书、测算文档；依赖：table.markdown_to_word
- [x] `table.header_repeat`：表头跨页重复（implemented）
  - 说明：首行设置 w:tblHeader。
  - 默认启用：是；可覆盖：是；适用：长表格文档；依赖：table.markdown_to_word
- [x] `table.cell_style`：表格文字样式（implemented）
  - 说明：表头和表体使用独立样式，不继承正文首行缩进。
  - 默认启用：是；可覆盖：是；适用：全部含表文档；依赖：table.markdown_to_word
- [ ] `table.autofit_content`：根据内容自适应列宽（partial） - 可启用但需复核
  - 说明：表格宽度和列宽根据内容/页面自动调整。
  - 默认启用：是；可覆盖：是；适用：全部含表文档；依赖：table.markdown_to_word
- [ ] `table.profile_borders`：文种表格边框规则（planned） - planned，不可作为 production 能力
  - 说明：合同/清单/报价等可使用网格表或自定义边框。
  - 默认启用：否；可覆盖：是；适用：合同、报价、清单；依赖：table.markdown_to_word
## 7. 图片与图题类

- [x] `image.insert_inline`：插入嵌入型图片（implemented）
  - 说明：Markdown 图片插入为 Word 内嵌图片。
  - 默认启用：是；可覆盖：是；适用：说明书、报告；依赖：无
- [x] `image.center`：图片居中（implemented）
  - 说明：图片段落居中、单倍行距。
  - 默认启用：是；可覆盖：是；适用：说明书、报告；依赖：image.insert_inline
- [x] `caption.figure`：图题（implemented）
  - 说明：图片下方生成图题样式。
  - 默认启用：是；可覆盖：是；适用：说明书、报告；依赖：image.insert_inline
- [ ] `image.profile_size`：图片尺寸规则（partial） - 可启用但需复核
  - 说明：profile 定义图片最大宽度、缩放规则。
  - 默认启用：否；可覆盖：是；适用：软著说明书、验收报告；依赖：image.insert_inline
- [ ] `caption.numbering`：图表自动编号（planned） - planned，不可作为 production 能力
  - 说明：自动生成“图 1”“表 1”等编号。
  - 默认启用：否；可覆盖：是；适用：说明书、报告；依赖：caption features
## 8. 文种结构类

- [ ] `structure.cover`：封面（planned） - planned，不可作为 production 能力
  - 说明：文档封面页结构。
  - 默认启用：否；可覆盖：是；适用：软著说明书、项目申报书、报告；依赖：page.section_breaks 可选
- [ ] `structure.toc_page`：目录页（partial） - 可启用但需复核
  - 说明：目录单独成页。
  - 默认启用：否；可覆盖：是；适用：软著说明书、长报告；依赖：toc.insert_field
- [ ] `structure.body_start_page`：正文另起页（implemented）
  - 说明：封面/目录后正文从新页开始。
  - 默认启用：否；可覆盖：是；适用：软著说明书、长报告；依赖：toc.page_breaks 或 page.section_breaks
- [ ] `structure.signature_block`：签署区（planned） - planned，不可作为 production 能力
  - 说明：合同、报告结尾签章/日期区域。
  - 默认启用：否；可覆盖：是；适用：合同、汇报、会议纪要；依赖：无
- [ ] `structure.appendix`：附件（planned） - planned，不可作为 production 能力
  - 说明：支持附件标题、附件正文和分页。
  - 默认启用：否；可覆盖：是；适用：公文、合同、项目材料；依赖：无
- [ ] `structure.red_head`：红头文件（planned） - planned，不可作为 production 能力
  - 说明：红头、发文字号、签发人、版记等。
  - 默认启用：否；可覆盖：是；适用：红头文件；依赖：页眉/版记/分隔线
## 9. 内容整理类

- [x] `content.normalize_text`：文本编码和换行清洗（implemented）
  - 说明：处理 pasted text、txt、md、content docx，输出 UTF-8 Markdown。
  - 默认启用：是；可覆盖：否；适用：全部文档；依赖：无
- [ ] `content.docx_to_markdown`：内容 Word 转 Markdown（partial） - 可启用但需复核
  - 说明：从内容型 docx 抽取标题、段落、表格为 Markdown。
  - 默认启用：否；可覆盖：是；适用：迁移旧文档；依赖：content.normalize_text
- [x] `content.markdown_structure`：Markdown 结构规范（implemented）
  - 说明：要求用 #、##、###、表格、图片表达结构。
  - 默认启用：是；可覆盖：是；适用：全部文档；依赖：无
- [ ] `content.profile_outline`：文种结构大纲（planned） - planned，不可作为 production 能力
  - 说明：根据文种提供默认章节结构。
  - 默认启用：否；可覆盖：是；适用：软著说明书、合同、工作汇报；依赖：profile 配置
- [ ] `content.writing_style`：文风规则（planned） - planned，不可作为 production 能力
  - 说明：按 profile 或用户文风生成/改写内容。
  - 默认启用：否；可覆盖：是；适用：可选写作流程；依赖：profile 文风规则
## 10. 校验类

- [x] `validation.basic_open`：基础打开校验（implemented）
  - 说明：.docx 可被 python-docx 打开。
  - 默认启用：是；可覆盖：否；适用：全部文档；依赖：无
- [x] `validation.required_styles`：必要样式校验（implemented）
  - 说明：检查 profile/template 必需样式是否存在。
  - 默认启用：是；可覆盖：是；适用：全部文档；依赖：profile rules
- [x] `validation.page_setup`：页面设置校验（implemented）
  - 说明：检查纸张、页边距、页眉页脚距离。
  - 默认启用：是；可覆盖：是；适用：一般公文默认；依赖：profile rules
- [x] `validation.indent_chars`：字符缩进校验（implemented）
  - 说明：禁止非 TOC 段落使用厘米/twip 首行缩进。
  - 默认启用：是；可覆盖：是；适用：公文、正式中文文档；依赖：paragraph.first_line_chars
- [x] `validation.footer_page_number`：页码位置校验（implemented）
  - 说明：检查奇偶页页码或 profile 页码规则。
  - 默认启用：是；可覆盖：是；适用：一般公文默认；依赖：page number features
- [x] `validation.table_rules`：表格规则校验（implemented）
  - 说明：检查三线表、跨页表头。
  - 默认启用：是；可覆盖：是；适用：含表文档；依赖：table features
- [ ] `validation.profile_rules`：profile 专属校验（planned） - planned，不可作为 production 能力
  - 说明：根据 profile 的 validation_rules.json 动态校验。
  - 默认启用：否；可覆盖：是；适用：多文种；依赖：profile 配置
- [ ] `validation.visual_review`：人工/截图复核（partial） - 可启用但需复核
  - 说明：对 Word 版式进行肉眼检查。
  - 默认启用：否；可覆盖：否；适用：重要交付文档；依赖：Word/WPS
