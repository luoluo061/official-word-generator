# Markdown Mapping

Default Markdown to Word style mapping.

- `#` -> document title / `公文标题`
- `##` -> level 1 heading / `一级标题`
- `###` -> level 2 heading / `二级标题`
- `####` and deeper -> body / `Normal` in the base official template. Use a derived skill/rule set if a document type needs third/fourth-level heading styles.
- Plain paragraph -> body / `Normal`
- `![caption](path)` -> centered image paragraph plus `图题` caption
- Markdown table -> Word table; add a `表题` paragraph before the table only when the Markdown provides one explicitly in nearby text
- Blockquote -> note / `备注`
- `[TOC]` or `[[TOC]]` -> Word table-of-contents field at that position. The field uses heading outline levels 1-2 only. Third/fourth-level items are kept in the body but excluded from the TOC. The generator inserts a page break before the TOC and another page break after it, so the cover, TOC, and body each start on separate pages. The field must be updated in Word/WPS, or generated with `--update-fields` on a machine with Microsoft Word available.

For Chinese official-style documents, Markdown expresses structure only. Final font, spacing, indentation, and page setup come from the Word template or `references/format_rules.md` fallback rules.

Recommended input structure:

```md
# 文档标题

## 一、一级标题

正文内容。

### （一）二级标题

正文内容。

普通正文内容。

![图1：首页看板页面](images/home.png)

| 项目 | 内容 |
| --- | --- |
| 建设目标 | 提升管理效率 |
```
