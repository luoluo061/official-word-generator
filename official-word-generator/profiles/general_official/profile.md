# Profile: general_official

## 中文名称

一般公文

## 用途

用于公司日常正式中文文档、普通公文、工作通知、工作汇报、会议纪要等基础场景。

## Profile 定位

- 状态：`production`
- 类型：基础公文模板
- 角色：reference production profile
- 用途：作为后续新增 profile、模板、格式规则和校验规则的开发参照。

新增文种 profile 时，优先复制本目录结构，再根据 `references/feature_catalog.md` 选择功能、修改配置和替换模板。不要直接从空目录开始临时拼装。

## 默认处理逻辑

1. 默认使用 `assets/base-official-template.docx`。
2. 默认使用 A4 纵向页面、公文页边距、奇偶页页码。
3. Markdown `#` 作为文档标题，`##` 作为一级标题，`###` 作为二级标题。
4. `####` 及更深层级暂按正文处理，除非用户明确要求扩展标题层级。
5. 不默认插入封面或目录；只有 Markdown 中出现 `[TOC]` 或用户明确要求时才启用目录。

## 当前实现状态

这是当前已实现能力最完整的 production profile。它复用现有生成脚本、基础模板和 profile-aware 校验器，可作为正式交付路径。

## Reference 文件

- `features.json`：声明 production 可用功能。
- `format_rules.json`：记录基础公文格式规则。
- `validation_rules.json`：记录基础公文校验规则。
- `example.md`：一般公文测试 Markdown。
- `expected_validation_report.md`：基于当前通过结果的标准预期报告。
- `template_placeholder.md`：说明本 profile 通过 `features.json` 指向共享模板 `../../assets/base-official-template.docx`。
