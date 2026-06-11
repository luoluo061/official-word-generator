# 一般公文 Profile 测试文档

## 一、总体要求

这是 `general_official` profile 的标准测试正文。该段用于验证正文样式、仿宋_GB2312、三号字、首行缩进 2 字符、两端对齐和固定行距 28 磅。

### （一）格式控制

一级标题应使用黑体三号、不加粗、首行缩进 2 字符，并设置为大纲 1 级。二级标题应使用楷体_GB2312 三号、不加粗、首行缩进 2 字符，并设置为大纲 2 级。

### （二）表格控制

| 项目 | 标准 | 验证点 |
| --- | --- | --- |
| 页面 | A4 纵向 | 页边距、页眉页脚距离 |
| 正文 | 仿宋_GB2312 三号 | 首行缩进、行距 |
| 表格 | 三线表 | 表头跨页重复、表格文字样式 |

## 二、校验要求

生成后必须通过 profile-aware validation。报告中应显示 `Profile ID: general_official`、`Status: production`、`Allow Formal Delivery: True`，且 `Errors` 应为 `none`。

> 备注：该文件是 reference production profile 的标准示例，新增 profile 时应准备同等用途的 example.md。
