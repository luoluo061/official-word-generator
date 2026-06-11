# Word 模板化生成格式标准

> 目标用途：作为 `official-word-generator` skill 的默认格式标准。  
> 使用方式：Codex 读取本文件理解格式要求；`generate_docx.py` 读取文末 JSON 规则块执行样式兜底。  
> 修改原则：如果只改说明文字，脚本不会自动变化；要影响生成结果，必须同步修改“机器可读规则”里的 JSON。

## 一、适用范围

| 项目 | 内容 |
| --- | --- |
| Skill 名称 | official-word-generator |
| 中文名称 | Word 模板化生成 Skill |
| 输入 | `template.docx` + `content.md` |
| 输出 | `output.docx` + `validation_report.md` |
| 适用文档 | 基础公文、正式报告、说明书、会议纪要、合同等模板化 Word 文档 |
| 当前默认风格 | 中文公文/正式文档基础格式 |

## 二、核心原则

1. 模板优先：模板中已有页面设置、样式、页眉页脚、表格样式时，优先继承模板。
2. 规则兜底：模板缺少必要样式时，使用本文件规则补齐。
3. Markdown 只表达结构：字体、字号、缩进、行距由 Word 模板和本文件规则决定。
4. 生成后必须校验：至少检查段落数、表格数、样式使用、必要样式是否存在。

## 三、页面设置标准

| 项目 | 标准 |
| --- | --- |
| 纸张大小 | A4 |
| 页面方向 | 纵向 |
| 上边距 | 3.7 厘米 |
| 下边距 | 3.5 厘米 |
| 左边距 | 2.7 厘米 |
| 右边距 | 2.7 厘米 |
| 页眉距边界 | 1.5 厘米 |
| 页脚距边界 | 2.8 厘米 |
| 是否奇偶页不同 | 公文页码默认支持奇偶页不同 |

## 四、标题格式标准

### 1. 文档标题

- Markdown：`# 文档标题`
- Word 样式：`公文标题`
- 字体：方正小标宋简体
- 字号：2号 / 22 磅
- 加粗：否
- 对齐：居中
- 首行缩进：0 字符
- 行距：最小值 0 磅
- 大纲级别：无

### 2. 一级标题

- Markdown：`## 一、一级标题`
- Word 样式：`一级标题`
- 字体：黑体
- 字号：3号 / 16 磅
- 加粗：否
- 对齐：左对齐
- 首行缩进：2 字符
- 段前段后：0
- 行距：固定值 28 磅
- 大纲级别：1 级

### 3. 二级标题

- Markdown：`### （一）二级标题`
- Word 样式：`二级标题`
- 字体：楷体_GB2312
- 字号：3号 / 16 磅
- 加粗：否
- 对齐：左对齐
- 首行缩进：2 字符
- 段前段后：0
- 行距：固定值 28 磅
- 大纲级别：2 级

### 4. 三级标题

- Markdown：`#### 1. 三级标题`
- Word 样式：`三级条目`
- 字体：楷体_GB2312
- 字号：3号 / 16 磅
- 加粗：否
- 对齐：左对齐
- 首行缩进：2 字符
- 段前段后：0
- 行距：固定值 28 磅
- 大纲级别：3 级

### 5. 四级标题或条目

- Markdown：`##### （1）四级条目`
- Word 样式：`四级条目`
- 字体：仿宋_GB2312
- 字号：3号 / 16 磅
- 加粗：否
- 对齐：左对齐
- 首行缩进：0 字符
- 段前段后：0
- 行距：固定值 28 磅
- 大纲级别：4 级

## 五、正文格式标准

| 项目 | 标准 |
| --- | --- |
| Markdown | 普通段落 |
| Word 样式 | `Normal` |
| 中文字体 | 仿宋_GB2312 |
| 英文和数字 | Times New Roman |
| 字号 | 3号 / 16 磅 |
| 加粗 | 否 |
| 对齐 | 两端对齐 |
| 首行缩进 | 2 字符 |
| 段前段后 | 0 |
| 行距 | 固定值 28 磅 |

## 六、图片格式标准

| 项目 | 标准 |
| --- | --- |
| Markdown | `![图题](images/demo.png)` |
| 插入方式 | 嵌入型图片 |
| 图片段落 | 居中 |
| 图片段落缩进 | 0 字符 |
| 图片段落行距 | 单倍行距 |
| 图题位置 | 图片下方 |
| 图题样式 | `图题` |
| 图题对齐 | 居中 |

## 七、表格格式标准

| 项目 | 标准 |
| --- | --- |
| Markdown | 标准 Markdown 表格 |
| Word 表格样式 | 优先使用模板表格样式，否则使用 `Table Grid` |
| 表格对齐 | 居中 |
| 表格内容字体 | 仿宋_GB2312 |
| 表格表头字体 | 仿宋_GB2312 |
| 表格表头字号 | 3号 / 16 磅 |
| 表格表头 | 加粗 |
| 表格正文字体 | 仿宋_GB2312 |
| 表格正文字号 | 3号 / 16 磅 |
| 表题样式 | `表题` |

## 八、页码格式标准

| 项目 | 标准 |
| --- | --- |
| 页码位置 | 页脚 |
| 公文默认位置 | 奇数页右侧，偶数页左侧 |
| 页码样式 | `—1—` |
| 字体 | Times New Roman |
| 字号 | 4号 / 14 磅 |

## 九、Markdown 映射规则

| Markdown | Word 样式 |
| --- | --- |
| `#` | `公文标题` |
| `##` | `一级标题` |
| `###` | `二级标题` |
| `####` | `三级条目` |
| `#####` | `四级条目` |
| 普通段落 | `Normal` |
| 图片 | 图片居中 + `图题` |
| 表格 | Word 表格 |
| 引用块 | `备注` |

## 十、修改格式时的操作顺序

1. 先修改本文件中对应的中文说明，方便人工审查。
2. 再修改下方 JSON 中对应字段，保证脚本能执行。
3. 重新运行 `generate_docx.py`。
4. 查看 `validation_report.md`。
5. 打开生成的 `.docx` 做最终肉眼检查。

## 十一、机器可读规则

下面这段 JSON 是脚本实际读取的配置。需要影响生成结果时，必须修改这里。

```json
{
  "styles": {
    "title": {
      "style_name": "公文标题",
      "east_asia": "方正小标宋简体",
      "ascii": "Times New Roman",
      "size_pt": 22,
      "bold": false,
      "first_line_chars": 0,
      "align": "center",
      "line_rule": "at_least",
      "line_pt": 0
    },
    "body": {
      "style_name": "Normal",
      "east_asia": "仿宋_GB2312",
      "ascii": "Times New Roman",
      "size_pt": 16,
      "bold": false,
      "first_line_chars": 2,
      "align": "justify",
      "line_rule": "exact",
      "line_pt": 28
    },
    "h1": {
      "style_name": "一级标题",
      "east_asia": "黑体",
      "ascii": "Times New Roman",
      "size_pt": 16,
      "bold": false,
      "first_line_chars": 2,
      "align": "left",
      "line_rule": "exact",
      "line_pt": 28,
      "outline_level": 0
    },
    "h2": {
      "style_name": "二级标题",
      "east_asia": "楷体_GB2312",
      "ascii": "Times New Roman",
      "size_pt": 16,
      "bold": false,
      "first_line_chars": 2,
      "align": "left",
      "line_rule": "exact",
      "line_pt": 28,
      "outline_level": 1
    },
    "figure_caption": {
      "style_name": "图题",
      "east_asia": "仿宋_GB2312",
      "ascii": "Times New Roman",
      "size_pt": 16,
      "bold": false,
      "first_line_chars": 0,
      "align": "center",
      "line_rule": "exact",
      "line_pt": 28
    },
    "table_caption": {
      "style_name": "表题",
      "east_asia": "仿宋_GB2312",
      "ascii": "Times New Roman",
      "size_pt": 16,
      "bold": false,
      "first_line_chars": 0,
      "align": "center",
      "line_rule": "exact",
      "line_pt": 28
    },
    "note": {
      "style_name": "备注",
      "east_asia": "仿宋_GB2312",
      "ascii": "Times New Roman",
      "size_pt": 16,
      "bold": false,
      "first_line_chars": 0,
      "align": "left",
      "line_rule": "exact",
      "line_pt": 28
    },
    "table_body": {
      "style_name": "表格正文",
      "east_asia": "仿宋_GB2312",
      "ascii": "Times New Roman",
      "size_pt": 16,
      "bold": false,
      "first_line_chars": 0,
      "align": "center",
      "line_rule": "single",
      "line_pt": 0
    },
    "table_header": {
      "style_name": "表格表头",
      "east_asia": "仿宋_GB2312",
      "ascii": "Times New Roman",
      "size_pt": 16,
      "bold": true,
      "first_line_chars": 0,
      "align": "center",
      "line_rule": "single",
      "line_pt": 0
    }
  }
}
```
