# official-word-generator

Local Word `.docx` generator for Chinese official-style documents.

## What It Does

- Converts pasted text, Markdown, or content `.docx` into normalized UTF-8 Markdown.
- Generates `.docx` from the bundled official Word template.
- Applies Chinese document styles, page setup, odd/even page numbers, figure/table captions, and strict validation.
- Supports optional `[TOC]` table of contents with Word field updates.
- Formats tables as three-line tables with repeating header rows.

## Daily Commands

Normalize input:

```powershell
py scripts\prepare_content.py --input input.txt --output content.md
```

Generate Word:

```powershell
py scripts\generate_docx.py --template assets\base-official-template.docx --content content.md --output output.docx --report report.md
```

Generate Word and update TOC/page fields:

```powershell
py scripts\generate_docx.py --template assets\base-official-template.docx --content content.md --output output.docx --report report.md --update-fields
```

Validate:

```powershell
py scripts\validate_docx.py --docx output.docx --output report.md
```

## Runtime

- Python 3.10+.
- Microsoft Word is required only for `--update-fields`.
- Python dependencies are declared in `requirements.txt`.
- Entry scripts auto-install missing Python packages into the active Python environment when possible.

## Markdown Contract

- `#` -> `公文标题`
- `##` -> `一级标题`
- `###` -> `二级标题`
- `####` and deeper -> body text in the base official template
- `[TOC]` / `[[TOC]]` -> Word table of contents, levels 1-2 only
- Markdown tables -> three-line Word tables with `表格表头` and `表格正文`

## Validation Gates

The generated document must pass:

- Required style availability.
- Official A4 page setup.
- Character-based first-line indent instead of point/cm indent.
- Odd/even page footer alignment.
- Table header repetition.
- Three-line table border structure.
