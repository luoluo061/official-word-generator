# official-word-generator

Codex Skill for generating Chinese official-style Word documents from pasted text, Markdown, or content `.docx` files.

## Features

- Normalizes pasted text, `.txt`, `.md`, and content `.docx` into UTF-8 Markdown.
- Generates formatted `.docx` documents from a bundled official Word template.
- Applies Chinese official-document page setup, title/body styles, odd/even page numbers, and character-based first-line indentation.
- Supports optional `[TOC]` table of contents with Microsoft Word field updates.
- Formats Markdown tables as three-line Word tables with repeating header rows.
- Produces strict validation reports.

## Install

Download or clone this repository, then copy the skill folder to Codex:

```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.codex\skills" | Out-Null
Copy-Item -Recurse -Force ".\official-word-generator" "$env:USERPROFILE\.codex\skills\official-word-generator"
Test-Path "$env:USERPROFILE\.codex\skills\official-word-generator\SKILL.md"
```

The last command should return `True`.

Alternatively, unzip `release\official-word-generator-skill.zip` directly into:

```text
%USERPROFILE%\.codex\skills
```

The final path must be:

```text
%USERPROFILE%\.codex\skills\official-word-generator\SKILL.md
```

## Usage

Normalize content:

```powershell
py "$env:USERPROFILE\.codex\skills\official-word-generator\scripts\prepare_content.py" `
  --input ".\input.txt" `
  --output ".\content.md"
```

Generate Word:

```powershell
py "$env:USERPROFILE\.codex\skills\official-word-generator\scripts\generate_docx.py" `
  --template "$env:USERPROFILE\.codex\skills\official-word-generator\assets\base-official-template.docx" `
  --content ".\content.md" `
  --output ".\output.docx" `
  --report ".\report.md"
```

Generate with TOC field update:

```powershell
py "$env:USERPROFILE\.codex\skills\official-word-generator\scripts\generate_docx.py" `
  --template "$env:USERPROFILE\.codex\skills\official-word-generator\assets\base-official-template.docx" `
  --content ".\content.md" `
  --output ".\output.docx" `
  --report ".\report.md" `
  --update-fields
```

## Markdown Contract

- `#` -> `公文标题`
- `##` -> `一级标题`
- `###` -> `二级标题`
- `####` and deeper -> body text in the base official template
- `[TOC]` / `[[TOC]]` -> Word table of contents, levels 1-2 only
- Markdown tables -> three-line Word tables with `表格表头` and `表格正文`

## Runtime

- Python 3.10+.
- `python-docx`.
- `lxml`.
- `pywin32` on Windows for `--update-fields`.
- Microsoft Word is required only when automatically updating fields.

The scripts auto-install missing Python packages when possible. Manual install:

```powershell
py -m pip install -r ".\official-word-generator\requirements.txt"
```

## Repository Layout

```text
official-word-generator/
  SKILL.md
  README.md
  requirements.txt
  assets/
  examples/
  references/
  scripts/
release/
  official-word-generator-skill.zip
```

## Status

This project is built for local Codex usage on Windows with Chinese Word documents. It is intentionally conservative: validation errors should block delivery.
