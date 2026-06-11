---
name: official-word-generator
description: Use when generating formatted Word `.docx` documents from a user-provided Word template and Markdown content, especially Chinese official-style documents where template styles, fonts, heading levels, page setup, captions, and validation reports must be preserved or checked.
---

# Word Template Generator

Use this skill when the user provides or asks to use:

- a Word template file such as `template.docx`
- Markdown content such as `content.md`
- a request to generate a formatted `.docx` using the template
- a request to inspect or validate Word template formatting

The skill's job is document production, not substantive writing. Codex may help create or revise `content.md`, then this skill turns it into Word using the template.

## Workflow

1. Confirm inputs. If the source content is pasted text, `.txt`, `.md`, or a content `.docx`, normalize it with `scripts/prepare_content.py` first.
2. Read `references/format_rules.md` for default formatting rules when formatting matters.
3. Read `references/markdown_mapping.md` if heading, image, table, or quote mapping is unclear.
4. Run `scripts/inspect_template.py` when the template is unfamiliar or the user asks what format it contains.
5. Run `scripts/generate_docx.py` to create the Word document.
6. Run `scripts/validate_docx.py` or use `--report` from `generate_docx.py` to produce a validation report.
7. Report the output `.docx` path and validation report path.

## Commands

Install or verify dependencies:

```powershell
py scripts\ensure_dependencies.py
```

Normalize pasted text, Markdown, or a content Word file into UTF-8 Markdown:

```powershell
py scripts\prepare_content.py --input path\to\input.txt --output path\to\content.md
```

Do not trust PowerShell `Get-Content` output for Chinese text display. It can render UTF-8 correctly stored files as mojibake depending on console encoding. Use `prepare_content.py` or Python UTF-8 reads when checking content.

Required runtime:

- Python 3.10+ recommended.
- `python-docx` for reading/writing `.docx`.
- `lxml` for strict OOXML validation and normalization.
- `pywin32` on Windows when using `--update-fields` to ask Microsoft Word to refresh TOC/page fields.

The entry scripts automatically install missing Python packages into the active Python environment with `python -m pip install ...`. If the environment cannot access the internet or cannot write to site-packages, install from `requirements.txt` manually before use.

Generate a Word document:

```powershell
py scripts\generate_docx.py --template path\to\template.docx --content path\to\content.md --output path\to\output.docx --report path\to\validation_report.md
```

Use a custom rules Markdown file:

```powershell
py scripts\generate_docx.py --template path\to\template.docx --content path\to\content.md --output path\to\output.docx --report path\to\validation_report.md --rules path\to\format_rules.md
```

Generate with a Word table of contents:

```powershell
py scripts\generate_docx.py --template path\to\template.docx --content path\to\content.md --output path\to\output.docx --report path\to\validation_report.md --update-fields
```

Put `[TOC]` or `[[TOC]]` in the Markdown where the table of contents should appear. The generator inserts page breaks so the cover, TOC, and body start on separate pages. The TOC includes only level-1 and level-2 headings; third/fourth-level items stay in the body but do not appear in the TOC. If Microsoft Word/pywin32 is unavailable, the TOC field is still inserted, but the user must update fields in Word/WPS to render final entries and page numbers.

Inspect a template:

```powershell
py scripts\inspect_template.py --template path\to\template.docx --output path\to\template_inspection.md
```

Validate an output document:

```powershell
py scripts\validate_docx.py --docx path\to\output.docx --output path\to\validation_report.md
```

## Required behavior

- Prefer template styles and page setup over hard-coded formatting.
- For official-document formatting, default to `assets/base-official-template.docx` unless the user explicitly provides a verified official template.
- Do not use an arbitrary content document as the template just because it is `.docx`; inspect it first and reject it if page setup, required styles, odd/even footer settings, or page-number alignment are wrong.
- Use fallback rules only when a required style is missing.
- When the user changes default formatting, edit the machine-readable JSON block in `references/format_rules.md`, then regenerate the document.
- Keep all Python and Markdown files UTF-8.
- Use `pathlib.Path` for paths, especially Chinese paths and filenames.
- Do not use `.doc` as the working template. Convert `.doc` to `.docx` first.
- Generate with `python-docx`; use Word COM only for conversion, PDF export, or Word-engine verification.
- Reopen the generated `.docx` and verify style usage, page setup, `firstLineChars`, and odd/even page footers before claiming completion.
- Treat validation errors as blocking; do not deliver a document when `OK` is `False`.

## Default style names

The scripts look for these preferred styles and fall back when missing:

- `公文标题`
- `一级标题`
- `二级标题`
- `图题`
- `表题`
- `备注`
- `Normal`

If a template uses only Word built-in styles, the scripts use `Heading 1` to `Heading 3` and `Normal` as fallbacks.
