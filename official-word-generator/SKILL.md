---
name: official-word-generator
description: Use when generating, formatting, inspecting, or validating formal Word `.docx` documents from templates and Markdown/content sources. Supports profile-based document types such as general official documents, software copyright manuals, contracts, reports, and meeting minutes, with reusable Word templates, format rules, feature flags, and validation rules.
---

# Official Word Generator

This skill standardizes formal Word document production. It is not a single-purpose official-document generator; it is a profile-driven workflow for applying Word templates, formatting rules, feature choices, and validation rules across multiple document types.

Primary responsibility: Word document generation, template application, format control, and validation. Do not turn this skill into a broad writing model. Codex may help create or revise Markdown content when needed, but the core output workflow is still `.md` or normalized content -> `.docx` -> validation report.

## Profile Model

The skill has one workflow and multiple document profiles:

- `profiles/general_official/`: 一般公文; current production-ready base profile.
- `profiles/software_copyright/`: 软著说明书; design profile, requires an approved template before production use.
- `profiles/contract/`: 合同; design profile, requires an approved template before production use.

Each profile may contain:

- `profile.md`: purpose, scenarios, default logic, current status.
- `features.json`: enabled/disabled/overridable `feature_id` values.
- `format_rules.json`: profile-specific format rules or placeholders.
- `validation_rules.json`: profile-specific validation expectations.
- `template.docx` or `template_placeholder.md`: approved template or template status.

Before adding a new profile or changing profile behavior, read `references/feature_catalog.md` and select feature IDs from that catalog. Do not hard-code profile-specific behavior in this file.

When adding a new template/profile, first read `references/profile_development_guide.md` and use `profiles/general_official/` as the reference production profile. The new profile must include the same core file structure, including `example.md` and `expected_validation_report.md`, before it can be treated as production-ready.

When the user asks to add a new template:

1. First show or generate the Word feature capability list from `references/feature_catalog.md`.
2. Do not hard-code template capabilities; select existing `feature_id` values.
3. Help map the user's natural-language requirements to `feature_id` values.
4. Use `scripts/create_profile.py` rather than manually creating a profile directory.
5. If the user has not specified exact features, create a draft profile from `general_official` and generate `feature_selection.md` for confirmation.
6. If the user does not provide a Word template, Codex may use `scripts/create_template_docx.py` to generate a draft `template.docx` candidate from a reference profile.
7. Automatically generated `template.docx` files are draft candidates only.
8. Planned features cannot be treated as production delivery capabilities.
9. A production profile must include an approved template, `features.json`, `format_rules.json`, `validation_rules.json`, `example.md`, and `expected_validation_report.md`.
10. Run `scripts/validate_profile.py` before inspecting, generating, or promoting a profile.

When the user asks to upgrade a profile to production:

- Run `py scripts\validate_profile.py --profile <profile_id>` first.
- If `OK: False`, do not mark the profile as production.
- Planned features cannot be delivered as production capabilities.
- Partial features require explicit review and should not be production-critical unless covered by manual verification.

Profile status rules:

- `production`: Can be used for formal generation and delivery after validation passes.
- `draft`: Can be used for test generation; warn the user before any delivery.
- `planned`: Design placeholder only; do not use as a formal output template.

## Profile Selection

1. If the user explicitly names a document type or template, use the matching profile.
2. If the document type is obvious from the request, select the profile and state the basis briefly.
3. If the document type is unclear, list available profiles and ask the user to choose.
4. If the user uploads a custom template, inspect it first and treat it as a temporary custom profile unless it is promoted into `profiles/`.
5. If the user requests feature overrides, apply those overrides on top of the selected profile only when the feature is marked overridable in `features.json`.
6. If a profile references only `template_placeholder.md`, do not claim it is production-ready. Either ask for an approved template or use `general_official` only with an explicit limitation.

## Core Workflow

1. Determine profile:
   - Read the selected `profiles/<profile_id>/profile.md`.
   - Read `profiles/<profile_id>/features.json`.
   - Read profile `format_rules.json` and `validation_rules.json` when relevant.
   - For new profile development, compare against `profiles/general_official/` and follow `references/profile_development_guide.md`.
2. Confirm inputs:
   - Template `.docx`, Markdown `.md`, pasted text, `.txt`, content `.docx`, or user-provided custom template.
3. Normalize content when needed:
   - Use `scripts/prepare_content.py` for pasted text, `.txt`, `.md`, or content `.docx`.
4. Inspect unfamiliar templates:
   - Use `scripts/inspect_template.py` before using a user-uploaded template or promoting a template into a profile.
5. Generate the Word document:
   - Use `scripts/generate_docx.py`.
6. Validate the output:
   - Use `--report` from `generate_docx.py` or run `scripts/validate_docx.py`.
7. Report:
   - Output `.docx` path.
   - Validation report path.
   - Selected profile and any feature overrides.
   - Any planned/partial feature limitations.

## Existing Command Surface

List profiles:

```powershell
py scripts\profile_resolver.py --list
```

Inspect a profile:

```powershell
py scripts\profile_resolver.py --profile general_official --inspect
```

List Word features:

```powershell
py scripts\list_features.py
py scripts\list_features.py --category table
py scripts\list_features.py --status implemented
py scripts\list_features.py --markdown
```

Create a draft profile from the reference profile:

```powershell
py scripts\create_profile.py --profile project_application --name 项目申报书 --from general_official
```

Create a draft template candidate when no Word template is provided:

```powershell
py scripts\create_template_docx.py --profile project_application --from general_official
```

Validate profile configuration:

```powershell
py scripts\validate_profile.py --profile general_official
py scripts\validate_profile.py --all
```

Install or verify dependencies:

```powershell
py scripts\ensure_dependencies.py
```

Normalize pasted text, Markdown, or a content Word file into UTF-8 Markdown:

```powershell
py scripts\prepare_content.py --input path\to\input.txt --output path\to\content.md
```

Do not trust PowerShell `Get-Content` output for Chinese text display. It can render UTF-8 files as mojibake depending on console encoding. Use `prepare_content.py` or Python UTF-8 reads when checking content.

Recommended profile-based generation:

```powershell
py scripts\generate_docx.py --profile general_official --content path\to\content.md --output path\to\output.docx --report path\to\validation_report.md
```

Compatible template-based generation:

```powershell
py scripts\generate_docx.py --template path\to\template.docx --content path\to\content.md --output path\to\output.docx --report path\to\validation_report.md
```

Use a custom rules Markdown file supported by the current generator:

```powershell
py scripts\generate_docx.py --template path\to\template.docx --content path\to\content.md --output path\to\output.docx --report path\to\validation_report.md --rules path\to\format_rules.md
```

Generate with a Word table of contents:

```powershell
py scripts\generate_docx.py --template path\to\template.docx --content path\to\content.md --output path\to\output.docx --report path\to\validation_report.md --update-fields
```

Put `[TOC]` or `[[TOC]]` in Markdown where the table of contents should appear. The current generator inserts a Word TOC field for heading levels 1-2 and page breaks around the TOC. If Microsoft Word/pywin32 is unavailable, the TOC field is inserted but final entries/page numbers must be updated in Word/WPS.

Inspect a template:

```powershell
py scripts\inspect_template.py --template path\to\template.docx --output path\to\template_inspection.md
```

Validate an output document:

```powershell
py scripts\validate_docx.py --docx path\to\output.docx --output path\to\validation_report.md
```

Profile-aware validation:

```powershell
py scripts\validate_docx.py --profile general_official --docx path\to\output.docx --output path\to\validation_report.md
```

## Required Behavior

- Prefer selected profile rules over global defaults.
- Prefer approved profile templates over arbitrary `.docx` files.
- For general official documents, default to `profiles/general_official/` and `assets/base-official-template.docx`.
- Do not use an arbitrary content document as the template just because it is `.docx`; inspect it first.
- When the user changes formatting for a profile, update the profile config first. If the current generator still needs `references/format_rules.md` compatibility, keep that file in sync.
- Keep all Python, Markdown, and JSON files UTF-8.
- Use `pathlib.Path` for paths, especially Chinese paths and filenames.
- Do not write generated `.docx` or validation reports into the skill directory. Use the user's specified output folder; if none is specified, use a Desktop/workspace output folder such as `F:\USER\Desktop\wordgen`.
- Do not use `.doc` as the working template. Convert `.doc` to `.docx` first.
- Generate with `python-docx`; use Word COM only for conversion, PDF export, field/TOC updates, or Word-engine verification.
- Reopen generated `.docx` and verify styles, page setup, indentation, page footers, tables, and profile-specific requirements before claiming completion.
- Treat validation errors as blocking; do not deliver a final document when `OK` is `False`.

## Current Implementation Notes

The scripts now support profile resolution for template/rule selection and profile-aware validation metadata. The `general_official` profile is the current production path. Draft/planned profiles can be inspected and tested only when a usable template exists.

Implemented or partially implemented today:

- Markdown/content normalization.
- `.docx` template inspection.
- Markdown to Word generation.
- A4 official page setup via template/rules.
- Title, level-1, level-2, body, figure caption, table caption, note, table body, and table header styles.
- Character-based first-line indentation.
- Odd/even official page-number validation for the base profile.
- Word TOC field insertion for levels 1-2.
- Optional Word COM field update.
- Markdown tables to three-line Word tables with repeating header rows.
- Inline image insertion and figure captions.
- Validation reports.

Planned profile-aware work:

- Enforcing feature flags in generation logic, not only reporting them.
- Full dynamic validation for every profile-specific rule.
- Template-specific page number modes.
- Software copyright cover/TOC/body structure.
- Contract signature blocks, appendix rules, and clause numbering.
- Dynamic table border modes per profile.
