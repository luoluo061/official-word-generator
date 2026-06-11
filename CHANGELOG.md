# Changelog

## 2026-06-11 Table Font Update

- Updated general official table header/body text to 仿宋_GB2312 3号.
- Set table header bold and table body non-bold.
- Documented that generated outputs should go to user-selected/Desktop folders, not the skill directory.

## 2026-06-11 Profile Validation

- Added `scripts/validate_profile.py` for profile configuration integrity checks.
- Added production/draft/planned delivery gates for profile validation.
- Updated profile inspection, development guide, and skill instructions to require profile validation before production promotion.

## 2026-06-11 Profile Scaffolding

- Added `scripts/list_features.py` to display the Word feature catalog by category/status/Markdown.
- Added `scripts/create_profile.py` to scaffold draft profiles from `general_official`.
- Updated the profile development guide and skill instructions for feature-selection-first profile creation.

## 2026-06-11 Reference Profile

- Promoted `profiles/general_official/` as the reference production profile.
- Added `example.md` and `expected_validation_report.md` for profile regression testing.
- Added `references/profile_development_guide.md`.

## 2026-06-11 Profile Execution

- Added `scripts/profile_resolver.py` for profile listing, inspection, and resolution.
- Added `--profile` support to generation and validation commands.
- Added profile status and feature-status reporting to validation reports.
- Preserved legacy `--template` generation flow.

## 2026-06-11

- Updated level-1 heading (`一级标题`) to use 黑体 3号 without bold.
- Regenerated the packaged skill archive.

## 0.1.0

- Initial public repository layout.
- Includes Codex skill files under `official-word-generator/`.
- Includes release zip under `release/`.
- Supports input normalization, Word generation, strict validation, TOC fields, odd/even page numbers, character indentation, and three-line tables.
