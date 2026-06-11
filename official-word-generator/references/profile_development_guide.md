# Profile Development Guide

Use this guide when adding or promoting a document profile in `official-word-generator`.

The reference production profile is:

`profiles/general_official/`

Do not start a new profile from an empty folder. Copy the reference structure first, then adapt it.

## 1. Choose `profile_id`

Use a stable lowercase snake_case id:

- Good: `software_copyright`, `service_contract`, `meeting_minutes`
- Avoid: Chinese folder names, spaces, temporary names, version numbers

The folder must be:

```text
profiles/<profile_id>/
```

## 2. Copy the reference structure

Start from:

```text
profiles/general_official/
```

Required files for every profile:

```text
profile.md
features.json
format_rules.json
validation_rules.json
template.docx or template_placeholder.md
example.md
expected_validation_report.md
```

If the profile does not yet have an approved Word template, keep it as `draft` and use `template_placeholder.md`.

## 3. Select features from `feature_catalog.md`

Read:

```text
references/feature_catalog.md
```

For every required capability, reference its existing `feature_id`. Do not invent feature ids in `features.json` before adding them to the catalog.

Feature status meanings:

- `implemented`: current scripts/templates support it.
- `partial`: partly supported; requires manual verification or profile-specific work.
- `planned`: design target only.

Do not mark a profile as `production` if it depends on `planned` features for formal delivery.

## 4. Configure `features.json`

`features.json` declares which capabilities the profile enables.

Rules:

- Set `profile_id` and `display_name`.
- Set `template` to `template.docx` or a relative approved template path.
- Keep placeholder profiles as draft by pointing to `template_placeholder.md`.
- Use only feature ids from `feature_catalog.md`.
- For every feature, set:
  - `enabled`
  - `overridable`

Example:

```json
{
  "profile_id": "example_profile",
  "display_name": "示例文档",
  "template": "template.docx",
  "features": {
    "page.a4_portrait": { "enabled": true, "overridable": true },
    "content.normalize_text": { "enabled": true, "overridable": false },
    "validation.basic_open": { "enabled": true, "overridable": false }
  }
}
```

## 5. Configure `format_rules.json`

`format_rules.json` records page and style expectations for the profile.

Include:

- Page size and orientation.
- Margins and header/footer distances.
- Title style.
- Body style.
- Heading styles.
- Table/caption/image style rules when relevant.

If the current generator cannot consume a profile-specific rule yet, still document it here and mark the related feature as `partial` or `planned` in the catalog.

## 6. Configure `validation_rules.json`

`validation_rules.json` declares what must be checked before delivery.

Include:

- `required_styles`
- `page_setup`
- `required_checks`
- `planned_checks` for rules not implemented yet

Use implemented validation features when possible:

- `validation.basic_open`
- `validation.required_styles`
- `validation.page_setup`
- `validation.indent_chars`
- `validation.footer_page_number`
- `validation.table_rules`

If a profile needs rules that are not implemented, keep it as `draft`.

## 7. Prepare `template.docx`

Before using a template in production:

1. Place it at `profiles/<profile_id>/template.docx` or set a stable relative path in `features.json`.
2. Run template inspection:

```powershell
py scripts\inspect_template.py --template profiles\<profile_id>\template.docx --output profiles\<profile_id>\template_inspection.md
```

3. Confirm:
   - Page setup is correct.
   - Required styles exist.
   - Header/footer/page-number behavior matches the profile.
   - Table/image/caption styles match the profile.

Do not use a content document as a template without inspection.

## 8. Prepare `example.md`

Every profile needs a representative Markdown example.

It should cover:

- Document title.
- Main heading levels used by the profile.
- Normal body paragraphs.
- Tables if the profile supports tables.
- Images if the profile requires images.
- Quotes/notes/appendices/signature blocks when relevant.

This example is the regression input for future changes.

## 9. Run inspect / generate / validate

List profiles:

```powershell
py scripts\profile_resolver.py --list
```

Inspect one profile:

```powershell
py scripts\profile_resolver.py --profile <profile_id> --inspect
```

Generate from the profile:

```powershell
py scripts\generate_docx.py --profile <profile_id> --content profiles\<profile_id>\example.md --output tmp\<profile_id>.docx --report tmp\<profile_id>-report.md
```

Validate explicitly:

```powershell
py scripts\validate_docx.py --profile <profile_id> --docx tmp\<profile_id>.docx --content profiles\<profile_id>\example.md --output tmp\<profile_id>-validate.md
```

Save the accepted report as:

```text
profiles/<profile_id>/expected_validation_report.md
```

## 10. Promote draft to production

A profile may be promoted from `draft` to `production` only when all conditions are met:

- It has an approved `template.docx` or approved stable template path.
- `profile_resolver.py --profile <profile_id> --inspect` passes.
- `generate_docx.py --profile <profile_id>` succeeds with `example.md`.
- `validate_docx.py --profile <profile_id>` reports `OK: True`.
- The validation report says `Allow Formal Delivery: True`.
- No formal-delivery-critical feature is still `planned`.
- Any `partial` feature is either non-critical or explicitly covered by manual verification.
- The generated `.docx` has been opened in Word/WPS for visual review when layout matters.

Only after these checks should the profile status be considered production-ready.
