# Skill goal

This skill is for Codex, not for a human Word user interface.

The goal is to preserve the stable workflow discovered during the Word template work: given a Word `.docx` template and Markdown content, Codex should generate a new `.docx` that inherits the template's page setup, styles, fonts, outline levels, headers, footers, page numbers, figure captions, table captions, and Chinese official-document formatting rules.

The skill should prevent repeated failures seen during the earlier workflow:

- Garbled Chinese paths or filenames.
- Python source files saved with corrupted Chinese text.
- Word styles appearing correct in body text but not actually written into the style system.
- Built-in `Heading 1` to `Heading 4` lacking outline levels.
- Word and WPS showing different style-gallery behavior.
- Missing fonts causing official-document typography to fall back silently.
- Theme-colored heading styles producing blue text.
- Generated documents that contain content but do not inherit page settings.
- Reliance on unstable `.doc` files instead of `.docx`.
- Treating Word COM as the primary generation path instead of a validation or conversion tool.
- Returning a document without opening it back up and checking styles.

The operating principle is: template first, built-in rules second. If the template contains usable styles, preserve and use them. If a required style is missing, create only the minimum missing style according to the official-document fallback rules.

Minimum accepted workflow:

1. Validate the template is `.docx`, or convert `.doc` to `.docx` before use.
2. Inspect template styles and page setup before writing content.
3. Parse Markdown into structural blocks.
4. Map Markdown headings and blocks to Word styles.
5. Generate the `.docx` using `python-docx`.
6. Reopen the generated `.docx` and verify style usage, outline levels, page setup, and paragraph counts.
7. Use Word COM only when needed for PDF export, `.doc` conversion, or final Word-engine verification.

Default Markdown mapping:

- `#` -> `公文标题`
- `##` -> `一级标题`
- `###` -> `二级标题`
- `####` -> `三级条目`
- `#####` -> `四级条目`
- Plain paragraph -> `Normal` or template body style
- Image -> centered image plus `图题`
- Table -> Word table plus `表题`
- Blockquote or explicit note -> `备注`

Fallback Chinese official-document styles live in `references/format_rules.md`. Load that file whenever the task involves building, updating, or validating the actual Word formatting rules.
