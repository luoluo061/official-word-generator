# Validation Rules

First-stage validation is intentionally basic. It should catch common generation failures, not prove perfect Word typography.

## Required checks

- The generated `.docx` opens with `python-docx`.
- The document has at least one paragraph.
- The document contains expected text from the Markdown title.
- Heading paragraphs use the expected Word styles where possible.
- Tables in Markdown produce Word tables.
- Missing images are reported.
- Required style names are reported as present or missing.
- Page setup is reported for the first section.

## Report format

Write a Markdown report with:

- Input file path.
- Output file path.
- Paragraph count.
- Table count.
- Styles used.
- Required styles present.
- Warnings.

Warnings are acceptable in first-stage output. Errors should be reserved for cases where generation failed or the output cannot be opened.
