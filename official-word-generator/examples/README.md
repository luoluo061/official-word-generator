# Example

Generate a Word document from the bundled template and Markdown sample:

```powershell
py ..\scripts\generate_docx.py --template ..\assets\base-official-template.docx --content .\content.md --output .\output.docx --report .\validation_report.md
```

Inspect the template:

```powershell
py ..\scripts\inspect_template.py --template ..\assets\base-official-template.docx --output .\template_inspection.md
```

Validate an existing document:

```powershell
py ..\scripts\validate_docx.py --docx .\output.docx --output .\validation_report.md
```
