# Anonymizing Workbook Fixtures

`excel-ir anonymize` creates a shareable workbook by replacing literal cell values while preserving workbook structure as much as openpyxl allows.

```bash
excel-ir anonymize private.xlsx anonymized.xlsx
```

Default behavior:

- numeric values become `0`
- emails become `user@example.com`
- long numeric identifiers become `000000`
- other text becomes length-based placeholders like `文本4`
- formulas are preserved by default
- `_excel_ir_metadata` is stripped because it can contain source labels

If formula text itself must be scrubbed, use:

```bash
excel-ir anonymize private.xlsx anonymized.xlsx --rewrite-formulas
```

Important: this is a best-effort fixture helper, not a legal privacy guarantee. Review the output before publishing.
