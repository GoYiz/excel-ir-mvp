# Metadata Commands

Excel IR semantic metadata captures logical table intent that should survive XLSX-only workflows.

The metadata carrier is a veryHidden sheet:

```text
_excel_ir_metadata!A1
```

The payload includes a SHA-256 checksum. If the checksum does not match, parsing ignores the payload.

## Status

```bash
excel-ir metadata status workbook.xlsx
```

Reports whether the hidden carrier exists, whether it parses, and whether the checksum is valid.

## Extract / verify / repair / strip

```bash
excel-ir metadata extract metadata.json --from-xlsx workbook.xlsx
excel-ir metadata verify metadata.json
excel-ir metadata verify --from-xlsx workbook.xlsx
excel-ir metadata repair repaired.xlsx --from-xlsx workbook.xlsx
excel-ir metadata strip stripped.xlsx --from-xlsx workbook.xlsx
```

Use `repair` to regenerate metadata. Use `strip` before sharing a workbook if labels in semantic metadata may be sensitive.
