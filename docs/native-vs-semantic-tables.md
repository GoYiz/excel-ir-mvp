# Native vs Semantic Tables

Excel IR MVP keeps table intent separate from the physical OOXML native table object.

## Native table

A table is `native` when it can safely round-trip as an Excel OOXML `Table`:

- one header row
- every heading is a string
- no merged or multi-row semantic header structure
- openpyxl can rebuild it without the warning `column headings must be strings`

Native tables preserve Excel features like structured references, built-in table style, filters, and totals metadata where supported.

## Semantic table

A table is `semantic` when it represents a human report region that should be treated as a logical table but should not be forced back into an unsafe OOXML native table. Typical examples:

- merged headers
- multi-row headers
- artificial section headers
- KPI grids whose logical fields differ from visible Excel column headings
- legacy workbooks that trigger unsafe native table warnings

Semantic tables keep field-map candidates and confirmed `field_map` metadata in IR and in the hidden `_excel_ir_metadata` sheet. During rebuild, the grid, styles, filters, formulas, and semantic metadata are preserved, but unsafe OOXML native table reconstruction is skipped.

## Why this matters

For complex human-authored reports, forcing every logical table into an Excel native table can create invalid or hard-to-open files. The alpha v2 line chooses correctness and auditability over pretending every region is a native table.

## Commands

```bash
excel-ir inspect workbook.xlsx --out inspect.json
excel-ir metadata extract metadata.json --from-xlsx workbook.xlsx
excel-ir metadata repair repaired.xlsx --from-xlsx workbook.xlsx
excel-ir metadata strip stripped.xlsx --from-xlsx workbook.xlsx
```

Use `inspect` to see `native_table_count` and `semantic_table_count` per sheet.
