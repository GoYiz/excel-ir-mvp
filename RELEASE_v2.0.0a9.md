# v2.0.0a9 IR Compare and Metadata Strip

Date: 2026-05-05

## Highlights

- Added direct IR comparison:
  ```bash
  excel-ir compare-ir a.ir.json b.ir.json ir_diff.json
  ```
- Added metadata strip command:
  ```bash
  excel-ir metadata strip stripped.xlsx --from-xlsx workbook.xlsx
  ```
- Added API helpers:
  - `compare_ir(a, b)`
  - `compare_ir_files(a_path, b_path)`
  - `strip_semantic_metadata_xlsx(src_path, out_path)`
- Added documentation:
  - `docs/native-vs-semantic-tables.md`
  - `tests/fixtures/real_world/README.md`
- CI now covers `compare-ir` and installed `metadata strip`.
- Coverage gate target raised to 70% and passed in source CI.

## Publishing

PyPI publishing intentionally skipped. GitHub release assets are used for this alpha line.
