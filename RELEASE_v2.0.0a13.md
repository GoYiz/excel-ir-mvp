# Release v2.0.0a13

Backend engine selection release. PyPI publishing remains intentionally skipped.

## Highlights

- Added a backend registry in `excel_ir_mvp.backends`.
- Added `excel-ir engines` to show available engines and optional backend status.
- Added `--engine openpyxl|wolfxl|auto` to workbook I/O CLI paths including `parse`, `rebuild`, `diff`, `inspect`, `anonymize`, `stream-edit`, and XLSX metadata helpers.
- Added API-level `engine=` parameters to core workbook operations.
- Added optional `wolfxl` detection. If `wolfxl` is not installed/importable, requesting it raises a clear `BackendUnavailableError`; `auto` falls back to `openpyxl`.
- Parse IR now records the selected backend under `workbook.engine`.
- Added [Backend Engines](docs/backends.md) documentation.

## Validation

- Source CI: unittest + CLI/golden/corpus checks + coverage gate 70%.
- Installed CI: wheel install plus installed CLI smoke checks.
- `twine check`: wheel and sdist pass.

## Example

```bash
excel-ir engines
excel-ir parse workbook.xlsx workbook.ir.json --engine auto
excel-ir stream-edit workbook.xlsx edited.xlsx --match 总计 --value 合计 --engine openpyxl
```
