# Release v2.0.0a19 - IR-first Multi-header Editing

## Why

Multi-header location and editing should not reopen `.xlsx` files after parsing. The package model is now explicit:

```text
xlsx -> parse to IR/JSON -> locate/update IR -> rebuild xlsx
```

This keeps edits deterministic, auditable and efficient for large workbooks.

## Changes

- Added IR-native helpers in `excel_ir.py`:
  - `multi_header_columns_ir`
  - `multi_header_rows_ir`
  - `locate_cell_by_multi_header_ir`
  - `update_cell_by_multi_header_ir`
- Public facade is now IR-first:
  - `xir.header_locate(ir, headers=..., options=...)`
  - `edited_ir, result = xir.header_edit(ir, headers=..., value=..., options=...)`
  - `xir.header_columns(ir, ...)`
  - `xir.header_rows(ir, ...)`
- CLI is now IR-first:
  - `excel-ir header-locate workbook.ir.json --headers 2026/5/8 ...`
  - `excel-ir header-edit workbook.ir.json edited.ir.json --headers 2026/5/8 --value 999`
  - then `excel-ir rebuild edited.ir.json edited.xlsx`
- The old XLSX direct helpers remain in internal modules for explicit compatibility, but are not the recommended public flow.
- Documentation and examples were updated to show parse -> IR edit -> rebuild.

## Example

```python
import excel_ir_mvp as xir

ir = xir.parse("dates.xlsx", sheets="日期表")
located = xir.header_locate(ir, headers=["2026", "5", "8"], options=xir.HeaderEditOptions(row_match="门店A"))
edited_ir, result = xir.header_edit(ir, headers=["2026", "5", "8"], value=999, options=xir.HeaderEditOptions(row_match="门店A"))
xir.rebuild(edited_ir, "edited.xlsx")
```

## Validation

- Source CI: 39 tests, coverage gate 70%.
- Installed CI: package installed from wheel and CLI smoke tests pass.
- Build: wheel and sdist pass `twine check`.
