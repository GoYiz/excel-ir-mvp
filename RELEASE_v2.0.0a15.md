# Release v2.0.0a15

Selective sheet parsing/rebuild plus compact IR output.

## Highlights

- `parse_workbook_plus(..., sheet_names=[...])` and `excel-ir parse --sheet NAME` parse only selected sheets.
- `rebuild_workbook_plus(..., sheet_names=[...])` and `excel-ir rebuild --sheet NAME` rebuild only selected sheets from a larger IR.
- IR records `workbook.sheet_names`; selective parses also record `workbook.selected_sheets`.
- Compact IR removes redundant formula `computed_value_source` while keeping `computed_value`.
- Compact IR omits common default values such as `hidden: false`, `locked: true`, `outlineLevel: 0`, default border flags, default alignment flags, visible `sheet_state`, and default sheet view/format/print option flags.
- Older IR remains rebuild-compatible because rebuild defaults are unchanged.

## CLI examples

```bash
excel-ir parse workbook.xlsx selected.ir.json --sheet 经营驾驶舱
excel-ir parse workbook.xlsx selected.ir.json --sheet Sheet1 --sheet Sheet2
excel-ir rebuild full.ir.json selected.xlsx --sheet Sheet1
```

## Validation

- Source CI: 37 tests, coverage gate 70%.
- Build: wheel + sdist.
- Twine check: passed.
- Installed CI: passed.

PyPI publishing remains skipped.
