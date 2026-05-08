# Multi-level Header Edits

`excel-ir header-edit` locates a column by expanding a multi-row header path, including merged cells, then edits the cell at a selected data row.

This covers reports where row 1 is year, row 2 is month, and row 3 is day:

```text
        2026           2027
        5              5
        7   8   9      7   8   9
门店A   10  20  30     40  50  60
```

Update the value for `2026 / 5 / 8` in the row whose first column is `门店A`:

```bash
excel-ir header-edit workbook.xlsx edited.xlsx \
  --sheet 日期表 \
  --header-rows 1:3 \
  --headers 2026/5/8 \
  --row-match 门店A \
  --value 999 \
  --as-number
```

Preview without writing:

```bash
excel-ir header-edit workbook.xlsx ignored.xlsx \
  --headers '["2026","5","8"]' \
  --row-match 门店A \
  --value 999 \
  --preview
```

Options:

- `--headers`: JSON array or slash-separated path.
- `--header-rows START:END`: header row span, default `1:3`.
- `--row N`: edit an explicit row.
- `--row-match VALUE`: find the target row by scanning `--row-match-col`, default column `1`.
- `--data-start-row N`: row search start; defaults to one row after headers.
- `--min-col` / `--max-col`: restrict header column scan.
- `--contains` / `--case-sensitive`: matching behavior.
- `--header-match-index N`: choose the Nth matching header path if duplicates exist.
- `--preview`: return the plan only.
- `--as-number`: coerce `--value` to int/float.

API:

```python
from excel_ir_mvp import locate_cell_by_multi_header_xlsx, update_cell_by_multi_header_xlsx

locate_cell_by_multi_header_xlsx(
    "workbook.xlsx",
    ["2026", "5", "8"],
    sheet="日期表",
    header_start_row=1,
    header_end_row=3,
    row_match="门店A",
)

update_cell_by_multi_header_xlsx(
    "workbook.xlsx",
    "edited.xlsx",
    ["2026", "5", "8"],
    999,
    row_match="门店A",
)
```
