# Multi-level Header Edits

`excel-ir header-edit` and `excel_ir_mvp.header_edit()` locate cells through human report headers rather than fixed coordinates.

Supported layouts:

1. **Horizontal headers**: multi-row headers identify a target column; then a data row is selected.
2. **Vertical headers**: multi-column headers identify a target row; then a data column is selected.

Header values are expanded through merged cells and adjacent blank cells, so a merged year/month/day header can be addressed as a path like `2026 / 5 / 8`.

## Horizontal headers

```text
        2026           2027
        5              5
        7   8   9      7   8   9
门店A   10  20  30     40  50  60
门店B   11  21  31     41  51  61
```

Update `2026 / 5 / 8` for row `门店A`:

```bash
excel-ir header-edit workbook.xlsx edited.xlsx \
  --sheet 日期表 \
  --header-rows 1:3 \
  --headers 2026/5/8 \
  --row-match 门店A \
  --value 999 \
  --as-number
```

If `--row` and `--row-match` are omitted, the target row defaults to the first data row after `--header-rows`:

```bash
excel-ir header-edit workbook.xlsx edited.xlsx \
  --headers 2026/5/8 \
  --value 999 \
  --as-number
```

## Regex and wildcard matching

Use `--match-mode regex` for regular expressions:

```bash
excel-ir header-edit workbook.xlsx ignored.xlsx \
  --headers '202[0-9]/5/[78]' \
  --match-mode regex \
  --value 999 \
  --preview
```

Use `--match-mode wildcard` for shell-style `*` / `?` matching:

```bash
excel-ir header-edit workbook.xlsx edited.xlsx \
  --headers '202?/*/8' \
  --match-mode wildcard \
  --value 999 \
  --as-number
```

Python supports per-level match dictionaries:

```python
import excel_ir_mvp as xir

xir.header_edit(
    "dates.xlsx",
    "edited.xlsx",
    headers=[{"regex": "202[0-9]"}, {"value": "5"}, {"wildcard": "*"}],
    value=999,
    options=xir.HeaderEditOptions(preview=True),
)
```

## Vertical headers

```text
指标    科目    Q1   Q2
收入    线上    100  110
        线下    120  130
成本    线上    60   70
        线下    65   75
```

Here `收入 / 线下` identifies a row, and `Q2` identifies the target column:

```bash
excel-ir header-edit vertical.xlsx edited.xlsx \
  --orientation vertical \
  --header-cols A:B \
  --min-row 2 \
  --headers 收入/线下 \
  --col-match Q2 \
  --value 999 \
  --as-number
```

If `--col` and `--col-match` are omitted, the target column defaults to the first data column after `--header-cols`.

## Main options

- `--headers`: JSON array or slash-separated path.
- `--match-mode exact|contains|wildcard|regex`: header and row/column selector matching.
- `--contains`: shorthand for `--match-mode contains`.
- `--case-sensitive`: make text/regex/wildcard matching case-sensitive.
- `--header-match-index N`: choose the Nth matching header path if duplicates exist.
- Horizontal mode:
  - `--orientation horizontal` (default)
  - `--header-rows START:END`, default `1:3`
  - `--row N` or optional `--row-match VALUE`
  - `--row-match-col COL`, default `1`
  - `--data-start-row N`
  - `--min-col` / `--max-col`
- Vertical mode:
  - `--orientation vertical`
  - `--header-cols START:END`, default `1:3` or e.g. `A:B`
  - `--col COL` or optional `--col-match VALUE`
  - `--col-match-row N`, default `1`
  - `--data-start-col COL`
  - `--min-row` / `--max-row`
- `--preview`: return the plan only.
- `--as-number`: coerce `--value` to int/float.

## Public API

```python
import excel_ir_mvp as xir

xir.header_edit(
    "workbook.xlsx",
    "edited.xlsx",
    headers=["2026", "5", "8"],
    value=999,
    options=xir.HeaderEditOptions(sheet="日期表", row_match="门店A"),
)

xir.header_edit(
    "vertical.xlsx",
    "edited.xlsx",
    headers=["收入", "线下"],
    value=999,
    options=xir.HeaderEditOptions(
        orientation="vertical",
        header_start_col="A",
        header_end_col="B",
        min_row=2,
        col_match="Q2",
    ),
)

cols = xir.header_columns("dates.xlsx", header_rows=(1, 3))
rows = xir.header_rows("vertical.xlsx", header_cols=("A", "B"), min_row=2)
```
