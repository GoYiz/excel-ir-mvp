# Release v2.0.0a14

Multi-level header targeting release. PyPI publishing remains intentionally skipped.

## Highlights

- Added `excel-ir header-edit` for locating columns by a multi-row header path.
- Header expansion understands merged cells and horizontal blank fill, so structures such as year/month/day headers work naturally.
- Supports `--headers 2026/5/8` or JSON form `--headers '["2026","5","8"]'`.
- Supports row targeting by explicit `--row` or by scanning a row-label column with `--row-match` and `--row-match-col`.
- Supports `--preview`, `--as-number`, `--contains`, `--case-sensitive`, `--min-col`, `--max-col`, and `--header-match-index`.
- Added API helpers: `multi_header_columns_xlsx`, `locate_cell_by_multi_header_xlsx`, and `update_cell_by_multi_header_xlsx`.
- Added `tests/fixtures/multi_header_dates.xlsx` and `docs/multi-header-edits.md`.

## Example

```bash
excel-ir header-edit workbook.xlsx edited.xlsx \
  --header-rows 1:3 \
  --headers 2026/5/8 \
  --row-match 门店A \
  --value 999 \
  --as-number
```

## Validation

- Source CI: unittest + CLI/golden/corpus checks + coverage gate 70%.
- Installed CI: wheel install plus installed CLI smoke checks.
- `twine check`: wheel and sdist pass.
