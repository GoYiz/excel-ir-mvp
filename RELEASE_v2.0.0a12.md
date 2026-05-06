# Release v2.0.0a12

Streaming edit hardening release. PyPI publishing remains intentionally skipped.

## Highlights

- Extended `excel-ir stream-edit` with `--preview` so callers can audit the exact anchor/target change plan without writing a workbook.
- Added `--all` to continue the streaming scan and update every matching cell in scan order.
- Added anchor-relative edits via `--offset-row` and `--offset-col`, allowing human-like workflows such as “find label X, edit the value one row down and two columns right”.
- `stream_find_cell_xlsx` now also accepts offsets and returns target metadata for anchor-relative discovery.
- Formula IR cells now include `computed_value_source: "xlsx_cached_value"` alongside `computed_value`.

## Validation

- Source CI: unittest + CLI/golden/corpus checks + coverage gate 70%.
- Installed CI: wheel install plus installed CLI smoke checks.
- `twine check`: wheel and sdist pass.

## Example

```bash
excel-ir stream-edit workbook.xlsx edited.xlsx --match 业务线 --value 收入本月 --offset-row 1 --offset-col 2 --preview
excel-ir stream-edit workbook.xlsx edited.xlsx --match 云业务 --value 云事业部 --all
```
