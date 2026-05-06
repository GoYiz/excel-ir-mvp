# v2.0.0a11 Streaming Edits and Formula Computed Values

Date: 2026-05-06

## Highlights

- Added human-like streaming cell search/edit without full IR parsing:
  ```bash
  excel-ir stream-edit workbook.xlsx edited.xlsx --match 总计 --value 合计
  excel-ir stream-edit workbook.xlsx edited.xlsx --match 备注 --value 说明 --start right
  ```
- Added APIs:
  - `stream_find_cell_xlsx(...)`
  - `stream_update_first_match_xlsx(...)`
- Formula cells in parsed IR now include a `computed_value` key from the workbook's cached `data_only=True` value.
- Added `docs/streaming-edits.md`.

## Notes

`computed_value` depends on cached values saved by Excel or another calculation engine. If no cached value exists, the key is still present for formula cells with value `null`.

## Publishing

PyPI publishing intentionally skipped. GitHub release assets are used for this alpha line.
