# Streaming Edits

`excel-ir stream-edit` simulates a human scanning a workbook until the needed cell is found, then stops and updates only that first match. It does not build the full Excel IR.

```bash
excel-ir stream-edit workbook.xlsx edited.xlsx --match 总计 --value 合计
excel-ir stream-edit workbook.xlsx edited.xlsx --match 备注 --value 说明 --start right
```

Options:

- `--start left|right`: scan each row left-to-right or right-to-left
- `--sheet SHEET`: restrict to one sheet
- `--contains`: substring matching instead of exact match
- `--case-sensitive`: preserve case when matching text
- `--max-cells N`: stop early after N visited cells
- `--as-number`: coerce `--value` to int/float

API:

```python
stream_find_cell_xlsx(path, match, start="left")
stream_update_first_match_xlsx(path, out, match, new_value, start="right")
```

The result includes `visited_cells` and `stopped_reason` so callers can audit that the scan stopped early.
