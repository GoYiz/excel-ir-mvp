# Streaming Edits

`excel-ir stream-edit` simulates a human scanning a workbook until the needed cell is found. It does not build the full Excel IR, which makes it useful for small targeted changes in large human-authored reports.

Basic first-match update:

```bash
excel-ir stream-edit workbook.xlsx edited.xlsx --match 总计 --value 合计
excel-ir stream-edit workbook.xlsx edited.xlsx --match 备注 --value 说明 --start right
```

Preview without writing a workbook:

```bash
excel-ir stream-edit workbook.xlsx ignored.xlsx --match 业务线 --value 收入本月 --offset-row 1 --offset-col 2 --preview
```

Update every match in streaming order:

```bash
excel-ir stream-edit workbook.xlsx edited.xlsx --match 云业务 --value 云事业部 --all
```

Anchor-relative update: find an anchor cell, then edit the target at `anchor + offset`:

```bash
excel-ir stream-edit workbook.xlsx edited.xlsx --match 业务线 --value 收入本月 --offset-row 1 --offset-col 2
```

Options:

- `--start left|right`: scan each row left-to-right or right-to-left
- `--sheet SHEET`: restrict to one sheet
- `--contains`: substring matching instead of exact match
- `--case-sensitive`: preserve case when matching text
- `--max-cells N`: stop early after N visited cells
- `--offset-row N` / `--offset-col N`: edit a cell relative to the matched anchor
- `--preview`: return the change plan without writing `out_xlsx`
- `--all`: continue scanning and update all matches instead of stopping after the first
- `--as-number`: coerce `--value` to int/float

API:

```python
import excel_ir_mvp as xir

xir.stream_edit(path, out, match="总计", value="合计")
xir.stream_edit(path, out, match="业务线", value="收入本月", options=xir.StreamEditOptions(offset_row=1, offset_col=2, preview=True))
xir.stream_edit(path, out, match="云业务", value="云事业部", options=xir.StreamEditOptions(update_all=True))
```

The result includes `visited_cells`, `stopped_reason`, and a `changes` list. For offset edits each change records both `anchor` and `target`, making audit trails explicit.
