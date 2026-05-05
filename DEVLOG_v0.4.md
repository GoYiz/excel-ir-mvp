# Excel IR MVP v0.4 迭代记录

## 本轮目标

v0.4 重点从“能 patch”推进到“更生产化地 patch”：

1. 多行表格语义追加。
2. 按条件更新表格行。
3. 按条件删除表格行。
4. 自动重算合计行公式。
5. 生成 HTML diff 报告，方便人工检查。

## 新增/更新文件

- `ir_patch.py`：新增 v0.4 表格语义操作。
- `diff_report.py`：新增 HTML diff 报告生成器。
- `v04_patch.json`：v0.4 patch 示例。
- `v04_patched_ir.json`：v0.4 patch 后 IR。
- `v04_patched_report.xlsx`：v0.4 patch 后重建 Excel。
- `v04_expected_diff.json`：原始复杂报表 vs v0.4 patch 后报表的结构化 diff。
- `v04_expected_diff.html`：HTML diff 报告。
- `complex_diff_v03_final.html`：v0.3 round-trip PASS 报告。

## 新增 patch 动作

### append_table_rows

批量追加多行表格数据。内部复用 `append_table_row`，每追加一行会自动：

- 找到 table
- 插入行
- 复制样式
- 写入值/公式
- 扩展 table ref
- 同步 auto_filter

示例：

```json
{
  "op": "append_table_rows",
  "sheet": "经营驾驶舱",
  "table": "KPI_Table",
  "copy_style_from_row": 11,
  "rows": [
    ["硬件", "华中", 760, 700, "=C12/D12", 120, 130, "=F12/G12", 28.5, 45, "B", "批量追加 1"],
    ["订阅", "华东", 1880, 1700, "=C13/D13", 410, 360, "=F13/G13", 61.2, 18, "A", "批量追加 2"]
  ]
}
```

### update_rows_where

按条件更新表格行。支持基础条件：

- `eq` / `==`
- `ne` / `!=`
- `contains`
- `not_contains`
- `gt` / `gte` / `lt` / `lte`
- `regex`
- `in`

示例：

```json
{
  "op": "update_rows_where",
  "sheet": "经营驾驶舱",
  "table": "KPI_Table",
  "header_rows": 2,
  "where": {"col": "B", "op": "eq", "value": "华南"},
  "updates": {"L": "已由 v0.4 条件更新"}
}
```

### delete_rows_where

按条件删除表格行。为了避免行号移动问题，内部会倒序删除。

示例：

```json
{
  "op": "delete_rows_where",
  "sheet": "经营驾驶舱",
  "table": "KPI_Table",
  "header_rows": 2,
  "where": {"col": "A", "op": "eq", "value": "服务"}
}
```

### recompute_totals

自动重算合计行公式。支持：

- `SUM`
- 自定义公式模板
- `{row}`
- `{col}`
- `{data_start}`
- `{data_end}`

示例：

```json
{
  "op": "recompute_totals",
  "sheet": "经营驾驶舱",
  "table": "KPI_Table",
  "header_rows": 2,
  "formulas": {
    "C": "SUM",
    "D": "SUM",
    "E": "=C{row}/D{row}",
    "F": "SUM",
    "G": "SUM",
    "H": "=F{row}/G{row}"
  }
}
```

## v0.4 patch 示例做了什么

`v04_patch.json`：

1. 批量追加两行：`硬件/华中`、`订阅/华东`。
2. 按条件更新 `区域 == 华南` 的备注列。
3. 按条件删除 `业务线 == 服务` 的原始行。
4. 自动重算合计行公式。
5. 扩展数据验证范围。
6. 修改图表标题为 `收入矩阵（v0.4 语义 patch）`。
7. 写入 patch stats 和 annotation。

验证抽样：

```text
table A5:L13
rows 订阅 总计
L8 已由 v0.4 条件更新
total C =SUM(C7:C12)
chart 收入矩阵（v0.4 语义 patch）
```

patch stats：

```json
[
  {"op": "update_rows_where", "count": 1},
  {"op": "delete_rows_where", "count": 1}
]
```

## HTML diff report

新增工具：

```bash
python3 diff_report.py diff.json report.html "Report Title"
```

已生成：

- `v04_expected_diff.html`：原始复杂报表 vs v0.4 patch 后报表。
- `complex_diff_v03_final.html`：v0.3 round-trip PASS 报告。

## 当前限制/观察

1. 语义 patch 可以显著减少动作数量，但复杂表格的“总计行位置/多级表头”仍需用户给 `header_rows` 或明确列号。
2. `delete_rows_where` 会同步移动下方区域，因此后续区域/图表也会跟随移动；这是 Excel 插入/删除行的合理语义，但 diff 会显示较多位置变化。
3. 公式仍是模板/正则级，不是 Excel AST。
4. 条件表达式当前是简单 AND；后续可加 `any/all/not` 组合。

## 下一步 v0.5 建议

1. patch dry-run：执行前输出将被影响的 rows/ranges/objects。
2. patch schema 校验：动作参数错误提前发现。
3. 更强 where 语法：`and/or/not`、列名 disambiguation、类型转换。
4. table field map：显式配置多级表头与字段名。
5. HTML 报告增强：按 sheet/region 聚合，显示 patch stats。
6. 公式 AST/tokenizer。
