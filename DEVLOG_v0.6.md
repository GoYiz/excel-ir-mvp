# Excel IR MVP v0.6 迭代记录

## 本轮目标

v0.6 继续增强生产可控性和可审计性：

1. table field_map 持久化进 IR。
2. dry-run 从静态预览升级为顺序模拟。
3. apply 生成事务式执行日志。
4. HTML 报告同时嵌入 dry-run plan 和 transaction log。

## 新增/更新文件

- `ir_patch.py`：新增 `persist_table_field_map`、顺序模拟 dry-run、`apply_patch_with_log`。
- `diff_report.py`：支持嵌入 transaction apply log。
- `v06_patch.json`：v0.6 patch 示例。
- `v06_dry_run.json`：顺序模拟 dry-run。
- `v06_tx_log.json`：事务式 apply 日志。
- `v06_patched_ir.json`：patch 后 IR，包含持久 field_map。
- `v06_patched_report.xlsx`：patch 后 Excel。
- `v06_expected_diff.html`：集成 plan + log + diff 的 HTML 报告。

## 新增 patch 动作

### persist_table_field_map

将字段映射写入 table IR：

```json
{
  "op": "persist_table_field_map",
  "sheet": "经营驾驶舱",
  "table": "KPI_Table",
  "header_rows": 2,
  "field_map": {
    "业务线": "A",
    "区域": "B",
    "本月收入": "C",
    "预算收入": "D",
    "收入达成率": "E",
    "备注": "L"
  }
}
```

持久化位置：

```json
{
  "extra": {
    "tables": [
      {
        "name": "KPI_Table",
        "ir": {
          "header_rows": 2,
          "field_map": {
            "业务线": "A",
            "区域": "B",
            "本月收入": "C"
          }
        }
      }
    ]
  }
}
```

后续 action 可以直接使用字段名，不必每个 action 重复传 `field_map`。

## 顺序模拟 dry-run

v0.5 dry-run 是基于原始 IR 的静态预览。v0.6 改为：

```text
copy IR → preview action 1 → simulate action 1
        → preview action 2 → simulate action 2
        → ...
```

因此后续 action 的预览会看到前面 action 造成的行号/table ref/chart anchor 等变化。

示例：

```json
{
  "mode": "sequential-simulated",
  "actions": [
    {"op": "append_table_rows", "insert_at": 12, "rows_added": 2},
    {"op": "update_rows_where", "matched_rows": [8, 10]},
    {"op": "delete_rows_where", "matched_rows": [11]},
    {"op": "recompute_totals", "total_row": 13}
  ]
}
```

## 事务式 apply log

新增 CLI 参数：

```bash
python3 ir_patch.py complex_ir_v03.json v06_patch.json v06_patched_ir.json --plan v06_apply_plan.json --log v06_tx_log.json
```

`v06_tx_log.json` 记录每个 action：

- index
- op
- before preview
- after preview

如果中途失败，apply 会抛错并保持“外部未写出结果”的事务语义。

## v0.6 patch 示例做了什么

`v06_patch.json`：

1. 持久化 `KPI_Table` 的 field_map。
2. 批量追加两行。
3. 使用持久字段名更新 `华南/华北 且 评级非 A` 的行。
4. 使用持久字段名删除 `业务线 == 服务` 的行。
5. 使用字段名重算总计：`本月收入`、`预算收入`、`收入达成率` 等。
6. 修改图表标题。

验证抽样：

```text
table A5:L13
fieldmap 本月收入 -> C  （存在于 v06_patched_ir.json）
L8 v0.6 持久字段映射更新
C13 =SUM(C7:C12)
chart 收入矩阵（v0.6 事务日志）
```

## HTML 报告增强

`diff_report.py` 现在支持：

```bash
python3 diff_report.py diff.json report.html "Title" plan.json tx_log.json
```

`v06_expected_diff.html` 同时包含：

1. dry-run plan
2. transaction apply log
3. diff 明细

## 当前限制

1. field_map 持久化在 table 对象内，但重建后 openpyxl 不会保留自定义 table.ir 元数据；它是 IR 层能力，不是 Excel 原生能力。
2. 事务语义是“内存 apply 成功后才写文件”，不是数据库式多文件事务。
3. dry-run 是顺序模拟，但还不是完整 impact graph。
4. diff 仍不能按 action 自动归因，只是报告中并列展示 log 与 diff。

## 下一步 v0.7 建议

1. 按 action 自动归因 diff：action-level before/after snapshots。
2. IR 内保存 patch history。
3. 完整 impact graph：cells/ranges/tables/charts/images/formulas。
4. JSON Schema/Pydantic 级校验。
5. field_map 从 patch 持久化升级为 parse 阶段自动候选 + 人工确认。
