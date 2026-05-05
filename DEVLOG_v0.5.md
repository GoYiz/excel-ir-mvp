# Excel IR MVP v0.5 迭代记录

## 本轮目标

v0.5 重点补“生产可控性”：

1. patch 参数校验。
2. dry-run 预览。
3. 更强 where 条件语法。
4. table field map，降低多级表头/重名表头风险。
5. HTML diff 报告集成 dry-run plan。

## 新增/更新文件

- `ir_patch.py`：新增 `validate_patch`、`dry_run`、复合条件、field_map。
- `diff_report.py`：报告中可嵌入 dry-run plan。
- `v05_patch.json`：v0.5 patch 示例。
- `v05_dry_run.json`：dry-run 计划。
- `v05_apply_plan.json`：实际 apply 前生成的计划。
- `v05_patched_ir.json`：patch 后 IR。
- `v05_patched_report.xlsx`：patch 后 Excel。
- `v05_expected_diff.json` / `v05_expected_diff.html`：增强 diff 报告。

## 新增能力

### 1. Patch validation

`ir_patch.py` 会在 apply 前进行基础校验：

- `actions` 必须是 list。
- `op` 必须已知。
- sheet 必须存在。
- table/chart 引用必须存在。
- 常见必填字段必须存在。

如果有 error，apply 会中止。

### 2. Dry-run

新增 CLI：

```bash
python3 ir_patch.py input_ir.json patch.json --dry-run --plan plan.json
```

返回：

- validation errors
- 每个 action 的影响预览
- update/delete 条件命中的行号
- append_table_rows 将插入几行
- recompute_totals 的 total_row 和列清单
- chart 当前 title/anchor

示例输出：

```json
{
  "ok": true,
  "validation": [],
  "actions": [
    {"op": "append_table_rows", "insert_at": 12, "rows_added": 2},
    {"op": "update_rows_where", "matched_rows": [8, 10], "count": 2},
    {"op": "delete_rows_where", "matched_rows": [11], "count": 1}
  ]
}
```

### 3. 更强 where 条件

v0.5 支持组合条件：

- `all` / `and`
- `any` / `or`
- `not`

示例：

```json
{
  "all": [
    {"col": "区域", "op": "in", "value": ["华南", "华北"]},
    {"not": {"col": "评级", "op": "eq", "value": "A"}}
  ]
}
```

### 4. Table field_map

对复杂多级表头，直接使用表头文本可能重名。v0.5 支持显式字段映射：

```json
"field_map": {
  "业务线": "A",
  "区域": "B",
  "备注": "L",
  "评级": "K"
}
```

然后 where/updates 中可以用稳定字段名：

```json
"where": {"col": "区域", "op": "in", "value": ["华南", "华北"]},
"updates": {"备注": "v0.5 复合条件更新"}
```

### 5. HTML report 集成 plan

`diff_report.py` 支持第四个参数：

```bash
python3 diff_report.py diff.json report.html "Title" plan.json
```

报告会先展示 dry-run plan，再展示 diff 明细。

## v0.5 patch 示例

`v05_patch.json` 做了：

1. 批量追加两行。
2. 用 `field_map` 和复合 where 条件更新 `华南/华北` 且评级非 A 的行。
3. 删除 `服务` 行。
4. 重算总计。
5. 修改图表标题。
6. 写入 annotation。

Dry-run 预览结果：

```text
append_table_rows: insert_at=12, rows_added=2
update_rows_where: matched_rows=[8, 10], count=2
delete_rows_where: matched_rows=[11], count=1
recompute_totals: total_row=12, columns=C,D,E,F,G,H,I,J
```

重建后抽样验证：

```text
table A5:L13
L8 v0.5 复合条件更新
L10 v0.5 复合条件更新
A12 订阅
chart 收入矩阵（v0.5 dry-run 控制）
```

## 当前限制

1. validation 仍是轻量校验，还不是完整 JSON Schema。
2. dry-run 只预览高层影响，不模拟每个下游对象的完整最终坐标。
3. field_map 目前是 action 局部配置，后续应进入 sheet/table schema。
4. where 条件已支持组合，但还没有类型声明/强制类型转换。

## 下一步 v0.6 建议

1. JSON Schema / Pydantic 风格 patch 校验。
2. dry-run 生成完整 impact graph。
3. table field map 持久化进 IR。
4. patch transaction：apply/rollback。
5. report 中按 action 聚合 diff。
6. 真实报表 corpus + golden tests。
