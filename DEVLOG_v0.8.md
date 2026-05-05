# Excel IR MVP v0.8 迭代记录

## 本轮目标

v0.8 继续补生产审计能力：

1. field_map_candidates → confirmed field_map 的确认动作。
2. action-level 逐项 cell diff 采样归因。
3. 简单 impact graph。
4. 独立审计报告。
5. golden tests。

## 新增/更新文件

- `ir_patch.py`：新增 `confirm_field_map`，事务日志加入 `cell_diffs_sample` 和 `impact_graph`。
- `audit_report.py`：独立审计报告生成器。
- `v08_patch.json`：v0.8 patch 示例。
- `v08_tx_log.json`：包含 action diff sample 和 impact graph 的事务日志。
- `v08_audit_report.html`：独立审计报告。
- `v08_expected_diff.html`：综合 diff 报告。
- `golden_tests.py`：基础 golden tests。

## confirm_field_map

v0.7 parse 阶段会生成候选：

```json
"field_map_candidates": {
  "收入/本月": "C",
  "利润/达成率": "H",
  "状态/评级": "K"
}
```

v0.8 新增动作：

```json
{
  "op": "confirm_field_map",
  "sheet": "经营驾驶舱",
  "table": "KPI_Table",
  "header_rows": 2,
  "field_map": {
    "本月收入": "C",
    "预算收入": "D",
    "评级": "K",
    "备注": "L"
  }
}
```

它会把 parser 候选与人工 override 合并，并持久化为 table 的 confirmed `ir.field_map`。

## Action-level cell diff attribution

`v08_tx_log.json` 的每个 action 现在包含：

- `impact`
- `cell_diffs_sample`

示例：

```json
{
  "op": "append_table_rows",
  "impact": {
    "cells_changed": 28,
    "cells_added": 34,
    "tables_changed": [...]
  },
  "cell_diffs_sample": [
    {"sheet": "经营驾驶舱", "coord": "A11", "before": {...}, "after": {...}}
  ]
}
```

这不是完整逐 cell 归因上限版，目前采样上限为 80 条，足以做 MVP 审计。

## Impact graph

事务日志新增：

```json
"impact_graph": {
  "nodes": {
    "cells": 64,
    "tables": 1,
    "charts": 1
  },
  "edges": ["cells->tables", "cells->charts", "tables->autofilter"]
}
```

这是简化图，表达本次 patch 涉及的对象类型和依赖方向。

## 独立审计报告

新增：

```bash
python3 audit_report.py v08_tx_log.json v08_audit_report.html "v0.8 Standalone Audit Report"
```

报告包含：

- overall impact
- impact graph
- 每个 action 的 impact
- 每个 action 的 cell diff sample

## Golden tests

新增：

```bash
python3 golden_tests.py
```

当前覆盖：

1. `complex_report.xlsx → IR → rebuild → diff_count == 0`
2. `v08_patch.json` 可成功 apply。
3. tx log ok。
4. 有 table impact。
5. 有 action cell diff sample。

运行结果：

```text
golden tests passed
```

## v0.8 产物

- `v08_patched_report.xlsx`
- `v08_expected_diff.html`
- `v08_audit_report.html`
- `v08_tx_log.json`

## 当前限制

1. cell diff attribution 是采样，不是无限完整输出。
2. impact graph 是静态摘要，不是完整依赖图。
3. confirmed field_map 仍保存在 IR，不写回 Excel 原生结构。
4. golden tests 仍是单样本，需要扩展为 corpus。

## 下一步 v0.9 建议

1. 多样本 corpus runner。
2. 完整 action diff 文件分片输出。
3. formula dependency graph。
4. chart/table/image 依赖追踪。
5. 可交互 field_map 确认 UI/HTML。
