# Excel IR MVP v0.9 迭代记录

## 本轮目标

v0.9 补齐更完整的验证与交互辅助：

1. 多样本 corpus runner。
2. 公式依赖提取并写入 impact graph。
3. field_map 候选确认 HTML。
4. v0.9 审计报告与 diff 报告。

## 新增/更新文件

- `ir_patch.py`：impact_graph 增加 `formulas` 节点和公式依赖采样。
- `corpus_runner.py`：多样本 round-trip/patch runner。
- `field_map_review.py`：生成字段映射候选确认 HTML。
- `field_map_review.html`：字段映射确认页面。
- `v09_tx_log.json`：包含公式依赖 sample 的事务日志。
- `v09_audit_report.html`：v0.9 审计报告。
- `corpus_results/summary.json`：corpus 执行汇总。

## 公式依赖图

`ir_patch.py` 新增 `formula_dependencies`，扫描公式单元格并提取引用：

```json
{
  "sheet": "经营驾驶舱",
  "cell": "E13",
  "formula": "=C13/D13",
  "refs": ["C13", "D13"]
}
```

`impact_graph` 升级为：

```json
{
  "nodes": {
    "cells": 64,
    "tables": 1,
    "charts": 1,
    "formulas": 20
  },
  "edges": [
    "cells->formulas",
    "formulas->charts",
    "cells->tables",
    "tables->autofilter"
  ],
  "formula_dependencies_sample": [...]
}
```

## Field map review HTML

新增：

```bash
python3 field_map_review.py complex_ir_v07.json field_map_review.html
```

输出页面会列出每个 table 的：

- candidate field
- detected column
- confirm/override input

这是 v0.8 `confirm_field_map` 的人工确认辅助页面雏形。

## Corpus runner

新增：

```bash
python3 corpus_runner.py
```

当前样本：

1. `complex_report.xlsx`
2. `v08_patched_report.xlsx`

每个样本执行：

```text
parse -> rebuild -> diff
```

对配置了 patch 的样本额外执行 patch apply。

结果：

```json
{
  "ok": true,
  "results": [
    {"name": "base_complex", "diff_count": 0, "patch_ok": true},
    {"name": "v08_patched_roundtrip", "diff_count": 0}
  ]
}
```

## v0.9 验证结果

执行完成：

```bash
python3 corpus_runner.py
python3 field_map_review.py complex_ir_v07.json field_map_review.html
python3 ir_patch.py complex_ir_v07.json v08_patch.json v09_patched_ir.json --plan v09_apply_plan.json --log v09_tx_log.json
python3 audit_report.py v09_tx_log.json v09_audit_report.html
```

产物：

- `corpus_results/summary.json`
- `field_map_review.html`
- `v09_tx_log.json`
- `v09_audit_report.html`
- `v09_expected_diff.html`

## 当前限制

1. field_map_review.html 还是静态 HTML，未实现提交回写。
2. formula dependency 是正则提取，不是完整 Excel formula parser。
3. corpus runner 当前样本数少，但框架已具备扩展能力。
4. impact graph 仍是摘要图，不是可遍历的完整依赖图。

## 下一步 v1.0 建议

1. 固化项目结构与 CLI。
2. 增加配置文件驱动 corpus。
3. 输出完整 JSON Schema。
4. 生成项目文档和架构图。
5. 封装为可导入 Python package。
