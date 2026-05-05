# Excel IR MVP v1.0 原型收敛版

## 本轮目标

v1.0 不再继续扩散功能，而是把 v0.1-v0.9 的能力收敛成一个可使用的原型工程：

1. 统一 CLI。
2. 配置驱动 corpus runner。
3. JSON Schema。
4. Python package 入口。
5. 架构文档。
6. 端到端验证报告。

## 新增/更新文件

- `excel_ir_cli.py`：统一 CLI。
- `__init__.py`：可导入 package 入口。
- `corpus_config.json`：配置驱动 corpus。
- `corpus_runner.py`：改造为配置驱动。
- `ir.schema.json`：IR JSON Schema 雏形。
- `patch.schema.json`：Patch JSON Schema 雏形。
- `ARCHITECTURE.md`：架构文档。
- `v10_ir.json` / `v10_rebuilt.xlsx` / `v10_diff.json`：v1.0 round-trip 验证。
- `v10_patched_ir.json` / `v10_tx.json` / `v10_audit.html`：v1.0 patch/audit 验证。

## 统一 CLI

```bash
python3 excel_ir_cli.py parse input.xlsx out.ir.json
python3 excel_ir_cli.py rebuild out.ir.json rebuilt.xlsx
python3 excel_ir_cli.py diff input.xlsx rebuilt.xlsx diff.json
python3 excel_ir_cli.py patch out.ir.json patch.json patched.ir.json --plan plan.json --log tx.json
python3 excel_ir_cli.py report diff.json report.html --plan plan.json --log tx.json
python3 excel_ir_cli.py audit tx.json audit.html
python3 excel_ir_cli.py corpus --config corpus_config.json
```

## Python package 入口

`__init__.py` 暴露：

```python
from excel_ir_mvp import (
    parse_workbook_plus,
    rebuild_workbook_plus,
    diff_workbooks_plus,
    apply_patch,
    apply_patch_with_log,
    dry_run,
    validate_patch,
)
```

## JSON Schema

新增：

- `ir.schema.json`
- `patch.schema.json`

当前是 MVP 级 schema，主要约束顶层结构和常见 action 字段；后续可升级为严格 schema。

## Corpus 配置

`corpus_config.json`：

```json
{
  "output_dir": "corpus_results",
  "samples": [
    {"name": "base_complex", "xlsx": "complex_report.xlsx", "patch": "v08_patch.json"},
    {"name": "v08_patched_roundtrip", "xlsx": "v08_patched_report.xlsx"}
  ]
}
```

运行：

```bash
python3 excel_ir_cli.py corpus --config corpus_config.json
```

结果：`ok=true`，样本 round-trip diff 均为 0。

## 端到端验证

执行：

```bash
python3 excel_ir_cli.py parse complex_report.xlsx v10_ir.json
python3 excel_ir_cli.py rebuild v10_ir.json v10_rebuilt.xlsx
python3 excel_ir_cli.py diff complex_report.xlsx v10_rebuilt.xlsx v10_diff.json
python3 excel_ir_cli.py patch v10_ir.json v08_patch.json v10_patched_ir.json --plan v10_plan.json --log v10_tx.json
python3 excel_ir_cli.py report v10_diff.json v10_roundtrip_report.html --title "v1.0 Roundtrip Report" --plan v10_plan.json --log v10_tx.json
python3 excel_ir_cli.py audit v10_tx.json v10_audit.html --title "v1.0 Audit Report"
```

结果：

```json
{
  "diff_count": 0,
  "diffs": [],
  "truncated": false
}
```

说明 v1.0 统一 CLI 的 parse/rebuild/diff 闭环通过。

## 当前 v1.0 原型能力

- Excel → IR
- IR → Excel
- structural diff
- semantic patch
- dry-run
- transaction log
- patch history
- action impact
- cell diff sample
- formula dependency sample
- field_map candidates / confirm action
- HTML diff report
- HTML audit report
- corpus runner
- golden tests
- JSON Schema 雏形

## 后续方向

1. 严格 JSON Schema / Pydantic 模型。
2. 完整公式 AST。
3. 更完整图表/图片/条件格式 IR。
4. 真实报表 corpus。
5. field_map review 页面回写 patch。
6. 封装为 pip 包或服务 API。
