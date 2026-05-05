# Excel IR MVP v1.1 Production Hardening

## 本轮目标

v1.1 在 v1.0 原型收敛基础上做 production hardening：

1. 更强公式引用解析与迁移。
2. 独立 validate 工具。
3. 统一 CLI 增加 validate / field-map-review 子命令。
4. 可交互 field_map review HTML，可生成 confirm_field_map patch。
5. pyproject 打包入口。

## 新增/更新文件

- `formula_utils.py`：公式引用提取与迁移工具。
- `validate_ir.py`：IR/Patch 独立校验工具。
- `field_map_review_app.py`：可交互 field_map 确认页面。
- `pyproject.toml`：打包配置和 `excel-ir` console script。
- `excel_ir_cli.py`：新增 `validate`、`field-map-review` 子命令。
- `v11_patched_report.xlsx`、`v11_tx.json`、`v11_audit.html`：v1.1 验证产物。

## 公式工具

`formula_utils.py` 支持：

- 提取引用：`extract_references(formula)`
- 移动引用：`shift_formula_references(formula, row_delta, col_delta, row_at, col_at)`
- 扫描 workbook 公式依赖：`workbook_formula_dependencies(ir)`

它会跳过字符串字面量中的伪引用，例如：

```text
=SUM(C7:C12)+"A1"
```

只提取/迁移 `C7:C12`，不会处理 `"A1"`。

也支持 quoted sheet：

```text
='Sheet 1'!A10+B$2+$C3
```

绝对行/列会按 Excel 语义保留。

## validate 工具

独立运行：

```bash
python3 validate_ir.py ir complex_ir_v07.json
python3 validate_ir.py patch v08_patch.json
```

统一 CLI：

```bash
python3 excel_ir_cli.py validate ir complex_ir_v07.json
python3 excel_ir_cli.py validate patch v08_patch.json
```

当前校验包括：

- JSON Schema fallback/可选 jsonschema 校验
- patch semantic validation
- IR formula dependency count/sample

## 可交互 field_map review

新增：

```bash
python3 field_map_review_app.py complex_ir_v07.json field_map_review_app.html
```

或：

```bash
python3 excel_ir_cli.py field-map-review complex_ir_v07.json field_map_review_app.html
```

页面可以编辑字段名和列号，然后生成 `confirm_field_map` patch JSON。

## pyproject

新增 `pyproject.toml`：

```toml
[project.scripts]
excel-ir = "excel_ir_cli:main"
```

后续可以安装为：

```bash
pip install -e .
excel-ir parse input.xlsx out.ir.json
```

## v1.1 验证

已执行：

```bash
python3 validate_ir.py ir complex_ir_v07.json
python3 validate_ir.py patch v08_patch.json
python3 excel_ir_cli.py field-map-review complex_ir_v07.json field_map_review_app.html
python3 excel_ir_cli.py patch complex_ir_v07.json v08_patch.json v11_patched_ir.json --plan v11_plan.json --log v11_tx.json
python3 excel_ir_cli.py rebuild v11_patched_ir.json v11_patched_report.xlsx
python3 excel_ir_cli.py audit v11_tx.json v11_audit.html --title "v1.1 Audit Report"
```

产物：

- `field_map_review_app.html`
- `v11_patched_report.xlsx`
- `v11_tx.json`
- `v11_audit.html`

## 当前限制

1. formula parser 仍不是完整 Excel AST，但比之前正则更安全，支持 sheet 前缀和字符串遮罩。
2. validate schema 仍偏宽松。
3. field_map review 仍是客户端生成 patch，未自动写回文件。
4. pyproject 还未实际发布/安装测试。

## 下一步 v1.2 建议

1. 引入 Pydantic 或 dataclasses 模型。
2. 真正安装包并测试 console script。
3. field_map review 生成 patch 文件下载。
4. 更完整公式 tokenizer。
5. CI 风格测试脚本。
