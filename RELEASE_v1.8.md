# Excel IR MVP v1.8 Coverage Threshold and API Test Expansion

## 目标

v1.8 聚焦提高 coverage 门槛和扩大直接导入式 API 测试范围。v1.7 已把 coverage 从 subprocess 盲区提升到约 59%，v1.8 在此基础上把门槛提升到 50%，并补充更多核心模块的函数级测试。

## 主要变更

### coverage 门槛提升

`ci_check.py` 和 GitHub Actions 中的 coverage report 从：

```bash
--fail-under=5
```

提高到：

```bash
--fail-under=50
```

本轮 source CI 结果：

```text
TOTAL 2503 statements, 917 missing, 63% coverage
```

### API 测试扩展

`tests/test_excel_ir_mvp.py` 测试数从 8 增加到 12，新增覆盖：

- `models.CellIR/TableIR/SheetIR/WorkbookIR/validate_basic_types`
- `diff_report.render/render_plan/render_tx_log`
- `audit_report.render`
- `field_map_review_app.main(argv)`
- `corpus_runner.load_config/run_corpus`
- `validate_ir.load/validate_json_schema`
- `formula_utils.workbook_formula_dependencies`
- `ir_patch.dry_run`

### package / version

- `pyproject.toml` 版本升级到 `1.8.0`
- `src/excel_ir_mvp/__init__.py` `__version__` 升级到 `1.8.0`

## 验证结果

### unittest

```bash
python3 -m unittest -v tests.test_excel_ir_mvp
```

结果：12 tests 通过。

### source CI

```bash
python3 ci_check.py
```

结果：`ok: true`。

coverage：`63%`，`--fail-under=50` 通过。

### installed CI

安装 v1.8 wheel 后：

```bash
python3 ci_check.py --installed
```

结果：`ok: true`。

### build / twine

成功构建：

```text
dist/excel_ir_mvp-1.8.0-py3-none-any.whl
dist/excel_ir_mvp-1.8.0.tar.gz
```

`twine check`：通过。

SHA256：

```text
97bc343eede86220698fb7b28a4d96936315535246035a4583461461cf33106c  dist/excel_ir_mvp-1.8.0-py3-none-any.whl
306ae90c2f091388515dce1ae7ab103be1cdaeeb1769cac5b1e0a7f56670ad50  dist/excel_ir_mvp-1.8.0.tar.gz
```

## 当前限制

- 仍有大量 `excel_ir.py` / `excel_ir_plus.py` 细分分支未被小粒度测试覆盖。
- `validate_ir.py` CLI main 分支覆盖仍有限。
- Excel Table + 多级/合并表头 warning 仍存在。

## 下一步 v1.9 建议

1. 把 coverage 门槛提升到 60%。
2. 给 `excel_ir_plus.py` 的图表、图片、数据验证、条件格式路径补 fixture/API tests。
3. 给 `ir_patch.py` 的每类 op 增加小型 unit tests。
4. 专门修复或隔离 Excel Table 多级表头 warning。
5. 增加性能预算门槛，防止回归变慢。
