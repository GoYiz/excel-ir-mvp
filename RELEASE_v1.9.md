# Excel IR MVP v1.9 Coverage 60 and Performance Budget

## 目标

v1.9 聚焦三件事：

1. 将 coverage 门槛从 50% 提升到 60%。
2. 补更多 `ir_patch` / `excel_ir_plus` 小粒度 API tests。
3. 给 bench 增加可检查的性能预算 smoke test。

## 主要变更

### coverage 门槛提升到 60%

`ci_check.py` 和 GitHub Actions 中：

```bash
python3 -m coverage report --show-missing --fail-under=60
```

source CI 实测：

```text
TOTAL 2510 statements, 898 missing, 64% coverage
```

60% 门槛通过。

### 测试数增加到 16

`tests/test_excel_ir_mvp.py` 从 12 项增至 16 项，新增覆盖：

- `excel_ir_plus.formula_ref_to_range`
- `excel_ir_plus.chart_title_text`
- `excel_ir_plus.table_style_to_dict/from_dict`
- `excel_ir_plus._strip_none`
- `ir_patch.coord_to_rc/rc_to_coord/iter_range`
- `ir_patch.infer_data_type`
- `ir_patch.compare_value`
- `ir_patch.render_value_template`
- `ir_patch.shift_range_ref`
- `ir_patch.ensure_cell`
- `ir_patch.update_rows_where`
- `ir_patch.append_table_row`
- `ir_patch.recompute_totals`

### bench 性能预算

`bench.py` 和 `src/excel_ir_mvp/bench.py` 新增：

```python
run_bench(max_total_seconds=120.0)
```

输出增加：

- `total_seconds`
- `max_total_seconds`
- `ok` 同时检查 returncode 和总耗时

测试中使用较宽松预算：

```python
summary = bench.run_bench(max_total_seconds=180.0)
```

用于防止明显性能回归。

## 验证结果

### unittest

```bash
python3 -m unittest -v tests.test_excel_ir_mvp
```

结果：16 tests 通过。

### source CI

```bash
python3 ci_check.py
```

结果：`ok: true`。

coverage：`64%`，`--fail-under=60` 通过。

### installed CI

安装 v1.9 wheel 后：

```bash
python3 ci_check.py --installed
```

结果：`ok: true`。

### build / twine

成功构建：

```text
dist/excel_ir_mvp-1.9.0-py3-none-any.whl
dist/excel_ir_mvp-1.9.0.tar.gz
```

`twine check`：通过。

SHA256：

```text
efc71af87b74db224f59d8ca844f20ac8050d14643ac46c1ef57447e2a816c03  dist/excel_ir_mvp-1.9.0-py3-none-any.whl
d7edc16f8ce13a9b8543ee12befc28fad8cc4de1ad10cf7704e6d769949c0c21  dist/excel_ir_mvp-1.9.0.tar.gz
```

## 当前限制

- coverage 已过 60%，但仍有大量 `excel_ir_plus.py` 和 `ir_patch.py` 分支未覆盖。
- bench 预算目前较宽松，只用于 smoke，而非严格性能 SLA。
- Excel Table 多级表头 warning 仍存在。

## 下一步 v2.0 建议

1. 专门解决 Excel Table + 多级/合并表头 warning。
2. 将 patch op 测试拆成独立文件/测试类。
3. 增加更多真实 Excel fixture。
4. 建立性能基线和趋势记录。
5. 发布 v2.0 alpha 结构化路线图。
