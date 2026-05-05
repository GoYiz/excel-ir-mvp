# Excel IR MVP v1.7 Coverage and API Test Hardening

## 目标

v1.7 聚焦覆盖率与 API 级测试硬化：新增直接导入式 API tests 和 fixture loader，减少只靠 subprocess 的测试盲区；给 coverage 增加最小门槛；补 `python -m excel_ir_mvp` 包入口；并重新构建/安装验证 v1.7 包。

## 主要变更

### 直接导入式 API tests

`tests/test_excel_ir_mvp.py` 新增 API 级测试：

- `test_api_parse_diff_validate`
- `test_api_patch_and_formula_tools`
- `test_package_module_main_smoke`

直接覆盖：

- `excel_ir_mvp.excel_ir_plus.parse_workbook_plus`
- `rebuild_workbook_plus`
- `diff_workbooks_plus`
- `validate_json_schema`
- `validate_basic_types`
- `apply_patch_with_log`
- `validate_patch`
- `formula_utils.extract_references`
- `formula_utils.shift_formula_references`

总测试数从 5 增加到 8。

### fixture loader

新增：

```text
tests/fixtures_loader.py
```

提供：

```python
fixture_path(name)
load_json_fixture(name)
```

用于统一 fixture 路径和 JSON 加载。

### python -m excel_ir_mvp

新增：

```text
src/excel_ir_mvp/__main__.py
```

因此安装后可运行：

```bash
python3 -m excel_ir_mvp doctor
```

### coverage 门槛

`ci_check.py` 中 coverage report 增加：

```bash
python3 -m coverage report --show-missing --fail-under=5
```

v1.7 API tests 让 coverage 从 v1.6 的 subprocess 盲区提升到约 59%。当前门槛先设 5%，避免早期重构阶段过早阻塞；后续可逐步提高。

### installed CI 增强

`ci_check.py --installed` 新增：

```bash
python3 -m excel_ir_mvp doctor
```

验证 wheel 安装态的 `__main__` 入口。

## 验证结果

### unittest

```bash
python3 -m unittest -v tests.test_excel_ir_mvp
```

结果：8 tests 通过。

### source CI

```bash
python3 ci_check.py
```

结果：`ok: true`。

coverage report：约 `59%`，`--fail-under=5` 通过。

### installed CI

安装 v1.7 wheel 后：

```bash
python3 ci_check.py --installed
```

结果：通过。

## 构建产物

- `dist/excel_ir_mvp-1.7.0-py3-none-any.whl`
- `dist/excel_ir_mvp-1.7.0.tar.gz`

## 当前限制

- coverage 门槛仍低，后续应逐步提高。
- 仍有大量测试通过 CLI/subprocess 间接覆盖，后续要拆更多 API 级 tests。
- Excel Table + 多级/合并表头仍会触发 openpyxl warning：`column headings must be strings`，当前作为已知限制保留。

## 下一步 v1.8 建议

1. 将 coverage 门槛提高到 50% 或按核心模块分模块设门槛。
2. 给 `excel_ir_plus.py` 和 `ir_patch.py` 增加更多小粒度 API tests。
3. 将 `golden_tests.py` / `corpus_runner.py` 拆为可导入函数测试。
4. 为已知 openpyxl Table warning 增加专门 issue/文档和修复实验。
5. 增加更多真实复杂报表 fixture。
