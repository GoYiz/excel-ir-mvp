# Excel IR MVP v1.6 Pipeline and Coverage Hardening

## 目标

v1.6 进一步把项目从“可发布包”推进到“可持续 CI 维护”的状态：将测试和 fixture 规范化，补 coverage，增加 `python -m` 启动烟雾测试，并把 source/installed CI 的输入统一到 `tests/fixtures/`。

## 主要变更

### tests/fixtures 规范化

新增/整理：

```text
tests/fixtures/complex_report.xlsx
tests/fixtures/complex_ir_v07.json
tests/fixtures/v08_patch.json
tests/fixtures/v08_patched_report.xlsx
tests/fixtures/corpus_config.json
```

原因：

- 避免 CI 和 unittest 直接依赖工作区根目录散文件。
- 让测试 fixture 更集中、可复制、可移植。

### package-aware unittest

`tests/test_excel_ir_mvp.py` 增加：

- `test_package_import`
- `test_module_cli_smoke`

并继续覆盖：

- round-trip diff=0
- patch transaction log impact
- validate IR / patch

根目录 `test_excel_ir_mvp.py` 保持兼容转发。

### ci_check.py 增加 coverage

source CI 现在包含：

```bash
python3 -m coverage run --source=src/excel_ir_mvp -m unittest tests.test_excel_ir_mvp
python3 -m coverage report --show-missing
```

并把测试输入改为：

```text
tests/fixtures/...
```

同时保留：

- `python3 ci_check.py`
- `python3 ci_check.py --installed`

### corpus_runner / golden_tests fixture 化

`golden_tests.py` 和 `corpus_runner.py` 已切换为 `tests/fixtures/` 输入路径，提升可移植性。

### 版本与许可证

- 版本升级为 `1.6.0`
- 新增 `LICENSE`（MIT）

## 验证结果

### unittest

```bash
python3 -m unittest -v tests.test_excel_ir_mvp
```

结果：5 tests 通过。

### source CI

```bash
python3 ci_check.py
```

结果：`ok: true`。

### installed CI

```bash
python3 ci_check.py --installed
```

结果：`ok: true`。

### coverage

已安装 `py3-coverage`，source CI 中完成 coverage run / report。

### build / twine

成功构建：

```text
excel_ir_mvp-1.6.0-py3-none-any.whl
excel_ir_mvp-1.6.0.tar.gz
```

`twine check` 通过。

SHA256：

```text
2af910d68eaa728512b0242964d328106f5cac7bfde11327a695766acbc8cad8  dist/excel_ir_mvp-1.6.0-py3-none-any.whl
9722421912baa828dc9187f93df1e777088c8e75f5f4cedfa89a584ed369ddb2  dist/excel_ir_mvp-1.6.0.tar.gz
```

## 当前限制

- coverage 因 subprocess 驱动测试，报告偏低，后续要改成更多直接导入式测试。
- corpus / golden 仍是 subprocess 型回归，性能还可以继续优化。
- 现在是测试/发布硬化阶段，还没引入 pytest/coverage 门禁阈值。

## 下一步 v1.7 建议

1. 让 source/installed CI 都调用同一套包内 fixture loader。
2. 增加 coverage 阈值。
3. 把 `golden_tests.py`、`corpus_runner.py` 逐步拆成可重用测试函数。
4. 增加 `python -m excel_ir_mvp.excel_ir_cli` 以及 `excel_ir_mvp` API 的 smoke tests。
5. 进一步减少 subprocess，提升回归速度。
