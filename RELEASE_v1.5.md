# Excel IR MVP v1.5 Test and Release Pipeline Hardening

## 目标

v1.5 聚焦测试与发布流程硬化：将测试迁移到 `tests/`，让 CI 同时覆盖 source 和 installed CLI 模式，schema 读取使用 `importlib.resources`，补齐 sdist、twine check、MANIFEST 和发布 checklist。

## 主要变更

### tests/ package-aware 回归测试

新增：

```text
tests/__init__.py
tests/test_excel_ir_mvp.py
```

覆盖：

- package import 和版本检查
- parse → rebuild → diff round-trip diff=0
- patch transaction log impact
- validate IR / patch

根目录 `test_excel_ir_mvp.py` 改为兼容转发：

```python
from tests.test_excel_ir_mvp import *
```

### ci_check.py 支持 --installed

`ci_check.py` 新增两种模式：

```bash
python3 ci_check.py
python3 ci_check.py --installed
```

source 模式覆盖：

- py_compile 包内核心模块
- root wrapper validate
- unittest tests/test_excel_ir_mvp.py
- golden_tests
- corpus_runner

installed 模式覆盖：

- `excel-ir doctor`
- `excel-ir validate ir ...`
- `excel-ir validate patch ...`
- `excel-ir field-map-review ...`

输出：

- `ci_results.json`
- `ci_results_installed.json`

### schema 使用 importlib.resources

`src/excel_ir_mvp/validate_ir.py` 的 `load()` 顺序：

1. 显式文件路径
2. 模块目录文件
3. `importlib.resources.files(__package__)`
4. embedded fallback

这样安装态优先读取 wheel 内 package data，而不是只依赖 embedded schema。

### MANIFEST 和 release checklist

新增：

- `MANIFEST.in`
- `RELEASE_CHECKLIST.md`

`MANIFEST.in` 包含 README、CLI_REFERENCE、ARCHITECTURE、package JSON/schema、tests。

### sdist / wheel / twine check

成功构建：

```text
dist/excel_ir_mvp-1.5.0-py3-none-any.whl
dist/excel_ir_mvp-1.5.0.tar.gz
```

`twine check`：

```text
PASSED
PASSED
```

SHA256：

```text
60486e8e1d1f3aef0a6d326a5c24a18f50346a0cd76761f2a4726352475783d7  dist/excel_ir_mvp-1.5.0-py3-none-any.whl
e78cbc69ac0c418543807f83383351fbbb45d2022aa3b6616359e721fbb53a33  dist/excel_ir_mvp-1.5.0.tar.gz
```

### GitHub Actions 更新

`.github/workflows/ci.yml` 现在包含：

- editable install
- source CI
- unittest tests
- build sdist+wheel
- twine check
- force reinstall wheel
- installed CI

## 验证结果

### tests

```bash
python3 -m unittest -v tests/test_excel_ir_mvp.py
```

结果：

```text
Ran 4 tests in ~39s
OK
```

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

### build + twine check

```bash
python3 -m build --sdist --wheel --outdir /tmp/excel_ir_v15_dist
python3 -m twine check /tmp/excel_ir_v15_dist/*
```

结果：通过。

### wheel install verification

```bash
python3 -m pip install --force-reinstall /tmp/excel_ir_v15_dist/excel_ir_mvp-1.5.0-py3-none-any.whl --break-system-packages
python3 ci_check.py --installed
```

结果：通过。

## 当前限制

- workspace 直接 build 在 iSH 共享目录仍可能受符号链接清理问题影响；release 构建继续使用 clean source `/tmp`。
- source CI 仍依赖根目录历史脚本 `golden_tests.py`、`corpus_runner.py`，后续可迁移到包内命令。
- tests 仍是 unittest；后续可引入 pytest/coverage。

## 下一步 v1.6 建议

1. 将 `golden_tests.py`、`corpus_runner.py` 的根目录运行迁移为包内/installed CLI 测试。
2. 增加 coverage。
3. 增加 `python -m excel_ir_mvp.excel_ir_cli` smoke tests。
4. 增加 API-level tests，减少 subprocess 测试时间。
5. 建立真实样例 corpus 目录和 fixture 管理。
