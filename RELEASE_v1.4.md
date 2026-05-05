# Excel IR MVP v1.4 Package Layout Refactor

## 目标

v1.4 将项目从 flat `py-modules` 迁移到标准 `src/excel_ir_mvp/` 包结构，解决后续维护、打包、package data 和 CLI 入口问题，同时保留根目录 `python3 excel_ir_cli.py ...` 的开发兼容入口。

## 结构变化

新增标准包：

```text
src/excel_ir_mvp/
  __init__.py
  excel_ir.py
  excel_ir_plus.py
  ir_patch.py
  excel_ir_cli.py
  validate_ir.py
  models.py
  formula_utils.py
  diff_report.py
  audit_report.py
  corpus_runner.py
  bench.py
  field_map_review.py
  field_map_review_app.py
  ir.schema.json
  patch.schema.json
  corpus_config.json
```

根目录 `excel_ir_cli.py` 变成 wrapper：

```python
from pathlib import Path
import sys
_SRC = Path(__file__).resolve().parent / "src"
if _SRC.exists():
    sys.path.insert(0, str(_SRC))
from excel_ir_mvp.excel_ir_cli import main
```

因此本地仍可运行：

```bash
python3 excel_ir_cli.py doctor
```

安装后运行：

```bash
excel-ir doctor
```

## pyproject.toml

改为标准 package-dir：

```toml
[project.scripts]
excel-ir = "excel_ir_mvp.excel_ir_cli:main"

[tool.setuptools]
package-dir = {"" = "src"}

[tool.setuptools.packages.find]
where = ["src"]
include = ["excel_ir_mvp*"]

[tool.setuptools.package-data]
excel_ir_mvp = ["*.schema.json", "corpus_config.json"]
```

版本：`1.4.0`。

## Import 修复

包内模块改为相对导入并保留 flat-source fallback，例如：

```python
try:
    from . import excel_ir_plus, ir_patch
except ImportError:
    import excel_ir_plus
    import ir_patch
```

修复模块包括：

- `src/excel_ir_mvp/excel_ir_cli.py`
- `src/excel_ir_mvp/validate_ir.py`
- `src/excel_ir_mvp/ir_patch.py`
- `src/excel_ir_mvp/excel_ir_plus.py`

## field_map review 修复

`field_map_review_app.main()` 改为支持 `argv` 参数，CLI 不再 subprocess 调用根目录脚本，而是包内直接调用：

```python
field_map_review_app.main([args.ir_json, args.html])
```

验证：

```bash
excel-ir field-map-review complex_ir_v07.json field_map_review_v14_wheel.html
```

通过。

## 验证结果

### 本地 wrapper

```bash
python3 excel_ir_cli.py doctor
python3 excel_ir_cli.py validate ir complex_ir_v07.json
```

通过。

### editable install

```bash
python3 -m pip install -e . --break-system-packages
excel-ir doctor
excel-ir validate ir complex_ir_v07.json
excel-ir field-map-review complex_ir_v07.json field_map_review_v14.html
```

通过。

### CI check

```bash
python3 ci_check.py
```

结果：`ok: true`。

### unittest

```bash
python3 -m unittest -v test_excel_ir_mvp.py
```

结果：`OK`。

### wheel build / install

工作区直接构建 wheel 在当前 iSH 共享目录仍受 setuptools 清理 egg-info 的符号链接兼容问题影响：

```text
OSError: Cannot call rmtree on a symbolic link
```

v1.4 使用 clean source 复制到 `/tmp` 构建，成功：

```bash
rm -rf /tmp/excel_ir_v14_src /tmp/excel_ir_v14_dist
mkdir -p /tmp/excel_ir_v14_src
cd /var/minis/workspace/excel_ir_mvp
cp pyproject.toml README.md CLI_REFERENCE.md /tmp/excel_ir_v14_src/
mkdir -p /tmp/excel_ir_v14_src/src
cp -R src/excel_ir_mvp /tmp/excel_ir_v14_src/src/
cd /tmp/excel_ir_v14_src
python3 -m build --wheel --outdir /tmp/excel_ir_v14_dist
```

产物：

```text
dist/excel_ir_mvp-1.4.0-py3-none-any.whl
```

SHA256：

```text
5879258cb33272aa873789a84e2e06f4772441a34cd0ef3a803f61dd67585f6a
```

安装验证：

```bash
python3 -m pip install --force-reinstall /tmp/excel_ir_v14_dist/excel_ir_mvp-1.4.0-py3-none-any.whl --break-system-packages
excel-ir doctor
excel-ir validate ir complex_ir_v07.json
excel-ir field-map-review complex_ir_v07.json field_map_review_v14_wheel.html
```

通过。

## CI 更新

`.github/workflows/ci.yml` 增加：

- `python -m build --wheel`
- `pip install --force-reinstall dist/*.whl`
- `excel-ir doctor`

## 当前限制

- 根目录仍保留旧 flat 模块作为开发/历史兼容；后续可逐步标记 deprecated 或迁移测试到包内 API。
- `ci_check.py` 仍主要运行根目录开发脚本；v1.5 可迁移为调用 installed CLI。
- iSH 共享目录 wheel 构建问题仍存在，但 clean source `/tmp` 构建可稳定通过；真实 Linux/GitHub Actions 预期不受影响。

## 下一步 v1.5 建议

1. 新建 `tests/`，迁移 unittest 到 package-aware tests。
2. `ci_check.py` 改成支持 `--installed` 模式。
3. 根目录旧 flat 模块标记 deprecated 或只保留 wrapper。
4. schema 统一从 `importlib.resources` 读取。
5. 增加 sdist 构建和 twine check。
