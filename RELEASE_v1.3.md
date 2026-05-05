# Excel IR MVP v1.3 Release Hardening

## 目标

v1.3 聚焦“发布与测试硬化”：在 v1.2 已具备 CLI、bench、CI 脚本和 editable install 的基础上，补齐更强类型校验、unittest、wheel 构建验证、CLI 文档、GitHub Actions 草案，以及 field_map review 的交互体验。

## 新增/更新

- `models.py`：增强 dataclass 校验，增加坐标合法性、行列号一致性、table ref、field_map 列号检查。
- `test_excel_ir_mvp.py`：unittest 风格回归测试。
- `CLI_REFERENCE.md`：统一 CLI 命令参考。
- `.github/workflows/ci.yml`：GitHub Actions CI 草案。
- `field_map_review_app.py`：增加新增映射行、复制 patch JSON、下载 patch JSON。
- `pyproject.toml`：版本更新为 `1.3.0`。
- `validate_ir.py`：增加内置 schema fallback，修复 wheel 安装后找不到 `ir.schema.json` / `patch.schema.json` 的问题。

## 测试结果

### unittest

```bash
python3 -m unittest -v test_excel_ir_mvp.py
```

结果：

```text
Ran 3 tests in ~32s
OK
```

覆盖：

- round-trip diff 为 0
- validate IR / patch 通过
- patch transaction log 包含 impact

### CI check

```bash
python3 ci_check.py
```

结果：`ok: true`。

覆盖：

- py_compile
- validate IR
- validate patch
- golden_tests
- corpus_runner

### wheel build

直接在 `/var/minis/workspace` 下构建 wheel 时，iSH / 文件系统在清理 egg-info 临时目录时出现：

```text
OSError: Cannot call rmtree on a symbolic link
```

解决方式：复制源码到 `/tmp/excel_ir_build_src` 后构建：

```bash
rm -rf /tmp/excel_ir_build_src /tmp/excel_ir_dist
mkdir -p /tmp/excel_ir_build_src /tmp/excel_ir_dist
cd /var/minis/workspace/excel_ir_mvp
cp *.py pyproject.toml README.md CLI_REFERENCE.md /tmp/excel_ir_build_src/
cd /tmp/excel_ir_build_src
python3 -m build --wheel --outdir /tmp/excel_ir_dist
```

成功产物：

```text
excel_ir_mvp-1.3.0-py3-none-any.whl
```

复制到项目：

- `dist/excel_ir_mvp-1.3.0-py3-none-any.whl`
- `dist/SHA256SUMS`

SHA256：

```text
63370fa6cb7192f1e2e324a86602103b4bd19d268ed3d1b395027db77b8d31db
```

### wheel install verification

```bash
python3 -m pip install --force-reinstall /tmp/excel_ir_dist/excel_ir_mvp-1.3.0-py3-none-any.whl --break-system-packages
cd /var/minis/workspace/excel_ir_mvp
excel-ir doctor
excel-ir validate ir complex_ir_v07.json
excel-ir validate patch v08_patch.json
```

结果均通过。

修复点：wheel 安装后 `validate_ir.py` 的 `ROOT` 指向 site-packages，无法找到项目目录下的 schema 文件。v1.3 已增加 `EMBEDDED_SCHEMAS` fallback，安装态 CLI 也可 validate。

## 当前限制

- 仍采用 flat `py-modules` 打包，后续建议迁移到标准 package layout：`src/excel_ir_mvp/`。
- schema fallback 是内置简化版；后续可以把 schema 作为 package data。
- GitHub Actions 是草案，真实仓库环境还需调整样例文件与依赖。

## 下一步 v1.4 建议

1. 迁移到 `src/excel_ir_mvp/` 标准包结构。
2. 将 schema、示例配置作为 package data。
3. pytest/coverage 集成。
4. 更完整的公式 tokenizer/AST。
5. 更大 corpus 和性能基准矩阵。
