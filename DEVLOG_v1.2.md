# Excel IR MVP v1.2 Engineering Hardening

## 本轮目标

v1.2 做打包与工程化硬化：

1. 轻量 dataclass 类型模型。
2. field_map review 页面支持下载 patch JSON。
3. CI 风格检查脚本。
4. 性能基准脚本。
5. 统一 CLI 增加 `doctor` / `bench`。
6. 实际测试 editable install 和 console script。

## 新增/更新文件

- `models.py`：轻量 dataclass 模型和基础类型校验。
- `bench.py`：parse/rebuild/diff/patch 性能基准。
- `ci_check.py`：CI 风格检查脚本。
- `field_map_review_app.py`：新增下载 patch.json 按钮。
- `excel_ir_cli.py`：新增 `doctor`、`bench` 子命令。
- `pyproject.toml`：补齐 `models`、`bench` 模块。
- `bench_results.json`：基准结果。
- `ci_results.json`：CI 检查结果。
- `field_map_review_v12.html`：v1.2 field map review 页面。

## dataclass 模型

新增：

- `CellIR`
- `TableIR`
- `SheetIR`
- `WorkbookIR`

并在 `validate_ir.py` 中加入 `validate_basic_types`。

## field_map review 下载 patch

`field_map_review_app.html` 现在除了生成 patch JSON，还支持点击下载：

```text
下载 patch.json
```

浏览器端会生成 `confirm_field_map.patch.json`。

## CLI doctor

```bash
python3 excel_ir_cli.py doctor
# 或 editable install 后
excel-ir doctor
```

输出：

```json
{
  "ok": true,
  "python": "3.12.13 ...",
  "openpyxl": "3.1.5",
  "project_root": "/var/minis/workspace/excel_ir_mvp"
}
```

## CLI bench

```bash
python3 excel_ir_cli.py bench
```

运行：

- parse
- rebuild
- diff
- patch

并输出 `bench_results.json`。

本次环境样例结果：

```json
{
  "ok": true,
  "results": [
    {"cmd": ["parse"], "seconds": 5.4479},
    {"cmd": ["rebuild"], "seconds": 6.1903},
    {"cmd": ["diff"], "seconds": 6.6418},
    {"cmd": ["patch"], "seconds": 4.9941}
  ]
}
```

## CI check

```bash
python3 ci_check.py
```

覆盖：

1. py_compile 核心脚本。
2. CLI validate IR。
3. CLI validate patch。
4. golden tests。
5. corpus runner。

结果：`ci_results.json`，本次 `ok=true`。

## Editable install 测试

已启用 pip 并执行：

```bash
python3 -m pip install -e . --break-system-packages
excel-ir doctor
```

结果通过。

过程中发现 `pyproject.toml` 漏列 `models` 和 `bench`，已补齐。

## 当前限制

1. dataclass 模型仍是轻量校验，不是完整类型系统。
2. bench 是单样本微基准。
3. CI 脚本是本地脚本，不是真正 GitHub Actions。
4. editable install 已测试，但 wheel 构建尚未测试。

## 下一步 v1.3 建议

1. wheel 构建测试。
2. pytest 化测试。
3. 更严格类型模型或 Pydantic。
4. 大文件性能优化。
5. CLI 文档生成。
