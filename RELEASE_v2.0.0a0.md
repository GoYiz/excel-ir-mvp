# Excel IR MVP v2.0.0a0 Semantic Tables Alpha

## 目标

v2.0 alpha 专门处理长期遗留的 Excel native Table + 多级/合并表头 warning：

```text
UserWarning: File may not be readable: column headings must be strings.
```

核心原则：复杂人工报表的多级/合并表头不一定满足 Excel 原生 Table 约束。v2.0 alpha 将这类表降级为“semantic table”：保留 IR 里的 ref / field_map / patch 语义，但不在重建时强行写回 OOXML native Table，从而避免生成潜在不可读文件和 openpyxl warning。

## 主要变更

### unsafe native table 检测

新增：

```python
excel_table_native_status(ws, ref)
```

判定 Excel native Table 是否可安全重建。规则：

- 表格范围不能与 merged cells 相交。
- 第一行 header 必须全部是非空字符串。
- header 不能重复。

不满足则返回：

```json
{
  "native_table_supported": false,
  "native_table_skip_reason": "merged_cells_intersect_table"
}
```

### parse_tables 增加 native 状态

`parse_tables(ws)` 现在会写入：

- `native_table_supported`
- `native_table_skip_reason`

并继续保留：

- `ref`
- `ir.header_rows`
- `ir.field_map_candidates`
- `ir.confidence`

### apply_tables 跳过 unsafe native table

`apply_tables(ws, items)` 遇到：

```python
item.get("native_table_supported") is False
```

会跳过 native Excel Table 重建，只保留单元格网格/样式/auto_filter/field_map 语义。

### canonical diff 语义化

对于 unsafe native table，`canonical_for_diff` 会把它视为 semantic layer 信息，不参与物理结构 round-trip diff。这样：

```bash
complex_report.xlsx -> IR -> rebuilt.xlsx -> diff
```

仍然 diff=0，同时不再触发 openpyxl warning。

### patch op 测试拆分

新增：

```text
tests/test_patch_ops.py
```

把 patch helper / row ops / warning-focused table tests 从主测试中拆出。

## 验证结果

### unittest

```bash
python3 -m unittest -v tests.test_excel_ir_mvp tests.test_patch_ops
```

结果：17 tests 通过。

### warning-focused test

新增：

```python
test_native_table_warning_is_suppressed_for_multilevel_headers
```

验证：

- `native_table_supported == False`
- `native_table_skip_reason == "merged_cells_intersect_table"`
- rebuild 不再出现 `column headings must be strings` warning
- round-trip diff=0

### source CI

```bash
python3 ci_check.py
```

结果：`ok: true`。

coverage：约 64%，`--fail-under=60` 通过。

### installed CI

安装 v2.0 alpha wheel 后：

```bash
python3 ci_check.py --installed
```

结果：`ok: true`。

### build / twine

成功构建：

```text
dist/excel_ir_mvp-2.0.0a0-py3-none-any.whl
dist/excel_ir_mvp-2.0.0a0.tar.gz
```

`twine check`：通过。

SHA256：

```text
2f868e3081930f8bbcaa3afa66e3650eb8d12e041c2373a89adbd1daac6e98c7  dist/excel_ir_mvp-2.0.0a0-py3-none-any.whl
95f6c94f8d4b51750a784efd4818f8ea8fc76a337af80fcfe6c7ac97f0fafcb5  dist/excel_ir_mvp-2.0.0a0.tar.gz
```

## 设计取舍

v2.0 alpha 明确区分：

1. **native Excel Table**：满足 Excel 原生约束，可安全重建为 OOXML Table。
2. **semantic table**：复杂人工表头，不强行 native rebuild，但保留语义 patch 能力。

这比“为了保留 Table 对象而生成 warning 文件”更适合生产级复杂报表。

## 当前限制

- semantic table 不会在重建文件中以 Excel native Table 对象存在。
- 但 auto_filter、单元格、样式、公式和 field_map 语义仍可用。
- 后续可考虑在 hidden sheet / custom properties 中持久化 semantic metadata。

## 下一步 v2.0 alpha.1 建议

1. 给 semantic table 增加显式 `table_kind: semantic|native` 字段。
2. 将 field_map 持久化到 hidden metadata sheet。
3. 增加 native single-row table fixture，确保 native table 仍会被重建。
4. 增加 semantic table 专用 patch replay tests。
5. 补充文档说明 native vs semantic table 的边界。
