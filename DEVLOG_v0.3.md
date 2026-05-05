# Excel IR MVP v0.3 迭代记录

## 本轮目标

在 v0.2 的复杂对象支持基础上继续推进：

1. 将图表从“仅记录元数据”推进到“基础结构化 IR + 基础重建”。
2. 增加列级结构变更动作。
3. 增加表格级语义动作，减少用户手动维护 table/filter/range。
4. 增加图表 patch 动作。
5. 继续用复杂报表进行 round-trip 与 patch 验证。

## 新增/更新文件

- `excel_ir_plus.py`：升级为 schema `0.3`，新增基础图表 IR 与重建。
- `ir_patch.py`：新增列操作、表格语义动作、图表 patch 动作。
- `v03_patch.json`：v0.3 patch 示例。
- `complex_ir_v03.json`：v0.3 解析后的复杂报表 IR。
- `complex_rebuilt_v03.xlsx`：v0.3 从 IR 重建的复杂报表。
- `complex_diff_v03_final.json`：v0.3 round-trip diff。
- `v03_patched_ir.json`：v0.3 patch 后 IR。
- `v03_patched_report.xlsx`：v0.3 patch 后重建 Excel。

## v0.3 新增能力

### 1. 基础图表 IR

当前可解析/记录：

- chart type：`BarChart` / `LineChart` / `PieChart` 基础类型
- anchor cell
- title
- x_axis_title
- y_axis_title
- style
- height / width 意图记录
- series
  - title_ref
  - title range
  - values range
  - categories range

示例：

```json
{
  "type": "BarChart",
  "anchor": "G16",
  "title": "收入矩阵",
  "x_axis_title": "区域",
  "y_axis_title": "万元",
  "series": [
    {
      "title_ref": "'经营驾驶舱'!B16",
      "values": {"ref": "B17:B20"},
      "categories": {"ref": "A17:A20"}
    }
  ]
}
```

### 2. 基础图表重建

`excel_ir_plus.py rebuild` 现在会根据图表 IR 重建基础图表：

- BarChart
- LineChart
- PieChart
- 标题
- 坐标轴标题
- series value ranges
- category ranges
- anchor

注意：openpyxl 会在读取/保存后规范化 chart height/width，因此 diff 中对 chart size 做了语义归一化。size patch 会保留在 IR 意图里，但重载后的 openpyxl 可能显示默认尺寸。

### 3. 列级结构动作

新增 patch 动作：

- `insert_cols`
- `delete_cols`

插入/删除列时同步移动：

- cells
- merged_ranges
- row/col dimensions
- auto_filter
- tables
- data_validations
- conditional_formatting
- image anchors
- chart anchors
- chart series ranges
- formulas 中的相对引用

### 4. 表格级语义动作

新增：

- `append_table_row`

这个动作会基于 Excel Table 的 ref 自动决定插入位置，并同步扩展 table ref / auto filter，减少手动维护区域的负担。

示例：

```json
{
  "op": "append_table_row",
  "sheet": "经营驾驶舱",
  "table": "KPI_Table",
  "copy_style_from_row": 11,
  "values": ["硬件", "华中", 760, 700, "=C12/D12", 120, 130, "=F12/G12", 28.5, 45, "B", "v0.3 表格语义追加"]
}
```

### 5. 图表 patch 动作

新增：

- `set_chart_title`
- `set_chart_anchor`
- `set_chart_size`
- `set_chart_series_ranges`
- `add_chart`

示例：

```json
{
  "op": "set_chart_title",
  "sheet": "经营驾驶舱",
  "chart": "chart1",
  "title": "收入矩阵（v0.3 重建）",
  "x_axis_title": "区域",
  "y_axis_title": "万元"
}
```

## 验证结果

### 1. v0.3 round-trip

执行：

```bash
cd /var/minis/workspace/excel_ir_mvp
python3 excel_ir_plus.py parse complex_report.xlsx complex_ir_v03.json
python3 excel_ir_plus.py rebuild complex_ir_v03.json complex_rebuilt_v03.xlsx
python3 excel_ir_plus.py diff complex_report.xlsx complex_rebuilt_v03.xlsx complex_diff_v03_final.json
```

结果：

```json
{
  "diff_count": 0,
  "diffs": [],
  "truncated": false
}
```

说明 v0.3 的复杂报表可逆核心，包括基础图表 IR，已通过结构化 diff。

### 2. v0.3 patch 验证

执行：

```bash
python3 ir_patch.py complex_ir_v03.json v03_patch.json v03_patched_ir.json
python3 excel_ir_plus.py rebuild v03_patched_ir.json v03_patched_report.xlsx
python3 excel_ir_plus.py parse v03_patched_report.xlsx v03_patched_roundtrip_ir.json
```

验证要点：

- `append_table_row` 成功追加业务行。
- `insert_cols` 成功增加 `M` 列负责人字段。
- Table ref 更新为 `A5:M13`。
- AutoFilter 更新为 `A5:M13`。
- 数据验证扩展为 `K7:K12`。
- 图表标题更新为 `收入矩阵（v0.3 重建）`。
- 图表 anchor 更新为 `H16`。
- 图表 series 数量为 3。

抽样输出：

```text
row 硬件 负责人 王六
table A5:M13
chart 收入矩阵（v0.3 重建） H16
series 3
```

## 关键修正

### append_table_row 双重扩展 bug

最初 `append_table_row` 在调用 `insert_rows` 后又手动 shift table ref，导致 table ref 被双重扩展。已修正：

- `insert_rows` 内部已经通过 `shift_refs` 移动/扩展 table ref。
- `append_table_row` 不再二次 shift，只同步 auto_filter。

### 图表 title range 重建偏移 bug

openpyxl 的 `add_data(..., titles_from_data=True)` 要求 data reference 包含标题行。解析出的 values range 不含标题行，因此重建时一度发生 series title/values 下移。已修正：

- 重建 Reference 时，如果 title range 正好在 values 上一行，则自动把 min_row 上移到 title row。

### chart size 归一化

openpyxl 对 chart height/width 的 load/save 处理会回到默认值。v0.3 暂时将 chart size 作为“意图字段”保留，但 diff 不以重载后的尺寸为失败条件。

## 当前限制

1. 图表重建仍是基础版：复杂图表格式、数据标签、图例布局、颜色等尚未完整 IR 化。
2. 公式更新仍是正则级，不是完整公式 AST。
3. `append_table_row` 对多级表头/复杂表格映射仍需增强。
4. 列插入删除对部分高级对象只做基础范围移动。
5. openpyxl 的 Excel Table warning 仍存在，需要后续处理合并表头与 Table header 的冲突。

## 下一步建议 v0.4

1. 实现公式 tokenizer/AST 级引用更新。
2. 增强 table semantic actions：`append_rows`、`update_rows_where`、`recompute_totals`。
3. 完善 chart style/color/legend/data label IR。
4. 增加 image two-cell anchor 精确重建。
5. 加入 patch schema 校验与 dry-run 预览。
6. 加入真实报表 corpus 和 HTML diff report。
