# Excel 报表 IR MVP：第一性原理猜想设计

## 0. 问题定义

输入是人工制作、格式复杂、面向人阅读的 Excel 报表。目标不是简单 `read_excel` 成二维数据，而是提取两类信息：

1. **物理/位置层**：工作簿、工作表、单元格坐标、值、公式、合并区域、行高列宽、样式、边框、填充、对齐、数字格式、冻结窗格等。
2. **逻辑/语义层**：报表标题、元信息、表头、数据区、合计行、分节、空白分隔区域等。

然后可以基于该 IR 做局部修改，再重建 `.xlsx`。

## 1. 第一性原理拆解

一个 Excel 报表本质是：

```text
Workbook = Sheets + Shared Resources + Relationships
Sheet = 2D Grid + Layout + Objects + Rules
Cell = Address + Value/Formula + StyleRef + Links/Comments/Validation...
Human Report = Visual Encoding + Spatial Grouping + Text Semantics
```

因此 MVP 应分两层：

- **Fidelity IR**：尽可能可逆，服务于 round-trip 重建。
- **Logical IR**：从 Fidelity IR 推断出人类报表结构，服务于理解与局部修改。

## 2. 初始猜想

### 猜想 A：复杂报表的最小可逆核心

第一版不追求图表、图片、宏、透视表等高级对象，先证明以下对象足以覆盖大量人工报表的主体视觉结构：

- sheet 名称、大小
- cell value/data_type/formula
- style：font/fill/border/alignment/number_format/protection
- merged ranges
- row heights、hidden
- column widths、hidden
- freeze panes

### 猜想 B：逻辑结构可由低阶视觉线索推断

无需 LLM，先用可解释启发式：

- 大号/加粗/居中/横向合并的顶部文本 → title
- `:`、`：`、日期/单位/制表人等关键词 → metadata
- 连续非空区域 + 边框密度高 → table region
- 表格上方多行文本且下方是数值密集区 → header region
- 包含“合计/小计/总计/累计” → total rows
- 空行/空列/边框变化 → section boundary

### 猜想 C：MVP 验证方式

构造一个合成生产风格报表，做闭环：

```text
create_sample.xlsx → parse_to_ir.json → rebuild.xlsx → diff(original, rebuilt)
```

若 diff 发现重建缺失，就实时补充 IR 字段或修复重建逻辑。

## 3. 第一版 IR 草案

```json
{
  "schema_version": "0.1",
  "workbook": {
    "sheets": [
      {
        "name": "销售日报",
        "dimensions": {"max_row": 20, "max_col": 10},
        "freeze_panes": "A8",
        "merged_ranges": ["A1:J1"],
        "rows": {"1": {"height": 28, "hidden": false}},
        "cols": {"A": {"width": 12, "hidden": false}},
        "cells": {
          "A1": {
            "value": "销售日报",
            "data_type": "s",
            "style_id": "s001"
          }
        },
        "logical": {
          "regions": []
        }
      }
    ],
    "styles": {
      "s001": {"font": {}, "fill": {}, "border": {}, "alignment": {}, "number_format": "General"}
    }
  }
}
```

## 4. MVP 边界

第一版支持：

- `.xlsx`
- 普通值与公式文本
- 样式、合并、行高列宽、冻结窗格
- 简单逻辑区域识别
- JSON IR 修改后重建

第一版暂不支持：

- 图表、图片、形状
- 宏/VBA
- 透视表
- 条件格式
- 数据验证
- 批注
- named range
- 外部链接

后续根据 diff 再补。
