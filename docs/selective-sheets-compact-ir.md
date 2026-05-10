# Selective Sheets and Compact IR

Large human-authored workbooks often contain many helper sheets, chart data sheets, or archived tabs. v2.0.0a15 lets callers parse and rebuild only the sheets they need.

## CLI

Parse one sheet:

```bash
excel-ir parse workbook.xlsx selected.ir.json --sheet 经营驾驶舱
```

Parse several sheets:

```bash
excel-ir parse workbook.xlsx selected.ir.json --sheet Sheet1 --sheet Sheet2
```

Rebuild only selected sheets from a larger IR:

```bash
excel-ir rebuild full.ir.json selected.xlsx --sheet Sheet1 --sheet Sheet2
```

## API

```python
from excel_ir_mvp import parse, rebuild

ir = parse("workbook.xlsx", sheets=["经营驾驶舱"])
rebuild(ir, "selected.xlsx", sheets=["经营驾驶舱"])
```

The resulting IR records:

```json
{
  "workbook": {
    "sheet_names": ["经营驾驶舱"],
    "selected_sheets": ["经营驾驶舱"]
  }
}
```

## Compact IR changes

v2.0.0a15 removes redundant/default fields to reduce JSON size:

- Formula cells keep `computed_value` but no longer emit `computed_value_source`.
- Default protection such as `locked: true` and `hidden: false` is omitted.
- Default row/column dimension values such as `hidden: false` and `outlineLevel: 0` are omitted.
- Default border/alignment/sheet view/format/print option flags are omitted when they match Excel/openpyxl defaults.

Rebuild functions keep backward compatibility with older IR that still contains those fields.
