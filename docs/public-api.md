# Public API

v2.0.0a16 introduces a small facade for third-party package users. New code should prefer this layer instead of importing the historical engine modules directly.

## Recommended imports

```python
import excel_ir_mvp as xir

ir = xir.parse("report.xlsx", sheets="经营驾驶舱")
xir.rebuild(ir, "rebuilt.xlsx")
summary = xir.inspect("report.xlsx")
```

## Stable facade functions

| Function | Purpose |
| --- | --- |
| `parse(path, sheets=None, engine="openpyxl")` | XLSX -> IR. |
| `rebuild(ir, path, sheets=None, engine="openpyxl")` | IR -> XLSX. |
| `diff(a, b, engine="openpyxl")` | Canonical XLSX diff. |
| `compare_ir(a, b, mode="full")` | Compare IR objects. |
| `inspect(path)` | Compact workbook overview. |
| `apply_patch(ir, patch, dry_run=False)` | Semantic patch facade. |
| `stream_edit(src, dst, match=..., value=..., options=...)` | Human-like streaming edit. |
| `header_edit(src, dst, headers=..., value=..., options=...)` | Multi-level header edit. |
| `header_columns(path, header_rows=(1, 3))` | Expand merged/multi-row headers. |
| `anonymize(src, dst)` | Produce shareable redacted workbook. |
| `engines()` | Engine status. |

## Options objects

```python
from excel_ir_mvp import ParseOptions, StreamEditOptions, HeaderEditOptions

ir = xir.parse_with_options("report.xlsx", ParseOptions(sheets=["Sheet1", "Sheet2"]))

preview = xir.stream_edit(
    "report.xlsx",
    "ignored.xlsx",
    match="业务线",
    value="收入本月",
    options=StreamEditOptions(offset_row=1, offset_col=2, preview=True),
)

result = xir.header_edit(
    "report.xlsx",
    "edited.xlsx",
    headers=["2026", "5", "8"],
    value=999,
    options=HeaderEditOptions(row_match="门店A"),
)
```

## Compatibility

The earlier functions such as `parse_workbook_plus`, `rebuild_workbook_plus`, `stream_update_first_match_xlsx`, and metadata helpers remain importable for compatibility and advanced use. They are no longer the recommended first thing to learn.

Internal modules follow this rough convention:

- `api.py`: concise package facade.
- `types.py`: option dataclasses and aliases.
- `excel_ir.py`: core fidelity IR engine.
- `excel_ir_plus.py`: extended OOXML objects and metadata.
- `ir_patch.py`: semantic patch engine.
- `excel_ir_cli.py`: CLI adapter.
