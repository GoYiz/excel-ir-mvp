# Public API

The package intentionally exposes a small top-level API. New code should import the package facade, not historical engine modules.

```python
import excel_ir_mvp as xir

ir = xir.parse("report.xlsx", sheets="经营驾驶舱")
xir.rebuild(ir, "rebuilt.xlsx")
summary = xir.inspect("report.xlsx")
```

## Stable top-level functions

| Function | Purpose |
| --- | --- |
| `parse(path, sheets=None, engine="openpyxl", profile="full")` | XLSX -> IR. |
| `rebuild(ir, path, sheets=None, engine="openpyxl")` | IR -> XLSX. |
| `diff(a, b, engine="openpyxl")` | Canonical XLSX diff. |
| `compare_ir(a, b, mode="full")` | Compare IR objects. |
| `inspect(path)` | Compact workbook overview. |
| `apply_patch(ir, patch, dry_run=False)` | Semantic patch facade. |
| `stream_edit(src, dst, match=..., value=..., options=...)` | Human-like streaming edit. |
| `header_edit(src, dst, headers=..., value=..., options=...)` | Multi-level header edit. |
| `header_columns(path, header_rows=(1, 3))` | Expand merged/multi-row horizontal headers. |
| `header_rows(path, header_cols=(1, 3))` | Expand merged/multi-column vertical headers. |
| `anonymize(src, dst)` | Produce shareable redacted workbook. |
| `engines()` | Engine status. |

## Fast parse profile

For large workbooks with many cells/styles/images, start with:

```python
ir = xir.parse("large.xlsx", sheets="Data", profile="fast")
```

`profile="fast"` skips expensive extras: formula cache workbook loading, empty styled cells, logical inference, extended OOXML extras, charts, images, binary payloads and hidden semantic metadata. Use `profile="full"` when you need maximum round-trip fidelity.

## Options objects

```python
preview = xir.stream_edit(
    "report.xlsx",
    "ignored.xlsx",
    match="业务线",
    value="收入本月",
    options=xir.StreamEditOptions(offset_row=1, offset_col=2, preview=True),
)

ir = xir.parse("report.xlsx")

located = xir.header_locate(
    ir,
    headers=["2026", "5", "8"],
    options=xir.HeaderEditOptions(row_match="门店A"),
)

edited_ir, result = xir.header_edit(
    ir,
    headers=["2026", "5", "8"],
    value=999,
    options=xir.HeaderEditOptions(row_match="门店A"),
)

regex_ir, regex_preview = xir.header_edit(
    ir,
    headers=["202[0-9]", "5", "[78]"],
    value=999,
    options=xir.HeaderEditOptions(match_mode="regex", preview=True),
)

vertical_ir = xir.parse("vertical.xlsx")
vertical_ir2, vertical = xir.header_edit(
    vertical_ir,
    headers=["收入", "线下"],
    value=999,
    options=xir.HeaderEditOptions(orientation="vertical", header_start_col="A", header_end_col="B", min_row=2, col_match="Q2"),
)
xir.rebuild(edited_ir, "edited.xlsx")
```

## Removed old top-level API

As of v2.0.0a17, old names such as `parse_workbook_plus`, `rebuild_workbook_plus`, `stream_update_first_match_xlsx`, and metadata-specific helpers are no longer exported from `excel_ir_mvp.__init__`. Internal modules may still use implementation functions, but third-party users should use the facade above.
