# Release v2.0.0a17

Remove old top-level API compatibility aliases and add large-workbook parse performance controls.

## Breaking change

Old helper names are no longer exported from `excel_ir_mvp.__init__`:

- `parse_workbook_plus`
- `rebuild_workbook_plus`
- `diff_workbooks_plus`
- `stream_update_first_match_xlsx`
- metadata-specific helper aliases

Use the concise facade instead:

```python
import excel_ir_mvp as xir

ir = xir.parse("workbook.xlsx", sheets="Data")
xir.rebuild(ir, "rebuilt.xlsx")
```

## Performance work

Added `profile="fast"` / `excel-ir parse --fast` for large workbooks.

Fast profile skips the main expensive parts:

- read-only streaming mode where supported
- empty styled cells
- formula cache workbook loading
- logical inference
- extended OOXML extras
- images and binary payloads
- charts
- hidden semantic metadata merge

Also added fine-grained parse flags:

```bash
excel-ir parse large.xlsx out.ir.json \
  --sheet Data \
  --no-formula-cache \
  --no-extra \
  --no-images \
  --no-charts \
  --no-binary
```

Core parser optimizations:

- sparse cell iteration by default using existing worksheet cells instead of rectangular `max_row × max_column` scans
- style id caching by openpyxl style id to avoid repeated style dict serialization
- formula cache workbook is loaded only when requested
- image binary loading/base64 can be skipped

## Why large XLSX was slow

The primary causes were multiple workbook loads, rectangular empty-cell scans on inflated dimensions, per-cell style serialization/deduplication, image base64 embedding, extended OOXML object parsing, and dense logical inference structures.

## Docs

- `docs/performance.md`
- `docs/public-api.md` updated for removed old top-level API

PyPI publishing remains skipped.
