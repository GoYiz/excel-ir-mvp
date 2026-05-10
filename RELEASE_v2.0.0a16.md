# Release v2.0.0a16

Public API and project structure cleanup for third-party package usability.

## Highlights

- Added `excel_ir_mvp.api` as a concise facade for common workflows.
- Added option dataclasses in `excel_ir_mvp.types`:
  - `ParseOptions`
  - `RebuildOptions`
  - `StreamEditOptions`
  - `HeaderEditOptions`
- Reworked top-level package exports so `import excel_ir_mvp as xir` exposes a small, understandable API:
  - `parse`
  - `rebuild`
  - `diff`
  - `compare_ir`
  - `inspect`
  - `apply_patch`
  - `stream_edit`
  - `header_edit`
  - `header_columns`
  - `anonymize`
  - `engines`
- Kept older `parse_workbook_plus`, `rebuild_workbook_plus`, `*_xlsx`, metadata and engine helpers as compatibility/advanced APIs.
- Added docs:
  - `docs/public-api.md`
  - `docs/project-structure.md`
- Rewrote README and architecture docs around the new public facade and layered design.

## Example

```python
import excel_ir_mvp as xir

ir = xir.parse("workbook.xlsx", sheets="经营驾驶舱")
xir.rebuild(ir, "rebuilt.xlsx")

result = xir.header_edit(
    "workbook.xlsx",
    "edited.xlsx",
    headers=["2026", "5", "8"],
    value=999,
    options=xir.HeaderEditOptions(row_match="门店A"),
)
```

## Compatibility

No removal of legacy public helpers in this release. Existing imports should continue working.

## Validation

- Source CI: expected full suite + coverage gate.
- Build: wheel + sdist.
- Installed CI: package facade and legacy CLI remain covered.

PyPI publishing remains skipped.
