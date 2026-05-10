# Project Structure

`excel-ir-mvp` now separates the small user-facing facade from the historical implementation modules.

```text
excel_ir_mvp/
  __init__.py          # stable public exports
  api.py               # concise third-party package facade
  types.py             # option dataclasses and type aliases
  backends.py          # backend registry: openpyxl, wolfxl optional
  excel_ir.py          # core reversible cell/style/layout IR
  excel_ir_plus.py     # extended OOXML objects + semantic metadata
  ir_patch.py          # semantic patching and transaction logs
  excel_ir_cli.py      # CLI adapter around public/internal APIs
  validate_ir.py       # JSON schema + semantic validation helpers
  models.py            # lightweight IR model validation
  corpus_runner.py     # corpus smoke tests and report generation
  diff_report.py       # HTML diff reports
  audit_report.py      # HTML audit reports
```

## Design principles

1. **Small public surface**: `import excel_ir_mvp as xir` should be enough for common parse/rebuild/diff/edit workflows.
2. **Compatibility over churn**: older `*_xlsx` and `*_plus` helpers remain available, but are treated as advanced/compatibility APIs.
3. **Two-layer IR**: fidelity IR preserves workbook structure; semantic IR stores table intent, metadata and patch history.
4. **Action-oriented edits**: high-level changes are represented as patch actions or targeted edit calls, not raw JSON mutation.
5. **Backend abstraction**: `openpyxl` remains default; optional engines can be registered/probed without changing the public facade.

## Recommended entry point

```python
import excel_ir_mvp as xir

ir = xir.parse("workbook.xlsx", sheets="经营驾驶舱")
xir.rebuild(ir, "rebuilt.xlsx")
```

Use submodules only when you need implementation details or compatibility with older scripts.
