# Excel IR MVP v2 Architecture

`excel-ir-mvp` is organized around a small public facade and layered internal engines.

```text
Third-party users
  │
  ▼
excel_ir_mvp.__init__ / api.py
  - parse / rebuild / diff / inspect
  - stream_edit / header_edit / anonymize
  - option dataclasses in types.py
  │
  ├──────────────► backends.py
  │                 openpyxl default, wolfxl optional
  │
  ▼
Fidelity IR engine
  excel_ir.py
  - cells/styles/layout/merges/formulas
  - selective sheet parse/rebuild
  - compact default omission
  │
  ▼
Extended workbook engine
  excel_ir_plus.py
  - tables/charts/images/validations/metadata
  - veryHidden semantic metadata sheet
  │
  ├──────────────► ir_patch.py
  │                 action patches, dry-run, transaction log
  │
  └──────────────► reports / corpus / validation
                    diff_report.py, audit_report.py,
                    corpus_runner.py, validate_ir.py
```

## Public package surface

New code should start with:

```python
import excel_ir_mvp as xir

ir = xir.parse("workbook.xlsx", sheets="经营驾驶舱")
xir.rebuild(ir, "rebuilt.xlsx")
```

The stable facade is intentionally small:

- `parse`, `rebuild`, `diff`, `compare_ir`, `inspect`
- `apply_patch`
- `stream_edit`, `header_edit`, `header_columns`
- `anonymize`, `engines`
- option dataclasses: `ParseOptions`, `RebuildOptions`, `StreamEditOptions`, `HeaderEditOptions`

Historical helpers such as `parse_workbook_plus`, `rebuild_workbook_plus`, `stream_update_first_match_xlsx`, and metadata-specific helpers remain available for compatibility and advanced workflows, but they are not the primary learning path.

## Two-layer IR

### Fidelity IR

A reversible workbook representation:

- sheet names and dimensions
- cells and formulas
- styles and compact style dictionaries
- row/column layout
- merged ranges
- selected sheet subset when requested

### Semantic IR

Intent and domain metadata over the physical grid:

- `table_kind: native | semantic`
- field-map candidates and confirmed field maps
- semantic metadata persisted in a very-hidden `_excel_ir_metadata` sheet
- patch history and audit-friendly impact information

## Native vs semantic tables

- `native`: safe single-row Excel native table rebuilt as OOXML Table.
- `semantic`: complex human-report table, often merged or multi-row headers; preserved semantically without forcing unsafe native table reconstruction.

## Backend model

`backends.py` provides a registry/probe layer:

- `openpyxl`: default full-fidelity engine.
- `wolfxl`: optional compatible engine when installed/available.
- `auto`: choose available accelerated engine when possible, otherwise openpyxl.

## CLI adapter

`excel_ir_cli.py` maps CLI commands to the public/internal APIs while keeping JSON-oriented output for automation.
