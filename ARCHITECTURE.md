# Excel IR MVP v2 Architecture

```text
             ┌──────────────┐
             │  XLSX files  │
             └──────┬───────┘
                    │ parse
                    ▼
        ┌────────────────────────┐
        │ Excel IR JSON           │
        │ - cells/styles/layout   │
        │ - tables/charts/images  │
        │ - table_kind metadata   │
        │ - field_map candidates  │
        └──────┬─────────────────┘
               │ patch validation + dry-run
               ▼
        ┌────────────────────────┐
        │ Semantic Patch Actions  │
        │ - confirm_field_map     │
        │ - append/update/delete  │
        │ - recompute_totals      │
        └──────┬─────────────────┘
               │ transactional apply
               ▼
        ┌────────────────────────┐
        │ Patched IR              │
        │ - patch_history         │
        │ - semantic metadata     │
        └──────┬─────────────────┘
               │ rebuild
               ▼
             ┌──────────────┐
             │ Patched XLSX │
             │ + veryHidden │
             │   metadata   │
             └──────┬───────┘
                    │ diff / audit / reparse
                    ▼
        ┌────────────────────────┐
        │ Reports                 │
        │ - diff HTML             │
        │ - audit HTML            │
        │ - corpus summary        │
        └────────────────────────┘
```

## Main Modules

- `excel_ir_plus.py`: parse/rebuild/diff engine, semantic metadata embedding.
- `ir_patch.py`: semantic patch engine, dry-run, transaction log.
- `diff_report.py`: diff report with plan and tx log.
- `audit_report.py`: standalone audit report.
- `corpus_runner.py`: config-driven corpus tests.
- `excel_ir_cli.py`: unified CLI.

## Native table vs semantic table

v2 distinguishes two table kinds:

- `table_kind: native`: a safe single-row Excel native Table. It is rebuilt as an OOXML Table.
- `table_kind: semantic`: a complex human-report table, typically with merged or multi-level headers. It remains semantic IR and is not forced into an OOXML Table, avoiding `column headings must be strings` warnings.

Both paths retain the reversible cell grid, styles, filters, formulas, and semantic field maps.

## Semantic metadata persistence

Rebuilding an IR writes compact semantic table metadata to a very-hidden worksheet:

```text
_excel_ir_metadata!A1
```

Payload kind:

```json
"excel_ir_semantic_metadata"
```

This lets XLSX-only pipelines preserve:

- `table_kind`
- table `ref` / display name
- generated or confirmed `field_map_candidates`
- native/semantic support status

CLI helpers:

```bash
excel-ir metadata export out.ir.json semantic_metadata.json
excel-ir metadata import stripped.ir.json semantic_metadata.json restored.ir.json
excel-ir metadata diff a.semantic.json b.semantic.json metadata_diff.json
excel-ir metadata verify semantic_metadata.json
excel-ir metadata verify --from-xlsx workbook.xlsx
```

v2.0.0a3 metadata includes a SHA-256 `checksum` over the canonical payload. Parse rejects corrupted checksummed metadata rather than merging it silently.

## CLI Examples

```bash
python3 excel_ir_cli.py parse input.xlsx out.ir.json
python3 excel_ir_cli.py rebuild out.ir.json rebuilt.xlsx
python3 excel_ir_cli.py diff input.xlsx rebuilt.xlsx diff.json
python3 excel_ir_cli.py patch out.ir.json patch.json patched.ir.json --plan plan.json --log tx.json
python3 excel_ir_cli.py metadata export out.ir.json semantic_metadata.json
python3 excel_ir_cli.py metadata import stripped.ir.json semantic_metadata.json restored.ir.json
python3 excel_ir_cli.py report diff.json report.html --plan plan.json --log tx.json
python3 excel_ir_cli.py audit tx.json audit.html
python3 excel_ir_cli.py corpus --config corpus_config.json
```
