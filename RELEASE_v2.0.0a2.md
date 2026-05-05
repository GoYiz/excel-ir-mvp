# Release v2.0.0a2 - Semantic Metadata Persistence

This alpha closes the loop for semantic table metadata by embedding it into rebuilt XLSX files and adding explicit export/import commands.

## Highlights

- Added compact semantic metadata payload:
  - `kind: excel_ir_semantic_metadata`
  - `version: 1`
  - `table_kind`
  - table identity / `ref`
  - generated or confirmed `ir.field_map_candidates`
  - native/semantic support status
- Rebuild now writes metadata to a very-hidden worksheet:
  - `_excel_ir_metadata!A1`
- Parse reads the hidden sheet and merges metadata back into normal IR.
- Hidden metadata carrier sheet is stripped from parsed business sheets.
- Added public API helpers:
  - `collect_semantic_metadata`
  - `apply_semantic_metadata`
  - `export_semantic_metadata_from_ir`
  - `import_semantic_metadata_to_ir`
- Added CLI:
  - `excel-ir metadata export ir.json metadata.json`
  - `excel-ir metadata import stripped.ir.json metadata.json restored.ir.json`
- Updated IR schema with `semantic_metadata` and `table_kind` enum.
- Updated `ARCHITECTURE.md` and `CLI_REFERENCE.md`.

## Commands

```bash
excel-ir metadata export out.ir.json semantic_metadata.json
excel-ir metadata import stripped.ir.json semantic_metadata.json restored.ir.json
```

## Hidden sheet behavior

When rebuilding:

```text
_excel_ir_metadata!A1
```

stores minified JSON metadata, and the sheet state is:

```text
veryHidden
```

When parsing a rebuilt workbook, the metadata is merged into `sheet.extra.tables[*]`, but `_excel_ir_metadata` itself is not exposed as a normal IR worksheet.

## Tests

```bash
python3 -m unittest -v \
  tests.test_excel_ir_mvp \
  tests.test_patch_ops \
  tests.test_native_tables \
  tests.test_metadata
```

Result:

```text
Ran 21 tests ... OK
```

New metadata tests verify:

1. Hidden metadata sheet is written as `veryHidden`.
2. Re-parse restores `table_kind: semantic` and field map metadata.
3. Hidden carrier sheet is not exposed as a business sheet.
4. API export/import restores stripped metadata.
5. CLI export/import works.

## CI

```bash
python3 ci_check.py
python3 ci_check.py --installed
```

Source CI passed with coverage gate:

```bash
python3 -m coverage report --show-missing --fail-under=60
```

Observed coverage:

```text
TOTAL 2631 statements, 912 missing, 65% coverage
```

Installed CI passed after installing the built wheel and now includes:

```bash
excel-ir metadata export tests/fixtures/complex_ir_v07.json ci_installed_metadata.json
```

## Build

Built artifacts:

- `dist/excel_ir_mvp-2.0.0a2-py3-none-any.whl`
- `dist/excel_ir_mvp-2.0.0a2.tar.gz`

Twine check passed for both artifacts.

## SHA256

```text
1635b146e355fd92fac24a9762f8ac5dfb5c332d6b93ec251f8fcbae7a6ba488  dist/excel_ir_mvp-2.0.0a2-py3-none-any.whl
578ce2909aa6f893df2fc04e1852dcec005425053e40ced0e53bf90e5aa699a6  dist/excel_ir_mvp-2.0.0a2.tar.gz
```
