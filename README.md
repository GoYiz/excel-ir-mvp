# Excel IR MVP

Current version: `2.0.0a5`.

Repository: <https://github.com/GoYiz/excel-ir-mvp>

A Python MVP for parsing complex human Excel reports into an intermediate representation (IR), applying semantic patches, rebuilding XLSX workbooks, and generating validation/audit artifacts.

## Quick start

```bash
git clone https://github.com/GoYiz/excel-ir-mvp.git
cd excel-ir-mvp
python3 -m pip install -e .
python3 -m unittest -v tests.test_excel_ir_mvp tests.test_patch_ops tests.test_native_tables tests.test_metadata
python3 ci_check.py
```

## Installed commands

```bash
excel-ir doctor
python3 -m excel_ir_mvp doctor
excel-ir validate ir tests/fixtures/complex_ir_v07.json
excel-ir validate patch tests/fixtures/v08_patch.json
```

## v2 semantic tables

The v2 alpha line distinguishes two table kinds:

- `table_kind: native` — safe single-row Excel native Tables are rebuilt as OOXML Tables.
- `table_kind: semantic` — complex human-report tables with merged/multi-level headers remain semantic IR tables and are not forced into native Excel Table objects.

This avoids openpyxl's warning:

```text
column headings must be strings
```

while preserving cell grid, styles, auto filters, formulas, and semantic `field_map` metadata for patching.

## Semantic metadata persistence

Rebuilt XLSX files embed compact semantic metadata in a very-hidden worksheet:

```text
_excel_ir_metadata!A1
```

v2 metadata includes a SHA-256 checksum and supports export/import/diff/verify:

```bash
excel-ir metadata export out.ir.json semantic_metadata.json
excel-ir metadata import stripped.ir.json semantic_metadata.json restored.ir.json
excel-ir metadata diff a.semantic.json b.semantic.json metadata_diff.json
excel-ir metadata verify semantic_metadata.json
```

## Corpus categories

Corpus summaries now include category rollups such as:

- `synthetic_complex`
- `metadata_roundtrip`

See [tests/fixtures/README.md](tests/fixtures/README.md) for fixture conventions.

## Validation status

- Source CI: pass.
- Installed CI: pass.
- Tests: 26.
- Coverage: 69% with `--fail-under=65`.
- Build: wheel + sdist.
- Twine check: pass.

## Latest release artifact names

- `excel_ir_mvp-2.0.0a5-py3-none-any.whl`
- `excel_ir_mvp-2.0.0a5.tar.gz`
