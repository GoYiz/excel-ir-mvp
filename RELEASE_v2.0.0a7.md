# Release v2.0.0a7 - Corpus Fixtures and Release Assets

PyPI publishing is intentionally skipped for this alpha.

## Highlights

- Added CLI:

```bash
excel-ir metadata extract metadata.json --from-xlsx workbook.xlsx
```

- Added standalone corpus fixtures:
  - `tests/fixtures/native_table.xlsx`
  - `tests/fixtures/semantic_table.xlsx`
- Added fixture generator:
  - `tools/make_table_fixtures.py`
- Extended corpus config with categories:
  - `native_table`
  - `semantic_table`
- Corpus report/list coverage now includes native and semantic table samples.
- README and fixture guide updated.
- GitHub release asset upload for wheel/sdist is part of the release process.

## Validation

```bash
python3 ci_check.py
python3 ci_check.py --installed
```

## Build

Artifacts:

- `dist/excel_ir_mvp-2.0.0a7-py3-none-any.whl`
- `dist/excel_ir_mvp-2.0.0a7.tar.gz`

## Notes

This release focuses on corpus breadth, metadata extraction ergonomics, and GitHub release completeness. No PyPI/TestPyPI publish is performed.
