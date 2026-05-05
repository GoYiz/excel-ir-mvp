# Release v2.0.0a4 - Metadata Verify and Corpus Categories

This alpha adds a metadata verification command, a corrupted hidden-metadata regression test, and category-level corpus summaries.

## Highlights

- Added semantic metadata verification API:
  - `verify_semantic_metadata`
  - `verify_semantic_metadata_file`
- Added CLI:
  - `excel-ir metadata verify semantic_metadata.json`
- Hidden metadata checksum corruption is now explicitly covered by regression tests.
- Corpus runner now emits category rollups:
  - `synthetic_complex`
  - `metadata_roundtrip`
  - future categories such as `native_table` / `semantic_table` / `real_world`
- Corpus sample config supports `category` per sample.
- Source CI coverage collection changed to split/parallel coverage runs plus `coverage combine`, avoiding long iSH run data loss and raising observed coverage to 69%.
- Coverage gate remains strict at `--fail-under=65`.

## Commands

```bash
excel-ir metadata export out.ir.json semantic_metadata.json
excel-ir metadata import stripped.ir.json semantic_metadata.json restored.ir.json
excel-ir metadata diff a.semantic.json b.semantic.json metadata_diff.json
excel-ir metadata verify semantic_metadata.json
```

`metadata verify` returns:

```json
{
  "ok": true,
  "errors": [],
  "tables": 1,
  "checksum_ok": true
}
```

## Corpus categories

`corpus_runner.py` summary now includes:

```json
{
  "categories": {
    "synthetic_complex": {"count": 1, "ok": 1, "failed": 0, "diff_count_total": 0},
    "metadata_roundtrip": {"count": 1, "ok": 1, "failed": 0, "diff_count_total": 0}
  }
}
```

Each result also records `category`, `metadata_export_ok`, and `metadata_tables`.

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
Ran 26 tests ... OK
```

New/updated tests cover:

1. `metadata verify` CLI/API.
2. In-process metadata CLI dispatch for coverage.
3. Corrupted hidden sheet checksum is ignored.
4. Corpus category rollups.
5. Package `corpus_runner` category helper.

## CI

```bash
python3 ci_check.py
python3 ci_check.py --installed
```

Source CI passed with split coverage:

```bash
python3 -m coverage erase
python3 -m coverage run --parallel-mode --source=src/excel_ir_mvp -m unittest tests.test_excel_ir_mvp tests.test_patch_ops
python3 -m coverage run --parallel-mode --source=src/excel_ir_mvp -m unittest tests.test_native_tables tests.test_metadata
python3 -m coverage combine
python3 -m coverage report --show-missing --fail-under=65
```

Observed coverage:

```text
TOTAL 2745 statements, 864 missing, 69% coverage
```

Installed CI passed and now includes:

```bash
excel-ir metadata verify ci_installed_metadata.json
```

## Build

Built artifacts:

- `dist/excel_ir_mvp-2.0.0a4-py3-none-any.whl`
- `dist/excel_ir_mvp-2.0.0a4.tar.gz`

Twine check passed for both artifacts.

## SHA256

```text
d9df3167f61458d22fd134b8a43e15f38d94f9880d58363ecb6ffc5c59c999eb  dist/excel_ir_mvp-2.0.0a4-py3-none-any.whl
e0aa6bc245704ef0cdfdd152a2244cc2260a2e5cbda236ed5623a3077e08e28e  dist/excel_ir_mvp-2.0.0a4.tar.gz
```
