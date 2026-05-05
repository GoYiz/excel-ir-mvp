# Release v2.0.0a3 - Metadata Checksum and Diff

This alpha strengthens semantic metadata persistence with checksums, metadata diffing, confirmed field-map persistence tests, and a stricter coverage gate.

## Highlights

- Metadata payload version bumped to `2`.
- Added SHA-256 checksum support:
  - `metadata_checksum`
  - `attach_metadata_checksum`
  - `verify_metadata_checksum`
- Hidden sheet reader rejects corrupted checksummed metadata instead of merging it silently.
- Added metadata diff API:
  - `semantic_metadata_diff`
  - `semantic_metadata_diff_files`
- Added CLI:
  - `excel-ir metadata diff a.json b.json diff.json`
- `collect_semantic_metadata` now migrates old IR table metadata:
  - if `table_kind` is absent but multi-row header / field-map metadata exists, persist it as `table_kind: semantic`.
- Added regression test that `confirm_field_map` results are persisted into semantic metadata.
- Coverage gate raised from 60% to 65% in `ci_check.py` and GitHub Actions.

## Commands

```bash
excel-ir metadata export out.ir.json semantic_metadata.json
excel-ir metadata import stripped.ir.json semantic_metadata.json restored.ir.json
excel-ir metadata diff a.semantic.json b.semantic.json metadata_diff.json
```

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
Ran 22 tests ... OK
```

New/updated metadata assertions verify:

1. Metadata includes `checksum`.
2. Checksum verifies successfully.
3. CLI/API metadata diff returns `diff_count: 0` for identical semantic payloads.
4. `confirm_field_map` patch output persists confirmed mappings such as `本月收入 -> C` and `评级 -> K`.
5. Old IR fixtures without `table_kind` are migrated to semantic table metadata when field-map/multi-row-header evidence exists.

## CI

```bash
python3 ci_check.py
python3 ci_check.py --installed
```

Source CI passed with stricter coverage gate:

```bash
python3 -m coverage report --show-missing --fail-under=65
```

Observed coverage:

```text
TOTAL 2691 statements, 929 missing, 65% coverage
```

Installed CI passed and now includes:

```bash
excel-ir metadata export tests/fixtures/complex_ir_v07.json ci_installed_metadata.json
excel-ir metadata diff ci_installed_metadata.json ci_installed_metadata.json ci_installed_metadata_diff.json
```

## Build

Built artifacts:

- `dist/excel_ir_mvp-2.0.0a3-py3-none-any.whl`
- `dist/excel_ir_mvp-2.0.0a3.tar.gz`

Twine check passed for both artifacts.

## SHA256

```text
65d4143c3717df17b9e222542025c82a1c01041d2717fb61a9e6e5c810d0ef7b  dist/excel_ir_mvp-2.0.0a3-py3-none-any.whl
938eb16cd69c0ab52ade932fd7abbb1746e3f9e72ef448cf16b5dd873168fc47  dist/excel_ir_mvp-2.0.0a3.tar.gz
```
