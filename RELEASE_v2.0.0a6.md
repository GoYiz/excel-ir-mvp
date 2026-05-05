# Release v2.0.0a6 - Repository UX and XLSX Metadata Verify

This alpha improves open-source repository usability and adds direct XLSX metadata verification.

## Highlights

- Added README GitHub Actions / MIT badges.
- Added community files:
  - `SECURITY.md`
  - `CODE_OF_CONDUCT.md`
  - `.github/ISSUE_TEMPLATE/bug_report.md`
  - `.github/ISSUE_TEMPLATE/feature_request.md`
- Added metadata verification from XLSX input:

```bash
excel-ir metadata verify --from-xlsx rebuilt.xlsx
```

- Added metadata API helpers:
  - `extract_semantic_metadata_from_xlsx`
  - `verify_semantic_metadata_xlsx`
- Added corpus subcommands:

```bash
excel-ir corpus list --config corpus_config.json
excel-ir corpus run --config corpus_config.json
excel-ir corpus report corpus_results/summary.json corpus_report.html
```

- Corpus results now record `metadata_verify_xlsx_ok`.
- Added HTML corpus report rendering.
- Synced package and test corpus configs with categories.
- Bumped version to `2.0.0a6`.

## Validation

```bash
python3 -m unittest -v \
  tests.test_excel_ir_mvp \
  tests.test_patch_ops \
  tests.test_native_tables \
  tests.test_metadata
python3 ci_check.py
python3 ci_check.py --installed
```

## Build

Artifacts:

- `dist/excel_ir_mvp-2.0.0a6-py3-none-any.whl`
- `dist/excel_ir_mvp-2.0.0a6.tar.gz`

## Notes

This release targets repository UX and operational tooling. Core IR semantics remain compatible with the v2.0 alpha series.
