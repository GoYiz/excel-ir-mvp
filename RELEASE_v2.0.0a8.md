# v2.0.0a8 Inspect, Repair, and CI Corpus Artifact

Date: 2026-05-05

## Highlights

- Added workbook overview command:
  ```bash
  excel-ir inspect workbook.xlsx --out inspect.json
  ```
- Added semantic metadata repair command:
  ```bash
  excel-ir metadata repair repaired.xlsx --from-xlsx workbook.xlsx
  ```
- Added API helpers:
  - `inspect_workbook(path)`
  - `repair_semantic_metadata_xlsx(src_path, out_path)`
- `corpus run` now writes both `summary.json` and `report.html`.
- GitHub Actions uploads corpus/inspect outputs as a `corpus-report` artifact.

## Validation

- Source CI: `python3 ci_check.py`
- Installed CI: `python3 ci_check.py --installed`
- Unit tests include inspect/repair regressions.
- Build and `twine check` remain part of release validation.

## Publishing

PyPI publishing intentionally skipped. GitHub release assets are used for this alpha line.
