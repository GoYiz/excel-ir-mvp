# v2.0.0a10 Anonymize, Metadata Status, and Compare Modes

Date: 2026-05-05

## Highlights

- Added workbook anonymization:
  ```bash
  excel-ir anonymize workbook.xlsx anonymized.xlsx
  excel-ir anonymize workbook.xlsx anonymized.xlsx --rewrite-formulas
  ```
- Added metadata status:
  ```bash
  excel-ir metadata status workbook.xlsx
  ```
- Added compare modes:
  ```bash
  excel-ir compare-ir --semantic-only a.ir.json b.ir.json semantic_diff.json
  excel-ir compare-ir --structural-only a.ir.json b.ir.json structural_diff.json
  ```
- Added docs:
  - `docs/anonymization.md`
  - `docs/metadata.md`
- CI covers anonymize/status and compare modes.

## Publishing

PyPI publishing intentionally skipped. GitHub release assets are used for this alpha line.
