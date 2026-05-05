# Excel IR MVP

[![CI](https://github.com/GoYiz/excel-ir-mvp/actions/workflows/ci.yml/badge.svg)](https://github.com/GoYiz/excel-ir-mvp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Python prototype for parsing complex human-authored Excel reports into an abstract IR, applying semantic changes, rebuilding `.xlsx`, and auditing the result.

Repository: https://github.com/GoYiz/excel-ir-mvp

Current prerelease: **2.0.0a9**. PyPI publishing is intentionally skipped for now; use GitHub releases or install from source.

## Install

```bash
git clone https://github.com/GoYiz/excel-ir-mvp.git
cd excel-ir-mvp
python3 -m pip install -e .
```

## Quick start

```bash
excel-ir doctor
excel-ir inspect tests/fixtures/complex_report.xlsx --out inspect.json
excel-ir parse tests/fixtures/complex_report.xlsx out.ir.json
excel-ir rebuild out.ir.json rebuilt.xlsx
excel-ir diff tests/fixtures/complex_report.xlsx rebuilt.xlsx diff.json
excel-ir compare-ir out.ir.json out.ir.json ir_diff.json
```

## Semantic metadata

```bash
excel-ir metadata export out.ir.json metadata.json
excel-ir metadata import stripped.ir.json metadata.json restored.ir.json
excel-ir metadata extract metadata.json --from-xlsx rebuilt.xlsx
excel-ir metadata verify metadata.json
excel-ir metadata verify --from-xlsx rebuilt.xlsx
excel-ir metadata repair repaired.xlsx --from-xlsx workbook.xlsx
excel-ir metadata strip stripped.xlsx --from-xlsx workbook.xlsx
excel-ir metadata diff a.metadata.json b.metadata.json metadata_diff.json
```

Semantic table metadata is stored in a `veryHidden` sheet named `_excel_ir_metadata` with a SHA-256 checksum. See [Native vs Semantic Tables](docs/native-vs-semantic-tables.md).

## Corpus

```bash
excel-ir corpus list --config tests/fixtures/corpus_config.json
excel-ir corpus run --config tests/fixtures/corpus_config.json
excel-ir corpus report corpus_results/summary.json corpus_results/report.html
```

Current categories: `synthetic_complex`, `metadata_roundtrip`, `native_table`, `semantic_table`.

CI writes `corpus_results/summary.json`, `corpus_results/report.html`, and `ci_inspect.json`; GitHub Actions uploads them as a `corpus-report` artifact.

## Development checks

```bash
python3 -m unittest -v tests.test_excel_ir_mvp tests.test_patch_ops tests.test_native_tables tests.test_metadata
python3 ci_check.py
python3 -m build --sdist --wheel
python3 -m twine check dist/*
```

## License

MIT.
