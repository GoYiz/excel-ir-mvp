# Excel IR MVP

[![CI](https://github.com/GoYiz/excel-ir-mvp/actions/workflows/ci.yml/badge.svg)](https://github.com/GoYiz/excel-ir-mvp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Python prototype for parsing complex human-authored Excel reports into an abstract IR, applying semantic changes, rebuilding `.xlsx`, and auditing the result.

Repository: https://github.com/GoYiz/excel-ir-mvp

Current prerelease: **2.0.0a13**. PyPI publishing is intentionally skipped for now; use GitHub releases or install from source.

## Install

```bash
git clone https://github.com/GoYiz/excel-ir-mvp.git
cd excel-ir-mvp
python3 -m pip install -e .
```

## Quick start

```bash
excel-ir doctor
excel-ir engines
excel-ir parse tests/fixtures/complex_report.xlsx out.ir.json --engine openpyxl
excel-ir stream-edit tests/fixtures/complex_report.xlsx edited.xlsx --match 总计 --value 合计
excel-ir stream-edit tests/fixtures/complex_report.xlsx edited.xlsx --match 云业务 --value 云事业部 --all
excel-ir stream-edit tests/fixtures/complex_report.xlsx ignored.xlsx --match 业务线 --value 收入本月 --offset-row 1 --offset-col 2 --preview
excel-ir inspect tests/fixtures/complex_report.xlsx --out inspect.json
excel-ir parse tests/fixtures/complex_report.xlsx out.ir.json
excel-ir rebuild out.ir.json rebuilt.xlsx
excel-ir diff tests/fixtures/complex_report.xlsx rebuilt.xlsx diff.json
excel-ir compare-ir out.ir.json out.ir.json ir_diff.json
```

## Useful commands

```bash
excel-ir anonymize private.xlsx anonymized.xlsx
excel-ir metadata status workbook.xlsx
excel-ir compare-ir --semantic-only a.ir.json b.ir.json semantic_diff.json
excel-ir compare-ir --structural-only a.ir.json b.ir.json structural_diff.json
```

## Docs

- [Backend Engines](docs/backends.md)
- [Streaming Edits](docs/streaming-edits.md)
- [Native vs Semantic Tables](docs/native-vs-semantic-tables.md)
- [Metadata Commands](docs/metadata.md)
- [Anonymization](docs/anonymization.md)

## Corpus

```bash
excel-ir corpus list --config tests/fixtures/corpus_config.json
excel-ir corpus run --config tests/fixtures/corpus_config.json
excel-ir corpus report corpus_results/summary.json corpus_results/report.html
```

Current categories: `synthetic_complex`, `metadata_roundtrip`, `native_table`, `semantic_table`.

## Development checks

```bash
python3 -m unittest -v tests.test_excel_ir_mvp tests.test_patch_ops tests.test_native_tables tests.test_metadata
python3 ci_check.py
python3 -m build --sdist --wheel
python3 -m twine check dist/*
```

## License

MIT.
