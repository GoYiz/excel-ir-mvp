# Excel IR MVP

[![CI](https://github.com/GoYiz/excel-ir-mvp/actions/workflows/ci.yml/badge.svg)](https://github.com/GoYiz/excel-ir-mvp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Python package for parsing complex human-authored Excel reports into an abstract IR, applying targeted/semantic changes, rebuilding `.xlsx`, and auditing the result.

Repository: https://github.com/GoYiz/excel-ir-mvp

Current prerelease: **2.0.0a17**. PyPI publishing is intentionally skipped for now; use GitHub releases or install from source.

## Install

```bash
git clone https://github.com/GoYiz/excel-ir-mvp.git
cd excel-ir-mvp
python3 -m pip install -e .
```

## Python quick start

```python
import excel_ir_mvp as xir

ir = xir.parse("tests/fixtures/complex_report.xlsx", sheets="经营驾驶舱")
xir.rebuild(ir, "rebuilt.xlsx")
print(xir.diff("tests/fixtures/complex_report.xlsx", "rebuilt.xlsx"))
```

Targeted edits without learning the full IR:

```python
from excel_ir_mvp import HeaderEditOptions, StreamEditOptions

xir.header_edit(
    "tests/fixtures/multi_header_dates.xlsx",
    "edited.xlsx",
    headers=["2026", "5", "8"],
    value=999,
    options=HeaderEditOptions(row_match="门店A"),
)

xir.stream_edit(
    "tests/fixtures/complex_report.xlsx",
    "edited.xlsx",
    match="总计",
    value="合计",
)
```

## CLI quick start

```bash
excel-ir doctor
excel-ir engines
excel-ir parse tests/fixtures/complex_report.xlsx out.ir.json --sheet 经营驾驶舱
excel-ir parse large.xlsx fast.ir.json --sheet Data --fast
excel-ir rebuild out.ir.json rebuilt.xlsx --sheet 经营驾驶舱
excel-ir diff tests/fixtures/complex_report.xlsx rebuilt.xlsx diff.json
excel-ir header-edit tests/fixtures/multi_header_dates.xlsx edited.xlsx --headers 2026/5/8 --row-match 门店A --value 999 --as-number
excel-ir stream-edit tests/fixtures/complex_report.xlsx edited.xlsx --match 总计 --value 合计
```

## Main concepts

- **Fidelity IR**: cells, styles, layout, merges, formulas, charts/images/tables and print settings.
- **Semantic IR**: table kind, field maps, metadata, patch history and audit-friendly change intent.
- **Compact/selective IR**: parse only selected sheets and omit default values to keep JSON smaller.
- **Public facade**: `excel_ir_mvp.api` / top-level `import excel_ir_mvp as xir` expose the small API for third-party users; historical helpers remain for compatibility.

## Docs

- [Public API](docs/public-api.md)
- [Performance Guide](docs/performance.md)
- [Project Structure](docs/project-structure.md)
- [Backend Engines](docs/backends.md)
- [Selective Sheets and Compact IR](docs/selective-sheets-compact-ir.md)
- [Multi-level Header Edits](docs/multi-header-edits.md)
- [Streaming Edits](docs/streaming-edits.md)
- [Native vs Semantic Tables](docs/native-vs-semantic-tables.md)
- [Metadata Commands](docs/metadata.md)
- [Anonymization](docs/anonymization.md)

## Development checks

```bash
python3 -m unittest -v tests.test_excel_ir_mvp tests.test_patch_ops tests.test_native_tables tests.test_metadata
python3 ci_check.py
python3 -m build --sdist --wheel
python3 -m twine check dist/*
```

## License

MIT.
