# Excel IR MVP

[![CI](https://github.com/GoYiz/excel-ir-mvp/actions/workflows/ci.yml/badge.svg)](https://github.com/GoYiz/excel-ir-mvp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Excel IR MVP is a Python package for complex, human-authored Excel reports. It parses `.xlsx` files into a structured intermediate representation (IR), supports targeted semantic edits, rebuilds workbooks, compares/audits changes, and provides fast paths for large files.

Repository: https://github.com/GoYiz/excel-ir-mvp

Current prerelease: **2.0.0a18**. PyPI publishing is intentionally skipped for now; install from source or GitHub Releases.

## Why this exists

Many production Excel reports are not simple tables. They contain merged cells, multi-row or multi-column headers, formulas, styled empty templates, charts, images, semantic sections, native tables, and manually edited layouts. This project splits the problem into two layers:

- **Fidelity IR**: reversible physical workbook structure such as cells, styles, formulas, merges, row/column dimensions, tables, charts/images, validations, print settings and workbook metadata.
- **Semantic IR**: higher-level understanding such as table kind, field maps, patch history, audit information and human-oriented edit intent.

The package is optimized for pragmatic automation: parse what you need, modify safely, rebuild, diff, and audit.

## Install

```bash
git clone https://github.com/GoYiz/excel-ir-mvp.git
cd excel-ir-mvp
python3 -m pip install -e .
```

For a built artifact, download the wheel from GitHub Releases and install it:

```bash
python3 -m pip install excel_ir_mvp-2.0.0a18-py3-none-any.whl
```

## Python quick start

```python
import excel_ir_mvp as xir

ir = xir.parse("tests/fixtures/complex_report.xlsx", sheets="经营驾驶舱")
xir.rebuild(ir, "rebuilt.xlsx")
print(xir.diff("tests/fixtures/complex_report.xlsx", "rebuilt.xlsx"))
```

## Public API

The top-level package intentionally exposes a small, third-party-friendly facade:

```python
import excel_ir_mvp as xir

xir.parse(...)
xir.rebuild(...)
xir.diff(...)
xir.compare_ir(...)
xir.inspect(...)
xir.apply_patch(...)
xir.stream_edit(...)
xir.header_edit(...)
xir.header_columns(...)
xir.header_rows(...)
xir.anonymize(...)
xir.engines()
```

Old top-level implementation helpers such as `parse_workbook_plus` are not exported from `excel_ir_mvp`.

## Parse examples

Parse a workbook:

```python
ir = xir.parse("report.xlsx")
```

Parse selected sheets only:

```python
ir = xir.parse("report.xlsx", sheets=["经营驾驶舱", "明细"])
```

Fast parse for large files:

```python
ir = xir.parse("large.xlsx", sheets="Data", profile="fast")
```

Fine-grained performance options:

```python
ir = xir.parse(
    "large.xlsx",
    sheets="Data",
    read_only=True,
    include_formula_cache=False,
    include_extra=False,
    include_images=False,
    include_charts=False,
    include_binary=False,
)
```

Using options objects:

```python
opts = xir.ParseOptions(sheets=["Data"], profile="fast")
ir = xir.parse_with_options("large.xlsx", opts)
```

## Rebuild and diff

```python
xir.rebuild(ir, "rebuilt.xlsx")
xir.rebuild(ir, "selected.xlsx", sheets="经营驾驶舱")

diff = xir.diff("report.xlsx", "rebuilt.xlsx")
```

Compare two IR objects:

```python
result = xir.compare_ir(ir_a, ir_b, mode="semantic")      # full | semantic | structural
```

## Streaming human-style edits

Streaming edits scan cells directly and do not require a full IR parse:

```python
xir.stream_edit(
    "report.xlsx",
    "edited.xlsx",
    match="总计",
    value="合计",
)
```

Preview an offset edit:

```python
preview = xir.stream_edit(
    "report.xlsx",
    "ignored.xlsx",
    match="业务线",
    value="收入本月",
    options=xir.StreamEditOptions(offset_row=1, offset_col=2, preview=True),
)
```

Update all matches:

```python
xir.stream_edit(
    "report.xlsx",
    "edited.xlsx",
    match="云业务",
    value="云事业部",
    options=xir.StreamEditOptions(update_all=True),
)
```

## Multi-level headers

Horizontal headers locate a column from multi-row headers, then select a row. If `row` and `row_match` are omitted, the first data row after the header rows is used.

```python
xir.header_edit(
    "tests/fixtures/multi_header_dates.xlsx",
    "edited.xlsx",
    headers=["2026", "5", "8"],
    value=999,
    options=xir.HeaderEditOptions(sheet="日期表", row_match="门店A"),
)
```

Regex and wildcard header matching:

```python
xir.header_edit(
    "dates.xlsx",
    "edited.xlsx",
    headers=["202[0-9]", "5", "[78]"],
    value=999,
    options=xir.HeaderEditOptions(match_mode="regex", preview=True),
)

xir.header_edit(
    "dates.xlsx",
    "edited.xlsx",
    headers=["202?", "*", "8"],
    value=999,
    options=xir.HeaderEditOptions(match_mode="wildcard", header_match_index=1),
)
```

Per-level match dictionaries are also supported:

```python
xir.header_edit(
    "dates.xlsx",
    "edited.xlsx",
    headers=[{"regex": "202[0-9]"}, {"value": "5"}, {"wildcard": "*"}],
    value=999,
    options=xir.HeaderEditOptions(preview=True),
)
```

Vertical headers locate a row from multi-column headers, then select a column. If `col` and `col_match` are omitted, the first data column after the header columns is used.

```python
xir.header_edit(
    "vertical.xlsx",
    "edited.xlsx",
    headers=["收入", "线下"],
    value=999,
    options=xir.HeaderEditOptions(
        sheet="纵向表",
        orientation="vertical",
        header_start_col="A",
        header_end_col="B",
        min_row=2,
        col_match="Q2",
    ),
)
```

Inspect expanded headers:

```python
cols = xir.header_columns("dates.xlsx", sheet="日期表", header_rows=(1, 3))
rows = xir.header_rows("vertical.xlsx", sheet="纵向表", header_cols=("A", "B"), min_row=2)
```

## Semantic patch example

```python
patch = {
    "actions": [
        {"op": "set_cell", "sheet": "经营驾驶舱", "cell": "B2", "value": "新标题"}
    ]
}
plan = xir.apply_patch(ir, patch, dry_run=True)
new_ir = xir.apply_patch(ir, patch)
```

## Metadata and semantic tables

The package distinguishes Excel native tables from semantic report tables. Complex merged/multi-level headers are preserved as semantic tables instead of being forced into unsafe native Excel table objects. Semantic metadata can be stored in a veryHidden sheet with checksum verification through the CLI metadata commands.

## CLI examples

```bash
excel-ir doctor
excel-ir engines
excel-ir parse report.xlsx out.ir.json
excel-ir parse report.xlsx selected.ir.json --sheet 经营驾驶舱 --sheet 明细
excel-ir parse large.xlsx fast.ir.json --sheet Data --fast
excel-ir parse large.xlsx lean.ir.json --read-only --no-formula-cache --no-extra --no-binary
excel-ir rebuild out.ir.json rebuilt.xlsx
excel-ir rebuild out.ir.json selected.xlsx --sheet 经营驾驶舱
excel-ir diff report.xlsx rebuilt.xlsx diff.json
excel-ir compare-ir a.ir.json b.ir.json diff.json --semantic-only
excel-ir inspect report.xlsx --out inspect.json
excel-ir anonymize private.xlsx public.xlsx
```

Streaming edit CLI:

```bash
excel-ir stream-edit report.xlsx edited.xlsx --match 总计 --value 合计
excel-ir stream-edit report.xlsx edited.xlsx --match 云业务 --value 云事业部 --all
excel-ir stream-edit report.xlsx ignored.xlsx --match 业务线 --value 收入本月 --offset-row 1 --offset-col 2 --preview
```

Multi-header CLI:

```bash
excel-ir header-edit dates.xlsx edited.xlsx --sheet 日期表 --header-rows 1:3 --headers 2026/5/8 --row-match 门店A --value 999 --as-number
excel-ir header-edit dates.xlsx ignored.xlsx --headers '202[0-9]/5/[78]' --match-mode regex --value 999 --preview
excel-ir header-edit dates.xlsx edited.xlsx --headers '202?/*/8' --match-mode wildcard --value 999 --as-number
excel-ir header-edit vertical.xlsx edited.xlsx --orientation vertical --header-cols A:B --min-row 2 --headers 收入/线下 --col-match Q2 --value 999 --as-number
```

Metadata CLI:

```bash
excel-ir metadata status report.xlsx
excel-ir metadata extract --from-xlsx report.xlsx metadata.json
excel-ir metadata verify metadata.json
excel-ir metadata repair repaired.xlsx --from-xlsx report.xlsx
excel-ir metadata strip stripped.xlsx --from-xlsx report.xlsx
```

Corpus and reports:

```bash
excel-ir corpus list
excel-ir corpus run --config tests/fixtures/corpus_config.json
excel-ir report diff.json report.html
excel-ir audit tx_log.json audit.html
```

## Project layout

```text
src/excel_ir_mvp/
  __init__.py       small public facade exports
  api.py            public Python API
  types.py          option dataclasses
  backends.py       backend registry: openpyxl / wolfxl detection
  excel_ir.py       core cells/styles/layout/formula IR and streaming/header edits
  excel_ir_plus.py  extended OOXML objects, semantic metadata, rebuild/diff helpers
  ir_patch.py       semantic patch engine
  excel_ir_cli.py   CLI adapter
```

## Development checks

```bash
python3 -m unittest -v tests.test_excel_ir_mvp tests.test_patch_ops tests.test_native_tables tests.test_metadata
python3 ci_check.py
python3 -m build --sdist --wheel
python3 -m twine check dist/*
```

## Documentation

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

## License

MIT.
