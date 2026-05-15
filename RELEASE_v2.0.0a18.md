# Release v2.0.0a18 - Flexible Multi Headers and Complete README

## Highlights

- Multi-level header matching now supports exact, contains, wildcard/glob and regex modes.
- `header-edit` supports both horizontal and vertical headers.
- `row_match` is optional for horizontal headers; when omitted, the target row defaults to the first data row after the header rows.
- For vertical headers, `col_match` is optional; when omitted, the target column defaults to the first data column after the header columns.
- Added public `header_rows()` facade to inspect vertical multi-column headers.
- README is expanded with package introduction, conceptual design, Python API examples, CLI examples, metadata, performance, streaming and multi-header usage.

## Python examples

```python
import excel_ir_mvp as xir

# Horizontal multi-row headers with regex matching
xir.header_edit(
    "dates.xlsx",
    "edited.xlsx",
    headers=["202[0-9]", "5", "[78]"],
    value=999,
    options=xir.HeaderEditOptions(match_mode="regex", preview=True),
)

# Vertical multi-column headers
xir.header_edit(
    "vertical.xlsx",
    "edited.xlsx",
    headers=["收入", "线下"],
    value=999,
    options=xir.HeaderEditOptions(
        orientation="vertical",
        header_start_col="A",
        header_end_col="B",
        min_row=2,
        col_match="Q2",
    ),
)
```

## CLI examples

```bash
excel-ir header-edit dates.xlsx ignored.xlsx --headers '202[0-9]/5/[78]' --match-mode regex --value 999 --preview
excel-ir header-edit dates.xlsx edited.xlsx --headers '202?/*/8' --match-mode wildcard --value 999 --as-number
excel-ir header-edit vertical.xlsx edited.xlsx --orientation vertical --header-cols A:B --min-row 2 --headers 收入/线下 --col-match Q2 --value 999 --as-number
```

## Validation

- Source CI: 39 tests, coverage gate 70%.
- Installed CI: package installed from wheel and CLI smoke tests pass.
- Build: wheel and sdist pass `twine check`.
