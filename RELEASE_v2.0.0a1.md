# Release v2.0.0a1 - Explicit Table Kinds

This alpha tightens the semantic-table work from `2.0.0a0` by making table intent explicit in the IR.

## Highlights

- Added explicit `table_kind` metadata:
  - `native` for safe single-row Excel native Tables.
  - `semantic` for human-report tables that should remain semantic IR only.
- Kept backward compatibility with `native_table_supported` / `native_table_skip_reason`.
- Added native-table regression tests to prove simple Excel Tables still round-trip as native OOXML Tables.
- Added semantic override regression: `table_kind: semantic` is enough to suppress native table rebuild even if the legacy flag is absent.
- Expanded CI to include `tests.test_native_tables`.

## Table-kind semantics

`excel_table_native_status(ws, ref)` now returns:

```json
{
  "native_table_supported": true,
  "table_kind": "native"
}
```

for a safe native table, or:

```json
{
  "native_table_supported": false,
  "native_table_skip_reason": "merged_cells_intersect_table",
  "table_kind": "semantic"
}
```

for complex human-report tables with merged/multi-level headers.

`parse_tables` persists the metadata into `sheet.extra.tables[*]`.

`apply_tables` now skips native OOXML table rebuild when either condition is true:

```python
item.get("table_kind") == "semantic"
item.get("native_table_supported") is False
```

This keeps the v2.0a0 warning fix while giving downstream systems a stable, explicit knob.

## Tests

```bash
python3 -m unittest -v tests.test_excel_ir_mvp tests.test_patch_ops tests.test_native_tables
```

Result:

```text
Ran 19 tests ... OK
```

## CI

```bash
python3 ci_check.py
python3 ci_check.py --installed
```

Source CI passed with coverage gate:

```bash
python3 -m coverage report --show-missing --fail-under=60
```

Observed coverage:

```text
TOTAL 2539 statements, 900 missing, 65% coverage
```

Installed CI passed after installing the built wheel.

## Build

```bash
python3 -m build --sdist --wheel
python3 -m twine check dist/*
```

Built artifacts:

- `dist/excel_ir_mvp-2.0.0a1-py3-none-any.whl`
- `dist/excel_ir_mvp-2.0.0a1.tar.gz`

Twine check passed for both artifacts.

## SHA256

```text
b47b27e2854728f26385e77077ef929ee64e7144e864b474e982f4bef29dfce0  dist/excel_ir_mvp-2.0.0a1-py3-none-any.whl
dec84a4262609b12da153ffc6295efd8c264701681eeeecf31c7e8a880e409bb  dist/excel_ir_mvp-2.0.0a1.tar.gz
```

## Notes

The v2.0 alpha series now has two table paths under test:

1. Complex human report table: `table_kind=semantic`, warning suppressed, structural diff remains 0.
2. Simple one-row Excel table: `table_kind=native`, rebuilt as native table, structural diff remains 0.
