# Backend Engines

Excel IR supports a small backend registry so callers can choose the workbook I/O engine per operation.

Available choices:

- `openpyxl`: default, always used by CI and the fidelity pipeline today.
- `wolfxl`: optional Rust-backed openpyxl-compatible engine. If `wolfxl` is not importable, the CLI returns a clear error instead of silently falling back.
- `auto`: choose `wolfxl` when it is importable, otherwise choose `openpyxl`.

Inspect availability:

```bash
excel-ir engines
excel-ir doctor
```

Use an engine explicitly:

```bash
excel-ir parse workbook.xlsx workbook.ir.json --engine openpyxl
excel-ir inspect workbook.xlsx --engine auto
excel-ir stream-edit workbook.xlsx edited.xlsx --match 总计 --value 合计 --engine openpyxl
```

API:

```python
from excel_ir_mvp import parse_workbook_plus, rebuild_workbook_plus, available_engines, engine_status

print(available_engines())
print(engine_status())
ir = parse_workbook_plus("workbook.xlsx", engine="auto")
rebuild_workbook_plus(ir, "rebuilt.xlsx", engine="openpyxl")
```

Notes:

- The IR records the actual selected engine under `workbook.engine` for parse operations.
- The compatibility layer is intentionally conservative: `openpyxl` remains the reference backend, while `wolfxl` can be enabled in environments where its wheel is available.
- On Alpine/musl a `wolfxl` wheel may be unavailable; source install may require a Rust toolchain. In that case use `--engine openpyxl`.
