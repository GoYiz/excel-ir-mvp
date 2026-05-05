# Contributing

Thanks for your interest in Excel IR MVP.

## Development setup

```bash
python3 -m pip install -e . --break-system-packages
python3 -m unittest -v tests.test_excel_ir_mvp tests.test_patch_ops tests.test_native_tables tests.test_metadata
python3 ci_check.py
```

## Release checklist

1. Update version in `pyproject.toml` and `src/excel_ir_mvp/__init__.py`.
2. Run source CI:

```bash
python3 ci_check.py
```

3. Build from a clean temp copy if your filesystem has symlink cleanup issues.
4. Run `twine check` on wheel and sdist.
5. Install the wheel and run:

```bash
python3 ci_check.py --installed
```

## Design principles

- Preserve a reversible physical Excel IR where practical.
- Keep complex human-report tables as semantic tables when native OOXML Table constraints are unsafe.
- Make semantic changes action-based and auditable.
- Keep metadata checksummed and recoverable through XLSX-only flows.
