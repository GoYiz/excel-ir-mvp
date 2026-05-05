# Excel IR MVP Release Checklist

## Before release

- [ ] Update `pyproject.toml` version.
- [ ] Update `src/excel_ir_mvp/__init__.py` `__version__`.
- [ ] Run source CI: `python3 ci_check.py`.
- [ ] Run installed CLI CI: `python3 ci_check.py --installed` after installing package.
- [ ] Run tests: `python3 -m unittest -v tests/test_excel_ir_mvp.py`.
- [ ] Build wheel and sdist from clean source.
- [ ] Run `twine check dist/*`.
- [ ] Install wheel in a clean environment and smoke test `excel-ir doctor`.
- [ ] Verify package data: `ir.schema.json`, `patch.schema.json`, `corpus_config.json`.
- [ ] Update release notes.
- [ ] Update SHA256SUMS.

## Build commands

```bash
python3 -m build --sdist --wheel
python3 -m twine check dist/*
python3 -m pip install --force-reinstall dist/*.whl
excel-ir doctor
```
