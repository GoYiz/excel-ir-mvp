# Release v2.0.0a5 - GitHub Open Source Bootstrap

This alpha is a repository hygiene release after publishing the project to GitHub.

Repository:

- <https://github.com/GoYiz/excel-ir-mvp>

## Highlights

- Created public GitHub repository `GoYiz/excel-ir-mvp`.
- Pushed `main` and tag `v2.0.0a4`.
- Added repository URL to README.
- Added fixture guide at `tests/fixtures/README.md`.
- Removed obsolete root scratch scripts/tests that were superseded by package modules and `tests/`:
  - `test_formula_utils.py`
  - `test_img.py`
  - `test_img2.py`
  - `test_img3.py`
  - `check_v04.py`
  - `patch_ir_demo.py`
  - root `__init__.py`
- Bumped version to `2.0.0a5`.

## Validation

Run:

```bash
python3 -m unittest -v \
  tests.test_excel_ir_mvp \
  tests.test_patch_ops \
  tests.test_native_tables \
  tests.test_metadata
python3 ci_check.py
python3 ci_check.py --installed
```

## Notes

This release does not change core IR semantics. It prepares the repository for external review and ongoing open-source iteration.
