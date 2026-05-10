from __future__ import annotations

import copy
from pathlib import Path
from typing import Any, Dict, Iterable, Optional, Sequence, Union

from . import excel_ir_plus as _xp
from . import ir_patch as _patch
from .types import HeaderEditOptions, ParseOptions, RebuildOptions, SheetSelector, StreamEditOptions

WorkbookIR = Dict[str, Any]
PatchIR = Dict[str, Any]


def _coerce_sheets(sheets: SheetSelector) -> Optional[list[str]]:
    if sheets is None:
        return None
    if isinstance(sheets, str):
        return [sheets]
    return [str(s) for s in sheets if str(s)]


def parse(path: str | Path, *, sheets: SheetSelector = None, engine: str = "openpyxl", include_empty_styled: bool = True, infer_logic: bool = True, include_formula_cache: bool = True, include_extra: bool = True, include_images: bool = True, include_charts: bool = True, include_binary: bool = True, read_only: bool = False, sparse: bool = True, profile: str = "full") -> WorkbookIR:
    """Parse an XLSX file into Excel IR.

    This is the recommended public entry point. Use ``sheets`` to limit large
    workbooks to one or more sheets.
    """
    return _xp.parse_workbook_plus(str(path), include_empty_styled=include_empty_styled, infer_logic=infer_logic, engine=engine, sheet_names=_coerce_sheets(sheets), include_formula_cache=include_formula_cache, include_extra=include_extra, include_images=include_images, include_charts=include_charts, include_binary=include_binary, read_only=read_only, sparse=sparse, profile=profile)


def parse_with_options(path: str | Path, options: ParseOptions) -> WorkbookIR:
    return parse(path, sheets=options.sheets, engine=options.engine, include_empty_styled=options.include_empty_styled, infer_logic=options.infer_logic, include_formula_cache=options.include_formula_cache, include_extra=options.include_extra, include_images=options.include_images, include_charts=options.include_charts, include_binary=options.include_binary, read_only=options.read_only, sparse=options.sparse, profile=options.profile)


def rebuild(ir: WorkbookIR, path: str | Path, *, sheets: SheetSelector = None, engine: str = "openpyxl") -> None:
    """Rebuild an XLSX file from Excel IR."""
    _xp.rebuild_workbook_plus(ir, str(path), engine=engine, sheet_names=_coerce_sheets(sheets))


def rebuild_with_options(ir: WorkbookIR, path: str | Path, options: RebuildOptions) -> None:
    rebuild(ir, path, sheets=options.sheets, engine=options.engine)


def diff(a: str | Path, b: str | Path, *, engine: str = "openpyxl") -> Dict[str, Any]:
    """Diff two XLSX workbooks through canonical Excel IR."""
    return _xp.diff_workbooks_plus(str(a), str(b), engine=engine)


def compare_ir(a: WorkbookIR, b: WorkbookIR, *, mode: str = "full") -> Dict[str, Any]:
    """Compare two IR objects. mode: full, semantic, or structural."""
    return _xp.compare_ir(a, b, mode=mode)


def inspect(path: str | Path, *, engine: str = "openpyxl") -> Dict[str, Any]:
    """Return a compact workbook overview."""
    return _xp.inspect_workbook(str(path), engine=engine)


def apply_patch(ir: WorkbookIR, patch: PatchIR, *, dry_run: bool = False) -> WorkbookIR | Dict[str, Any]:
    """Apply or preview semantic patch actions to IR."""
    return _patch.dry_run(ir, patch) if dry_run else _patch.apply_patch(copy.deepcopy(ir), patch)


def stream_edit(src: str | Path, dst: str | Path, *, match: Any, value: Any, options: StreamEditOptions | None = None, **kwargs: Any) -> Dict[str, Any]:
    """Edit cells by streaming scan without parsing a full IR."""
    opts = options or StreamEditOptions()
    params = {
        "sheet": opts.sheet,
        "start": opts.start,
        "contains": opts.contains,
        "case_sensitive": opts.case_sensitive,
        "max_cells": opts.max_cells,
        "offset_row": opts.offset_row,
        "offset_col": opts.offset_col,
        "preview": opts.preview,
        "update_all": opts.update_all,
        "engine": opts.engine,
    }
    params.update(kwargs)
    return _xp.stream_update_first_match_xlsx(str(src), str(dst), match, value, **params)


def header_edit(src: str | Path, dst: str | Path, *, headers: Sequence[Any], value: Any, options: HeaderEditOptions | None = None, **kwargs: Any) -> Dict[str, Any]:
    """Edit a cell addressed by multi-level header path plus row selector."""
    opts = options or HeaderEditOptions()
    params = {
        "sheet": opts.sheet,
        "header_start_row": opts.header_start_row,
        "header_end_row": opts.header_end_row,
        "row": opts.row,
        "row_match": opts.row_match,
        "row_match_col": opts.row_match_col,
        "data_start_row": opts.data_start_row,
        "min_col": opts.min_col,
        "max_col": opts.max_col,
        "contains": opts.contains,
        "case_sensitive": opts.case_sensitive,
        "header_match_index": opts.header_match_index,
        "preview": opts.preview,
        "engine": opts.engine,
    }
    params.update(kwargs)
    return _xp.update_cell_by_multi_header_xlsx(str(src), str(dst), list(headers), value, **params)


def header_columns(path: str | Path, *, sheet: str | None = None, header_rows: tuple[int, int] = (1, 3), engine: str = "openpyxl", min_col: int = 1, max_col: int | None = None) -> Dict[str, Any]:
    """Expand merged/multi-row headers into per-column header paths."""
    return _xp.multi_header_columns_xlsx(str(path), sheet=sheet, header_start_row=header_rows[0], header_end_row=header_rows[1], min_col=min_col, max_col=max_col, engine=engine)


def anonymize(src: str | Path, dst: str | Path, *, rewrite_formulas: bool = False, engine: str = "openpyxl") -> Dict[str, Any]:
    return _xp.anonymize_workbook_xlsx(str(src), str(dst), keep_formulas=not rewrite_formulas, engine=engine)


def engines() -> Dict[str, Any]:
    return {"available": _xp.available_engines(), "engines": _xp.engine_status()}
