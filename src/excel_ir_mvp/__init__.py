from __future__ import annotations

from .api import (
    WorkbookIR,
    PatchIR,
    anonymize,
    apply_patch,
    compare_ir,
    diff,
    engines,
    header_columns,
    header_edit,
    inspect,
    parse,
    parse_with_options,
    rebuild,
    rebuild_with_options,
    stream_edit,
)
from .backends import BackendUnavailableError, available_engines, engine_status, resolve_engine
from .types import HeaderEditOptions, ParseOptions, RebuildOptions, StreamEditOptions

# Backward-compatible advanced API aliases. Prefer excel_ir_mvp.api for new code.
try:
    from .excel_ir_plus import (
        parse_workbook_plus,
        rebuild_workbook_plus,
        diff_workbooks_plus,
        inspect_workbook,
        compare_ir_files,
        metadata_status_xlsx,
        anonymize_workbook_xlsx,
        stream_find_cell_xlsx,
        stream_update_first_match_xlsx,
        multi_header_columns_xlsx,
        locate_cell_by_multi_header_xlsx,
        update_cell_by_multi_header_xlsx,
    )
except ImportError:  # pragma: no cover - flat source fallback
    pass

__version__ = "2.0.0a16"

__all__ = [
    # Stable, concise public API
    "WorkbookIR",
    "PatchIR",
    "parse",
    "parse_with_options",
    "rebuild",
    "rebuild_with_options",
    "diff",
    "compare_ir",
    "inspect",
    "apply_patch",
    "stream_edit",
    "header_edit",
    "header_columns",
    "anonymize",
    "engines",
    "ParseOptions",
    "RebuildOptions",
    "StreamEditOptions",
    "HeaderEditOptions",
    "available_engines",
    "engine_status",
    "resolve_engine",
    "BackendUnavailableError",
    # Compatibility aliases retained for existing users
    "parse_workbook_plus",
    "rebuild_workbook_plus",
    "diff_workbooks_plus",
    "inspect_workbook",
    "compare_ir_files",
    "metadata_status_xlsx",
    "anonymize_workbook_xlsx",
    "stream_find_cell_xlsx",
    "stream_update_first_match_xlsx",
    "multi_header_columns_xlsx",
    "locate_cell_by_multi_header_xlsx",
    "update_cell_by_multi_header_xlsx",
]
