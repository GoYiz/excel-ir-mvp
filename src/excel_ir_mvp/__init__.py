from __future__ import annotations

try:
    from .excel_ir_plus import (
        parse_workbook_plus, rebuild_workbook_plus, diff_workbooks_plus,
        collect_semantic_metadata, apply_semantic_metadata,
        export_semantic_metadata_from_ir, import_semantic_metadata_to_ir,
        semantic_metadata_diff, verify_metadata_checksum,
        verify_semantic_metadata, verify_semantic_metadata_file,
        extract_semantic_metadata_from_xlsx, verify_semantic_metadata_xlsx,
        repair_semantic_metadata_xlsx, inspect_workbook,
        compare_ir_files, strip_semantic_metadata_xlsx,
        metadata_status_xlsx, anonymize_workbook_xlsx,
        stream_find_cell_xlsx, stream_update_first_match_xlsx,
    )
    from .ir_patch import apply_patch, apply_patch_with_log, dry_run, validate_patch
except ImportError:  # dev fallback when imported from flat source tree
    from excel_ir_plus import (
        parse_workbook_plus, rebuild_workbook_plus, diff_workbooks_plus,
        collect_semantic_metadata, apply_semantic_metadata,
        export_semantic_metadata_from_ir, import_semantic_metadata_to_ir,
        semantic_metadata_diff, verify_metadata_checksum,
        verify_semantic_metadata, verify_semantic_metadata_file,
        extract_semantic_metadata_from_xlsx, verify_semantic_metadata_xlsx,
        repair_semantic_metadata_xlsx, inspect_workbook,
        compare_ir_files, strip_semantic_metadata_xlsx,
        metadata_status_xlsx, anonymize_workbook_xlsx,
        stream_find_cell_xlsx, stream_update_first_match_xlsx,
    )
    from ir_patch import apply_patch, apply_patch_with_log, dry_run, validate_patch

__all__ = [
    "parse_workbook_plus", "rebuild_workbook_plus", "diff_workbooks_plus",
    "collect_semantic_metadata", "apply_semantic_metadata",
    "export_semantic_metadata_from_ir", "import_semantic_metadata_to_ir",
    "semantic_metadata_diff", "verify_metadata_checksum",
    "verify_semantic_metadata", "verify_semantic_metadata_file",
    "extract_semantic_metadata_from_xlsx", "verify_semantic_metadata_xlsx",
    "repair_semantic_metadata_xlsx", "inspect_workbook",
    "compare_ir_files", "strip_semantic_metadata_xlsx",
    "metadata_status_xlsx", "anonymize_workbook_xlsx",
    "stream_find_cell_xlsx", "stream_update_first_match_xlsx",
    "apply_patch", "apply_patch_with_log", "dry_run", "validate_patch",
]

__version__ = "2.0.0a12"
