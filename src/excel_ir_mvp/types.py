from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Sequence, Union

SheetSelector = Optional[Union[str, Sequence[str]]]


@dataclass(frozen=True)
class ParseOptions:
    """Options for parsing an XLSX workbook into IR."""

    engine: str = "openpyxl"
    sheets: SheetSelector = None
    include_empty_styled: bool = True
    infer_logic: bool = True


@dataclass(frozen=True)
class RebuildOptions:
    """Options for rebuilding an XLSX workbook from IR."""

    engine: str = "openpyxl"
    sheets: SheetSelector = None


@dataclass(frozen=True)
class StreamEditOptions:
    """Options for streaming first-match or all-match edits."""

    engine: str = "openpyxl"
    sheet: Optional[str] = None
    start: str = "left"
    contains: bool = False
    case_sensitive: bool = False
    max_cells: Optional[int] = None
    offset_row: int = 0
    offset_col: int = 0
    preview: bool = False
    update_all: bool = False


@dataclass(frozen=True)
class HeaderEditOptions:
    """Options for edits addressed by multi-level header paths."""

    engine: str = "openpyxl"
    sheet: Optional[str] = None
    header_start_row: int = 1
    header_end_row: int = 3
    row: Optional[int] = None
    row_match: object = None
    row_match_col: Union[int, str] = 1
    data_start_row: Optional[int] = None
    min_col: int = 1
    max_col: Optional[int] = None
    contains: bool = False
    case_sensitive: bool = False
    header_match_index: int = 1
    preview: bool = False
