from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

A1_RE = re.compile(r'^[A-Z]{1,3}[1-9][0-9]*$')
RANGE_RE = re.compile(r'^[A-Z]{1,3}[1-9][0-9]*(?::[A-Z]{1,3}[1-9][0-9]*)?$')


@dataclass
class CellIR:
    row: int
    col: int
    value: Any = None
    data_type: str = 'n'
    style_id: Optional[str] = None

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'CellIR':
        return cls(row=int(d['row']), col=int(d['col']), value=d.get('value'), data_type=d.get('data_type', 'n'), style_id=d.get('style_id'))

    def validate(self, coord: str) -> List[str]:
        errs = []
        if self.row < 1 or self.col < 1:
            errs.append(f'{coord}: row/col must be positive')
        if self.data_type not in {'n', 's', 'str', 'inlineStr', 'f', 'b', 'd', 'e'}:
            errs.append(f'{coord}: unusual data_type {self.data_type!r}')
        return errs


@dataclass
class TableIR:
    name: str
    ref: str
    field_map: Dict[str, str] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'TableIR':
        return cls(name=d.get('name') or d.get('displayName') or '', ref=d.get('ref') or '', field_map=d.get('ir', {}).get('field_map', {}))

    def validate(self) -> List[str]:
        errs = []
        if not self.name:
            errs.append('table: missing name')
        if not RANGE_RE.match(self.ref):
            errs.append(f'table {self.name}: invalid ref {self.ref!r}')
        for k, v in self.field_map.items():
            if not re.match(r'^[A-Z]{1,3}$', str(v)):
                errs.append(f'table {self.name}: field_map {k!r} -> invalid column {v!r}')
        return errs


@dataclass
class SheetIR:
    name: str
    cells: Dict[str, CellIR]
    tables: List[TableIR] = field(default_factory=list)

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'SheetIR':
        cells = {k: CellIR.from_dict(v) for k, v in d.get('cells', {}).items()}
        tables = [TableIR.from_dict(t) for t in d.get('extra', {}).get('tables', [])]
        return cls(name=d['name'], cells=cells, tables=tables)

    def validate(self) -> List[str]:
        errs = []
        if not self.name:
            errs.append('sheet: missing name')
        for coord, cell in self.cells.items():
            if not A1_RE.match(coord):
                errs.append(f'invalid cell coordinate {coord!r}')
            errs.extend(cell.validate(coord))
        for table in self.tables:
            errs.extend(table.validate())
        return errs


@dataclass
class WorkbookIR:
    schema_version: str
    sheets: List[SheetIR]

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'WorkbookIR':
        return cls(schema_version=d.get('schema_version', ''), sheets=[SheetIR.from_dict(s) for s in d.get('workbook', {}).get('sheets', [])])

    def validate(self) -> List[str]:
        errs = []
        if not self.schema_version:
            errs.append('missing schema_version')
        if not self.sheets:
            errs.append('workbook has no sheets')
        names = set()
        for sheet in self.sheets:
            if sheet.name in names:
                errs.append(f'duplicate sheet name {sheet.name!r}')
            names.add(sheet.name)
            errs.extend(sheet.validate())
        return errs


def validate_basic_types(ir: Dict[str, Any]) -> List[Dict[str, str]]:
    errors: List[Dict[str, str]] = []
    try:
        wb = WorkbookIR.from_dict(ir)
        for msg in wb.validate():
            errors.append({'level': 'error', 'message': msg})
    except Exception as e:
        errors.append({'level': 'error', 'message': f'type model validation failed: {e}'})
    return errors
