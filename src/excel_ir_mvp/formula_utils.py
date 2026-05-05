from __future__ import annotations

import re
from openpyxl.utils import get_column_letter, column_index_from_string

# Optional sheet prefix: Sheet1!A1, 'Sheet 1'!$A$1:$B$2
REF_RE = re.compile(
    r"(?P<sheet>(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!)?"
    r"(?P<cabs>\$?)(?P<col>[A-Z]{1,3})(?P<rabs>\$?)(?P<row>\d+)"
    r"(?::(?P<cabs2>\$?)(?P<col2>[A-Z]{1,3})(?P<rabs2>\$?)(?P<row2>\d+))?"
)


def mask_string_literals(formula: str) -> str:
    out = []
    in_str = False
    i = 0
    while i < len(formula):
        ch = formula[i]
        if ch == '"':
            out.append(' ')
            if in_str and i + 1 < len(formula) and formula[i + 1] == '"':
                out.append(' ')
                i += 2
                continue
            in_str = not in_str
        else:
            out.append(' ' if in_str else ch)
        i += 1
    return ''.join(out)


def extract_references(formula: str) -> list[dict]:
    if not isinstance(formula, str):
        return []
    text = mask_string_literals(formula)
    refs = []
    for m in REF_RE.finditer(text):
        # Avoid matching parts of names like ABC123Foo by checking boundaries.
        start, end = m.span()
        if start > 0 and text[start - 1].isalnum():
            continue
        if end < len(text) and text[end:end+1].isalpha():
            continue
        sheet = m.group('sheet')[:-1] if m.group('sheet') else None
        if sheet and sheet.startswith("'") and sheet.endswith("'"):
            sheet = sheet[1:-1]
        col = m.group('col')
        row = int(m.group('row'))
        col2 = m.group('col2') or col
        row2 = int(m.group('row2') or row)
        refs.append({
            'sheet': sheet,
            'ref': f"{col}{row}" if (col, row) == (col2, row2) else f"{col}{row}:{col2}{row2}",
            'min_col': column_index_from_string(col),
            'min_row': row,
            'max_col': column_index_from_string(col2),
            'max_row': row2,
            'span': [start, end],
            'text': formula[start:end],
        })
    return refs


def _shift_one(m: re.Match, row_delta: int, col_delta: int, row_at: int, col_at: int) -> str:
    sheet = m.group('sheet') or ''
    cabs, col, rabs, row_s = m.group('cabs'), m.group('col'), m.group('rabs'), m.group('row')
    cabs2, col2, rabs2, row2_s = m.group('cabs2'), m.group('col2'), m.group('rabs2'), m.group('row2')

    def shift(cabs_v: str, col_v: str, rabs_v: str, row_v: str):
        c = column_index_from_string(col_v)
        r = int(row_v)
        if not cabs_v and c >= col_at:
            c = max(1, c + col_delta)
        if not rabs_v and r >= row_at:
            r = max(1, r + row_delta)
        return f"{cabs_v}{get_column_letter(c)}{rabs_v}{r}"

    first = shift(cabs, col, rabs, row_s)
    if col2:
        second = shift(cabs2 or '', col2, rabs2 or '', row2_s)
        return f"{sheet}{first}:{second}"
    return f"{sheet}{first}"


def shift_formula_references(formula: str, row_delta: int = 0, col_delta: int = 0, row_at: int = 1, col_at: int = 1) -> str:
    if not isinstance(formula, str) or not formula.startswith('='):
        return formula
    masked = mask_string_literals(formula)
    result = []
    last = 0
    for m in REF_RE.finditer(masked):
        start, end = m.span()
        if start > 0 and masked[start - 1].isalnum():
            continue
        if end < len(masked) and masked[end:end+1].isalpha():
            continue
        result.append(formula[last:start])
        # Re-run the same regex on original slice to preserve quoted sheet text.
        original_match = REF_RE.match(formula[start:end])
        if original_match:
            result.append(_shift_one(original_match, row_delta, col_delta, row_at, col_at))
        else:
            result.append(formula[start:end])
        last = end
    result.append(formula[last:])
    return ''.join(result)


def workbook_formula_dependencies(ir: dict) -> list[dict]:
    deps = []
    for sheet in ir.get('workbook', {}).get('sheets', []):
        sname = sheet.get('name')
        for coord, cell in sheet.get('cells', {}).items():
            v = cell.get('value')
            if isinstance(v, str) and v.startswith('='):
                refs = extract_references(v)
                if refs:
                    deps.append({'sheet': sname, 'cell': coord, 'formula': v, 'refs': refs})
    return deps
