from formula_utils import shift_formula_references, extract_references
cases = [
    "=SUM(C7:C12)+\"A1\"+'经营驾驶舱'!$D$8",
    "='Sheet 1'!A10+B$2+$C3",
]
for f in cases:
    print(f)
    print(shift_formula_references(f, row_delta=1, col_delta=1, row_at=10, col_at=3))
    print(extract_references(f))
