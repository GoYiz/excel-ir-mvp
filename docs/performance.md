# Performance Guide

Large XLSX files can become slow or memory-hungry because Excel files are not simple tables: they can contain a very large used range, repeated styles, images, charts, validations, conditional formatting, formulas and hidden metadata.

## Why parsing can be slow

The main pressure points are:

1. **Workbook object loading**: normal `openpyxl.load_workbook` materializes a large object graph.
2. **Formula cache loading**: reading formulas and cached values can require loading the workbook twice (`data_only=False` and `data_only=True`).
3. **Rectangular scans**: scanning `max_row × max_column` is dangerous when a workbook has far-away styled cells or inflated dimensions.
4. **Style serialization**: converting each cell style into nested dictionaries and JSON keys is expensive when there are many cells.
5. **Binary payloads**: images are base64 encoded in full-fidelity mode, increasing memory and JSON size.
6. **Extended extras**: charts, validations, conditional formatting, tables, metadata and logical inference add work.
7. **Dense logical inference**: building a dense occupancy matrix is not suitable for very large sparse sheets.

## Recommended large-file workflow

Start with selective sheets plus fast profile:

```bash
excel-ir parse large.xlsx selected.ir.json --sheet Data --fast
```

or in Python:

```python
import excel_ir_mvp as xir

ir = xir.parse("large.xlsx", sheets="Data", profile="fast")
```

`--fast` / `profile="fast"` currently uses read-only streaming mode where possible and skips:

- empty styled cells
- formula cache workbook loading
- logical inference
- extended OOXML extras
- images and binary payloads
- charts
- hidden semantic metadata merge

Use full mode only when you need maximum reconstruction fidelity:

```bash
excel-ir parse report.xlsx full.ir.json --profile full
```

## Fine-grained knobs

```bash
excel-ir parse large.xlsx out.ir.json \
  --sheet Data \
  --read-only \
  --no-formula-cache \
  --no-extra \
  --no-images \
  --no-charts \
  --no-binary
```

The parser also uses sparse cell iteration by default to avoid instantiating empty cells across inflated worksheet dimensions. Use `--dense` only for debugging.

## Trade-off

Fast IR is intended for analysis, inspection and targeted workflows. Full IR remains the right choice when exact charts/images/metadata round-trip fidelity matters.
