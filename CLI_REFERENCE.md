# CLI Reference

Generated for Excel IR MVP v2.0.0a4.

## Basic

```bash
excel-ir doctor
excel-ir parse input.xlsx out.ir.json
excel-ir rebuild out.ir.json rebuilt.xlsx
excel-ir diff input.xlsx rebuilt.xlsx diff.json
```

## Patch

```bash
excel-ir patch out.ir.json patch.json patched.ir.json --plan plan.json --log tx.json
excel-ir patch out.ir.json patch.json --dry-run --plan plan.json
```

## Semantic metadata

Export semantic table metadata (`table_kind`, `field_map`, confirmed mappings) from an IR file:

```bash
excel-ir metadata export out.ir.json semantic_metadata.json
```

Import semantic table metadata into another IR file:

```bash
excel-ir metadata import stripped.ir.json semantic_metadata.json restored.ir.json
```

Diff two metadata payloads:

```bash
excel-ir metadata diff a.semantic.json b.semantic.json metadata_diff.json
```

Verify metadata checksum and shape:

```bash
excel-ir metadata verify semantic_metadata.json
```

When rebuilding XLSX, the same metadata is embedded into a very-hidden sheet named `_excel_ir_metadata` so XLSX-only flows can preserve semantic table intent. v2 metadata includes a SHA-256 checksum.

## Reports

```bash
excel-ir report diff.json report.html --title "Report" --plan plan.json --log tx.json
excel-ir audit tx.json audit.html --title "Audit"
```

## Validation / Corpus / Bench

```bash
excel-ir validate ir out.ir.json
excel-ir validate patch patch.json
excel-ir corpus --config corpus_config.json
excel-ir bench
```

Corpus summary includes category rollups such as `synthetic_complex` and `metadata_roundtrip`.

## Field Map Review

```bash
excel-ir field-map-review out.ir.json field_map_review.html
```

The generated HTML can produce and download a `confirm_field_map.patch.json` file.
