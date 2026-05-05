# CLI Reference

Generated for Excel IR MVP v2.0.0a7.

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

```bash
excel-ir metadata export out.ir.json semantic_metadata.json
excel-ir metadata import stripped.ir.json semantic_metadata.json restored.ir.json
excel-ir metadata extract semantic_metadata.json --from-xlsx workbook.xlsx
excel-ir metadata diff a.semantic.json b.semantic.json metadata_diff.json
excel-ir metadata verify semantic_metadata.json
excel-ir metadata verify --from-xlsx rebuilt.xlsx
```

## Corpus

```bash
excel-ir corpus list --config corpus_config.json
excel-ir corpus run --config corpus_config.json
excel-ir corpus report corpus_results/summary.json corpus_report.html
```

Corpus summary includes category rollups such as `synthetic_complex`, `metadata_roundtrip`, `native_table`, and `semantic_table`.

## Reports

```bash
excel-ir report diff.json report.html --title "Report" --plan plan.json --log tx.json
excel-ir audit tx.json audit.html --title "Audit"
```

## Validation / Bench

```bash
excel-ir validate ir out.ir.json
excel-ir validate patch patch.json
excel-ir bench
```

## Field Map Review

```bash
excel-ir field-map-review out.ir.json field_map_review.html
```

The generated HTML can produce and download a `confirm_field_map.patch.json` file.
