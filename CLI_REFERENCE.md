# CLI Reference

```bash
excel-ir doctor
excel-ir bench
```

## Workbook IR

```bash
excel-ir inspect workbook.xlsx --out inspect.json
excel-ir parse workbook.xlsx out.ir.json
excel-ir rebuild out.ir.json rebuilt.xlsx
excel-ir diff original.xlsx rebuilt.xlsx diff.json
```

## Patch / audit

```bash
excel-ir patch in.ir.json patch.json out.ir.json --plan plan.json --log tx.json
excel-ir patch in.ir.json patch.json --dry-run --plan plan.json
excel-ir report diff.json report.html --plan plan.json --log tx.json
excel-ir audit tx.json audit.html
excel-ir field-map-review in.ir.json field_map_review.html
```

## Validation

```bash
excel-ir validate ir out.ir.json
excel-ir validate patch patch.json
```

## Semantic metadata

```bash
excel-ir metadata export out.ir.json metadata.json
excel-ir metadata import stripped.ir.json metadata.json restored.ir.json
excel-ir metadata extract metadata.json --from-xlsx workbook.xlsx
excel-ir metadata verify metadata.json
excel-ir metadata verify --from-xlsx workbook.xlsx
excel-ir metadata repair repaired.xlsx --from-xlsx workbook.xlsx
excel-ir metadata diff a.metadata.json b.metadata.json metadata_diff.json
```

## Corpus

```bash
excel-ir corpus list --config tests/fixtures/corpus_config.json
excel-ir corpus run --config tests/fixtures/corpus_config.json
excel-ir corpus report corpus_results/summary.json corpus_results/report.html
```

`corpus run` writes `summary.json` and `report.html` into the configured output directory.
