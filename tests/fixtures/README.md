# Fixture Guide

Fixtures are intentionally small and synthetic unless explicitly marked otherwise.

## Current layout

```text
tests/fixtures/
  complex_report.xlsx          # synthetic complex human-report workbook
  complex_ir_v07.json          # stable parsed IR fixture from v0.7 lineage
  v08_patch.json               # semantic patch with confirm_field_map + row ops
  v08_patched_report.xlsx      # patched workbook round-trip fixture
  corpus_config.json           # category-aware corpus configuration
```

## Planned layout

```text
tests/fixtures/synthetic/      # generated fixtures safe to redistribute
tests/fixtures/metadata/       # metadata corruption / round-trip fixtures
tests/fixtures/real_world/     # user-provided anonymized samples, not committed by default
```

## Corpus categories

Current categories:

- `synthetic_complex`
- `metadata_roundtrip`

Planned categories:

- `native_table`
- `semantic_table`
- `real_world`

## Privacy rule

Do not commit real business workbooks unless they are anonymized and explicitly approved for redistribution.
