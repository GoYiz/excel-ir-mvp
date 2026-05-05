# Real-world Fixture Privacy Guide

Do not commit proprietary or personal workbooks.

If a real-world pattern is needed for a regression test:

1. Recreate it as a synthetic workbook with the same layout behavior.
2. Remove company names, personal names, account numbers, IDs, emails, addresses, and financial values.
3. Replace business-sensitive formulas with equivalent toy formulas.
4. Keep only the minimum sheets/cells needed to reproduce the parser/rebuilder behavior.
5. Prefer generated fixtures under `tools/` so reviewers can inspect how data is made.

Safe fixture categories:

- `synthetic_complex`
- `metadata_roundtrip`
- `native_table`
- `semantic_table`
- future `real_world_anonymized` only after explicit review
