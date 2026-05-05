from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path('/var/minis/workspace/excel_ir_mvp')


def run(cmd):
    p = subprocess.run(cmd, cwd=ROOT, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if p.returncode != 0:
        print(p.stdout)
        print(p.stderr, file=sys.stderr)
        raise SystemExit(p.returncode)
    return p.stdout

run(['python3', 'excel_ir_cli.py', 'parse', 'tests/fixtures/complex_report.xlsx', 'golden_ir_tmp.json'])
run(['python3', 'excel_ir_cli.py', 'rebuild', 'golden_ir_tmp.json', 'golden_rebuilt_tmp.xlsx'])
run(['python3', 'excel_ir_cli.py', 'diff', 'tests/fixtures/complex_report.xlsx', 'golden_rebuilt_tmp.xlsx', 'golden_diff_tmp.json'])
diff = json.loads((ROOT / 'golden_diff_tmp.json').read_text())
assert diff['diff_count'] == 0, diff
run(['python3', 'ir_patch.py', 'golden_ir_tmp.json', 'tests/fixtures/v08_patch.json', 'golden_patched_tmp.json', '--log', 'golden_tx_tmp.json'])
log = json.loads((ROOT / 'golden_tx_tmp.json').read_text())
assert log['ok'] is True
assert log['impact']['tables_changed']
assert log['actions'][1]['cell_diffs_sample']
print('golden tests passed')

