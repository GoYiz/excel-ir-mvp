from __future__ import annotations

import json
import subprocess
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DEFAULT_MAX_SECONDS = 120.0


def run(cmd):
    t0 = time.perf_counter()
    p = subprocess.run(cmd, cwd=ROOT, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    return {'cmd': cmd, 'returncode': p.returncode, 'seconds': round(time.perf_counter() - t0, 4)}


def run_bench(max_total_seconds: float = DEFAULT_MAX_SECONDS):
    results = []
    results.append(run(['python3', 'excel_ir_cli.py', 'parse', 'tests/fixtures/complex_report.xlsx', 'bench_ir.json']))
    results.append(run(['python3', 'excel_ir_cli.py', 'rebuild', 'bench_ir.json', 'bench_rebuilt.xlsx']))
    results.append(run(['python3', 'excel_ir_cli.py', 'diff', 'tests/fixtures/complex_report.xlsx', 'bench_rebuilt.xlsx', 'bench_diff.json']))
    results.append(run(['python3', 'excel_ir_cli.py', 'patch', 'bench_ir.json', 'tests/fixtures/v08_patch.json', 'bench_patched_ir.json', '--log', 'bench_tx.json']))
    total_seconds = round(sum(r['seconds'] for r in results), 4)
    ok = all(r['returncode'] == 0 for r in results) and total_seconds <= max_total_seconds
    return {'ok': ok, 'max_total_seconds': max_total_seconds, 'total_seconds': total_seconds, 'results': results}


def main():
    summary = run_bench()
    (ROOT / 'bench_results.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if not summary['ok']:
        raise SystemExit(1)

if __name__ == '__main__':
    main()
