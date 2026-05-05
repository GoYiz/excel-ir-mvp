from __future__ import annotations

import argparse
import json
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def run(cmd):
    return subprocess.run(cmd, cwd=ROOT, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)


def load_config(path: str | None):
    if path:
        return json.loads((ROOT / path).read_text(encoding='utf-8'))
    default = ROOT / 'corpus_config.json'
    if default.exists():
        return json.loads(default.read_text(encoding='utf-8'))
    return {
        'output_dir': 'corpus_results',
        'samples': [
            {'name': 'base_complex', 'category': 'synthetic_complex', 'xlsx': 'complex_report.xlsx', 'patch': 'v08_patch.json'},
            {'name': 'v08_patched_roundtrip', 'category': 'metadata_roundtrip', 'xlsx': 'v08_patched_report.xlsx'},
        ],
    }


def _category_summary(results):
    cats = {}
    for r in results:
        c = r.get('category') or 'uncategorized'
        item = cats.setdefault(c, {'count': 0, 'ok': 0, 'failed': 0, 'diff_count_total': 0})
        item['count'] += 1
        ok = bool(r.get('parse_ok') and r.get('rebuild_ok') and r.get('diff_count') == 0)
        item['ok'] += 1 if ok else 0
        item['failed'] += 0 if ok else 1
        item['diff_count_total'] += int(r.get('diff_count') or 0)
    return cats


def run_corpus(config: dict):
    out = ROOT / config.get('output_dir', 'corpus_results')
    out.mkdir(exist_ok=True)
    results = []
    for sample in config.get('samples', []):
        name = sample['name']
        category = sample.get('category', 'uncategorized')
        ir = out / f'{name}.ir.json'
        rebuilt = out / f'{name}.rebuilt.xlsx'
        diff = out / f'{name}.diff.json'
        metadata = out / f'{name}.metadata.json'
        p1 = run(['python3', 'excel_ir_plus.py', 'parse', sample['xlsx'], str(ir)])
        pmeta = run(['python3', 'excel_ir_cli.py', 'metadata', 'export', str(ir), str(metadata)]) if p1.returncode == 0 else None
        p2 = run(['python3', 'excel_ir_plus.py', 'rebuild', str(ir), str(rebuilt)]) if p1.returncode == 0 else None
        p3 = run(['python3', 'excel_ir_plus.py', 'diff', sample['xlsx'], str(rebuilt), str(diff)]) if p2 and p2.returncode == 0 else None
        diff_count = json.loads(diff.read_text()).get('diff_count') if diff.exists() else None
        item = {
            'name': name,
            'category': category,
            'parse_ok': p1.returncode == 0,
            'metadata_export_ok': bool(pmeta and pmeta.returncode == 0),
            'rebuild_ok': bool(p2 and p2.returncode == 0),
            'diff_ok': bool(p3 and p3.returncode == 0),
            'diff_count': diff_count,
        }
        if metadata.exists():
            try:
                item['metadata_tables'] = sum(len(s.get('tables', [])) for s in json.loads(metadata.read_text(encoding='utf-8')).get('sheets', []))
            except Exception:
                item['metadata_tables'] = None
        if sample.get('patch') and item['parse_ok']:
            patched_ir = out / f'{name}.patched.ir.json'
            tx = out / f'{name}.tx.json'
            p4 = run(['python3', 'ir_patch.py', str(ir), sample['patch'], str(patched_ir), '--log', str(tx)])
            item['patch_ok'] = p4.returncode == 0
            if tx.exists():
                item['impact'] = json.loads(tx.read_text()).get('impact')
        results.append(item)
    summary = {'ok': all(r.get('diff_count') == 0 and r.get('parse_ok') and r.get('rebuild_ok') for r in results), 'categories': _category_summary(results), 'results': results}
    (out / 'summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    return summary


def main():
    ap = argparse.ArgumentParser(description='Run Excel IR corpus tests')
    ap.add_argument('--config', help='corpus config JSON, relative to project root')
    args = ap.parse_args()
    summary = run_corpus(load_config(args.config))
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if not summary['ok']:
        raise SystemExit(1)


if __name__ == '__main__':
    main()
