from __future__ import annotations

import argparse
import html
import json
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def run(cmd):
    return subprocess.run(cmd, cwd=ROOT, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)


def load_config(path: str | None):
    default = ROOT / 'tests' / 'fixtures' / 'corpus_config.json'
    if path:
        return json.loads((ROOT / path).read_text(encoding='utf-8'))
    if default.exists():
        return json.loads(default.read_text(encoding='utf-8'))
    return {
        'output_dir': 'corpus_results',
        'samples': [
            {'name': 'base_complex', 'category': 'synthetic_complex', 'xlsx': 'tests/fixtures/complex_report.xlsx', 'patch': 'tests/fixtures/v08_patch.json'},
            {'name': 'v08_patched_roundtrip', 'category': 'metadata_roundtrip', 'xlsx': 'tests/fixtures/v08_patched_report.xlsx'},
        ],
    }


def list_samples(config: dict):
    samples = []
    for sample in config.get('samples', []):
        samples.append({
            'name': sample.get('name'),
            'category': sample.get('category', 'uncategorized'),
            'xlsx': sample.get('xlsx'),
            'patch': sample.get('patch'),
        })
    cats = {}
    for s in samples:
        cats[s['category']] = cats.get(s['category'], 0) + 1
    return {'count': len(samples), 'categories': cats, 'samples': samples}


def _category_summary(results):
    cats = {}
    for r in results:
        c = r.get('category') or 'uncategorized'
        item = cats.setdefault(c, {'count': 0, 'ok': 0, 'failed': 0, 'diff_count_total': 0})
        item['count'] += 1
        ok = bool(r.get('parse_ok') and r.get('rebuild_ok') and r.get('diff_count') == 0)
        if ok:
            item['ok'] += 1
        else:
            item['failed'] += 1
        item['diff_count_total'] += int(r.get('diff_count') or 0)
    return cats


def run_corpus(config: dict):
    out = ROOT / config.get('output_dir', 'corpus_results')
    out.mkdir(exist_ok=True)
    results = []
    for sample in config.get('samples', []):
        name = sample['name']
        xlsx = sample['xlsx']
        category = sample.get('category', 'uncategorized')
        ir = out / f'{name}.ir.json'
        rebuilt = out / f'{name}.rebuilt.xlsx'
        diff = out / f'{name}.diff.json'
        metadata = out / f'{name}.metadata.json'
        p1 = run(['python3', 'excel_ir_cli.py', 'parse', xlsx, str(ir)])
        pmeta = run(['python3', 'excel_ir_cli.py', 'metadata', 'export', str(ir), str(metadata)]) if p1.returncode == 0 else None
        p2 = run(['python3', 'excel_ir_cli.py', 'rebuild', str(ir), str(rebuilt)]) if p1.returncode == 0 else None
        pverify = run(['python3', 'excel_ir_cli.py', 'metadata', 'verify', '--from-xlsx', str(rebuilt)]) if p2 and p2.returncode == 0 else None
        p3 = run(['python3', 'excel_ir_cli.py', 'diff', xlsx, str(rebuilt), str(diff)]) if p2 and p2.returncode == 0 else None
        diff_count = json.loads(diff.read_text()).get('diff_count') if diff.exists() else None
        item = {
            'name': name,
            'category': category,
            'parse_ok': p1.returncode == 0,
            'metadata_export_ok': bool(pmeta and pmeta.returncode == 0),
            'metadata_verify_xlsx_ok': bool(pverify and pverify.returncode == 0),
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
    ok = all(r.get('diff_count') == 0 and r.get('parse_ok') and r.get('rebuild_ok') for r in results)
    summary = {'ok': ok, 'categories': _category_summary(results), 'results': results}
    (out / 'summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    return summary


def render_summary_html(summary: dict, title: str = 'Excel IR Corpus Report') -> str:
    rows = []
    for r in summary.get('results', []):
        ok = bool(r.get('parse_ok') and r.get('rebuild_ok') and r.get('diff_count') == 0)
        rows.append('<tr>' + ''.join(f'<td>{html.escape(str(v))}</td>' for v in [
            r.get('name'), r.get('category'), ok, r.get('diff_count'), r.get('metadata_tables'), r.get('metadata_verify_xlsx_ok')
        ]) + '</tr>')
    cats = []
    for name, c in sorted((summary.get('categories') or {}).items()):
        cats.append(f"<li><b>{html.escape(name)}</b>: {c.get('ok')}/{c.get('count')} ok, diff_total={c.get('diff_count_total')}</li>")
    return f"""<!doctype html><meta charset='utf-8'><title>{html.escape(title)}</title>
<style>body{{font-family:-apple-system,BlinkMacSystemFont,sans-serif;margin:2rem}}table{{border-collapse:collapse}}td,th{{border:1px solid #ddd;padding:.35rem .5rem}}th{{background:#f5f5f5}}</style>
<h1>{html.escape(title)}</h1><p>ok: <b>{summary.get('ok')}</b></p><h2>Categories</h2><ul>{''.join(cats)}</ul>
<h2>Samples</h2><table><thead><tr><th>Name</th><th>Category</th><th>OK</th><th>Diffs</th><th>Metadata tables</th><th>XLSX metadata verify</th></tr></thead><tbody>{''.join(rows)}</tbody></table>
"""


def write_report(summary: dict, path: str, title: str = 'Excel IR Corpus Report') -> None:
    Path(path).write_text(render_summary_html(summary, title), encoding='utf-8')


def main():
    ap = argparse.ArgumentParser(description='Run Excel IR corpus tests')
    sub = ap.add_subparsers(dest='cmd')
    p = sub.add_parser('run')
    p.add_argument('--config', help='corpus config JSON, relative to project root')
    p = sub.add_parser('list')
    p.add_argument('--config')
    p = sub.add_parser('report')
    p.add_argument('summary_json')
    p.add_argument('html')
    p.add_argument('--title', default='Excel IR Corpus Report')
    ap.add_argument('--config', help='corpus config JSON, relative to project root')
    args = ap.parse_args()
    cmd = args.cmd or 'run'
    if cmd == 'list':
        print(json.dumps(list_samples(load_config(args.config)), ensure_ascii=False, indent=2))
        return
    if cmd == 'report':
        summary = json.loads(Path(args.summary_json).read_text(encoding='utf-8'))
        write_report(summary, args.html, args.title)
        print(json.dumps({'ok': True, 'output': args.html}, ensure_ascii=False))
        return
    summary = run_corpus(load_config(args.config))
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if not summary['ok']:
        raise SystemExit(1)


if __name__ == '__main__':
    main()
