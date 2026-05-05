from __future__ import annotations

import html
import json
import sys
from pathlib import Path


def esc(x):
    return html.escape(str(x))


def main():
    if len(sys.argv) < 3:
        print('usage: python3 field_map_review.py ir.json output.html', file=sys.stderr)
        raise SystemExit(2)
    ir = json.loads(Path(sys.argv[1]).read_text(encoding='utf-8'))
    blocks = []
    for s in ir.get('workbook', {}).get('sheets', []):
        for t in s.get('extra', {}).get('tables', []):
            cand = t.get('ir', {}).get('field_map_candidates', {})
            rows = ''.join(f'<tr><td>{esc(k)}</td><td>{esc(v)}</td><td><input value="{esc(v)}"></td></tr>' for k, v in cand.items())
            blocks.append(f'<h2>{esc(s.get("name"))} / {esc(t.get("name") or t.get("displayName"))}</h2><p>ref: {esc(t.get("ref"))}</p><table><tr><th>Candidate field</th><th>Column</th><th>Confirm/Override</th></tr>{rows}</table>')
    doc = f'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Field Map Review</title><style>body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;margin:24px}}table{{border-collapse:collapse}}td,th{{border:1px solid #ddd;padding:6px 10px}}input{{width:80px}}</style></head><body><h1>Field Map Review</h1>{''.join(blocks)}</body></html>'''
    Path(sys.argv[2]).write_text(doc, encoding='utf-8')
    print(json.dumps({'ok': True, 'output': sys.argv[2]}, ensure_ascii=False))

if __name__ == '__main__':
    main()
