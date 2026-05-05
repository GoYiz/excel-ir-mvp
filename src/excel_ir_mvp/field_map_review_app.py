from __future__ import annotations

import html
import json
import sys
from pathlib import Path
from typing import Sequence


def esc(x):
    return html.escape(str(x))


def main(argv: Sequence[str] | None = None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if len(argv) < 2:
        print('usage: python3 field_map_review_app.py ir.json output.html', file=sys.stderr)
        raise SystemExit(2)
    ir_path = Path(argv[0])
    out_path = Path(argv[1])
    ir = json.loads(ir_path.read_text(encoding='utf-8'))
    blocks = []
    for s in ir.get('workbook', {}).get('sheets', []):
        for t in s.get('extra', {}).get('tables', []):
            table_name = t.get('name') or t.get('displayName')
            cand = t.get('ir', {}).get('field_map_candidates', {})
            rows = ''.join(
                f'<tr><td><input class="field" value="{esc(k)}"></td><td><input class="col" value="{esc(v)}"></td><td>candidate</td></tr>'
                for k, v in cand.items()
            )
            blocks.append(
                f'<section data-sheet="{esc(s.get("name"))}" data-table="{esc(table_name)}">'
                f'<h2>{esc(s.get("name"))} / {esc(table_name)}</h2>'
                f'<button type="button" onclick="addRow(this)">新增映射行</button>'
                f'<table><thead><tr><th>Field</th><th>Column</th><th>Source</th></tr></thead><tbody>{rows}</tbody></table></section>'
            )
    doc = f'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Field Map Review App</title><style>body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;margin:24px}}table{{border-collapse:collapse;margin:12px 0 20px}}td,th{{border:1px solid #ddd;padding:6px}}input{{width:180px}}textarea{{width:100%;height:260px}}button{{margin:4px 6px 4px 0}}</style></head><body>
<h1>Field Map Review App</h1><p>编辑字段/列后点击生成 patch JSON。可新增人工映射行、复制或下载 patch。</p>{''.join(blocks)}
<button onclick="gen()">生成 confirm_field_map patch</button>
<button onclick="copyPatch()">复制 patch JSON</button>
<button onclick="downloadPatch()">下载 patch.json</button><textarea id="out"></textarea>
<script>
function gen(){{
 const actions=[];
 document.querySelectorAll('section').forEach(sec=>{{
   const fmap={{}};
   sec.querySelectorAll('tbody tr').forEach(tr=>{{ const f=tr.querySelector('.field').value.trim(); const c=tr.querySelector('.col').value.trim(); if(f&&c)fmap[f]=c; }});
   actions.push({{op:'confirm_field_map', sheet:sec.dataset.sheet, table:sec.dataset.table, header_rows:2, field_map:fmap}});
 }});
 document.getElementById('out').value=JSON.stringify({{name:'field_map_confirmation', actions}}, null, 2);
}}
function addRow(btn){{ const tbody=btn.closest('section').querySelector('tbody'); const tr=document.createElement('tr'); tr.innerHTML='<td><input class="field" value=""></td><td><input class="col" value="A"></td><td>manual</td>'; tbody.appendChild(tr); }}
function copyPatch(){{ if(!document.getElementById('out').value) gen(); navigator.clipboard?.writeText(document.getElementById('out').value); }}
function downloadPatch(){{
 if(!document.getElementById('out').value) gen();
 const blob=new Blob([document.getElementById('out').value], {{type:'application/json'}});
 const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='confirm_field_map.patch.json'; a.click(); URL.revokeObjectURL(a.href);
}}
</script></body></html>'''
    out_path.write_text(doc, encoding='utf-8')
    print(json.dumps({'ok': True, 'output': str(out_path)}, ensure_ascii=False))


if __name__ == '__main__':
    main()
