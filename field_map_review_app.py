from __future__ import annotations

import html
import json
import sys
from pathlib import Path


def esc(x):
    return html.escape(str(x))


def main():
    if len(sys.argv) < 3:
        print('usage: python3 field_map_review_app.py ir.json output.html', file=sys.stderr)
        raise SystemExit(2)
    ir_path = Path(sys.argv[1])
    ir = json.loads(ir_path.read_text(encoding='utf-8'))
    blocks = []
    for s in ir.get('workbook', {}).get('sheets', []):
        for t in s.get('extra', {}).get('tables', []):
            table_name = t.get('name') or t.get('displayName')
            cand = t.get('ir', {}).get('field_map_candidates', {})
            rows = ''.join(f'<tr><td><input class="field" value="{esc(k)}"></td><td><input class="col" value="{esc(v)}"></td></tr>' for k, v in cand.items())
            blocks.append(f'<section data-sheet="{esc(s.get("name"))}" data-table="{esc(table_name)}"><h2>{esc(s.get("name"))} / {esc(table_name)}</h2><table><tr><th>Field</th><th>Column</th></tr>{rows}</table></section>')
    doc = f'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>Field Map Review App</title><style>body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;margin:24px}}table{{border-collapse:collapse}}td,th{{border:1px solid #ddd;padding:6px}}input{{width:160px}}textarea{{width:100%;height:240px}}</style></head><body>
<h1>Field Map Review App</h1><p>编辑字段/列后点击生成 patch JSON。</p>{''.join(blocks)}
<button onclick="addRow()">新增映射行</button>
<button onclick="gen()">生成 confirm_field_map patch</button>
<button onclick="copyPatch()">复制 patch JSON</button>
<button onclick="downloadPatch()">下载 patch.json</button><textarea id="out"></textarea>
<script>
function gen(){{
 const actions=[];
 document.querySelectorAll('section').forEach(sec=>{{
   const fmap={{}};
   sec.querySelectorAll('tr').forEach((tr,i)=>{{if(i===0)return; const f=tr.querySelector('.field').value.trim(); const c=tr.querySelector('.col').value.trim(); if(f&&c)fmap[f]=c;}});
   actions.push({{op:'confirm_field_map', sheet:sec.dataset.sheet, table:sec.dataset.table, header_rows:2, field_map:fmap}});
 }});
 document.getElementById('out').value=JSON.stringify({{name:'field_map_confirmation', actions}}, null, 2);
}}
function addRow(){{ const tbody=document.querySelector('#tbl tbody'); const tr=document.createElement('tr'); tr.innerHTML='<td><input value=""></td><td><input value="A"></td><td>manual</td>'; tbody.appendChild(tr); }}
function copyPatch(){{ if(!document.getElementById('out').value) gen(); navigator.clipboard?.writeText(document.getElementById('out').value); }}
function downloadPatch(){{
 if(!document.getElementById('out').value) gen();
 const blob=new Blob([document.getElementById('out').value], {{type:'application/json'}});
 const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='confirm_field_map.patch.json'; a.click(); URL.revokeObjectURL(a.href);
}}
</script></body></html>'''
    Path(sys.argv[2]).write_text(doc, encoding='utf-8')
    print(json.dumps({'ok': True, 'output': sys.argv[2]}, ensure_ascii=False))

if __name__ == '__main__':
    main()
