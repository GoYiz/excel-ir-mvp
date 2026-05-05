from __future__ import annotations

import html
import json
import sys
from pathlib import Path
from typing import Any, Dict


def esc(x: Any) -> str:
    if isinstance(x, (dict, list)):
        x = json.dumps(x, ensure_ascii=False, indent=2)
    return html.escape(str(x))


def render(log: Dict[str, Any], title: str = "Excel IR Audit Report") -> str:
    rows = []
    for a in log.get("actions", []):
        diffs = a.get("cell_diffs_sample", [])[:12]
        diff_html = "".join(f"<li><code>{esc(d.get('sheet'))}!{esc(d.get('coord'))}</code></li>" for d in diffs)
        rows.append(f"""
        <tr>
          <td>{esc(a.get('index'))}</td>
          <td>{esc(a.get('op'))}</td>
          <td><pre>{esc(a.get('impact', {}))}</pre></td>
          <td><ul>{diff_html}</ul></td>
        </tr>""")
    return f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>{esc(title)}</title>
<style>body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;margin:24px}}table{{border-collapse:collapse;width:100%}}td,th{{border:1px solid #d0d7de;padding:8px;vertical-align:top}}th{{background:#f6f8fa}}pre{{white-space:pre-wrap;font-size:12px}}</style>
</head><body>
<h1>{esc(title)}</h1>
<h2>Overall Impact</h2><pre>{esc(log.get('impact', {}))}</pre>
<h2>Impact Graph</h2><pre>{esc(log.get('impact_graph', {}))}</pre>
<h2>Action Attribution</h2>
<table><thead><tr><th>#</th><th>Op</th><th>Impact</th><th>Cell diff sample</th></tr></thead><tbody>{''.join(rows)}</tbody></table>
</body></html>"""


def main() -> None:
    if len(sys.argv) < 3:
        print("usage: python3 audit_report.py tx_log.json output.html [title]", file=sys.stderr)
        raise SystemExit(2)
    log = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    title = sys.argv[3] if len(sys.argv) > 3 else "Excel IR Audit Report"
    Path(sys.argv[2]).write_text(render(log, title), encoding="utf-8")
    print(json.dumps({"ok": True, "output": sys.argv[2]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
