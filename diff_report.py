from __future__ import annotations

import html
import json
import sys
from pathlib import Path
from typing import Any, Dict, List


def load(path: str) -> Dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def esc(x: Any) -> str:
    if isinstance(x, (dict, list)):
        s = json.dumps(x, ensure_ascii=False, indent=2)
    else:
        s = str(x)
    return html.escape(s)


def render_plan(plan: Dict[str, Any]) -> str:
    rows = []
    for i, a in enumerate(plan.get("actions", []), start=1):
        rows.append(f"""
        <tr>
          <td>{i}</td>
          <td>{esc(a.get('op', ''))}</td>
          <td>{esc(a.get('sheet', ''))}</td>
          <td><pre>{esc({k:v for k,v in a.items() if k not in ('op','sheet')})}</pre></td>
        </tr>""")
    validation = plan.get("validation", [])
    val_html = "".join(f"<li><pre>{esc(v)}</pre></li>" for v in validation) or "<li>No validation errors.</li>"
    return f"""
<h2>Patch Dry-run Plan</h2>
<ul>{val_html}</ul>
<table>
<thead><tr><th>#</th><th>Op</th><th>Sheet</th><th>Preview</th></tr></thead>
<tbody>{''.join(rows)}</tbody>
</table>
"""


def render_tx_log(log: Dict[str, Any]) -> str:
    rows = []
    for a in log.get("actions", []):
        rows.append(f"""
        <tr>
          <td>{esc(a.get('index'))}</td>
          <td>{esc(a.get('op'))}</td>
          <td><pre>{esc(a.get('before', {}))}</pre></td>
          <td><pre>{esc(a.get('after', {}))}</pre></td>
        </tr>""")
    return f"""
<h2>Transaction Apply Log</h2>
<p>ok: <strong>{esc(log.get('ok'))}</strong></p>
<table>
<thead><tr><th>#</th><th>Op</th><th>Before</th><th>After</th></tr></thead>
<tbody>{''.join(rows)}</tbody>
</table>
"""


def render(diff: Dict[str, Any], title: str = "Excel IR Diff Report", plan: Dict[str, Any] | None = None, tx_log: Dict[str, Any] | None = None) -> str:
    rows = []
    for i, d in enumerate(diff.get("diffs", []), start=1):
        rows.append(f"""
        <tr>
          <td>{i}</td>
          <td>{esc(d.get('sheet', ''))}</td>
          <td>{esc(d.get('coord', ''))}</td>
          <td>{esc(d.get('type', ''))}</td>
          <td><pre>{esc(d.get('a', ''))}</pre></td>
          <td><pre>{esc(d.get('b', ''))}</pre></td>
        </tr>""")
    status = "PASS" if diff.get("diff_count") == 0 else "DIFF"
    color = "#107C10" if status == "PASS" else "#C00000"
    plan_html = render_plan(plan) if plan else ""
    tx_html = render_tx_log(tx_log) if tx_log else ""
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>{html.escape(title)}</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; margin: 24px; color: #222; }}
.badge {{ display: inline-block; padding: 6px 12px; border-radius: 999px; background: {color}; color: white; font-weight: 700; }}
.summary {{ margin: 16px 0; padding: 12px 16px; background: #f6f8fa; border: 1px solid #d0d7de; border-radius: 8px; }}
table {{ border-collapse: collapse; width: 100%; table-layout: fixed; }}
th, td {{ border: 1px solid #d0d7de; padding: 8px; vertical-align: top; }}
th {{ background: #f6f8fa; }}
pre {{ white-space: pre-wrap; word-break: break-word; margin: 0; font-size: 12px; }}
</style>
</head>
<body>
<h1>{html.escape(title)}</h1>
<div class="summary">
  <span class="badge">{status}</span>
  <p>diff_count: <strong>{diff.get('diff_count')}</strong></p>
  <p>truncated: <strong>{diff.get('truncated')}</strong></p>
</div>
{plan_html}
{tx_html}
<table>
<thead><tr><th>#</th><th>Sheet</th><th>Coord</th><th>Type</th><th>A</th><th>B</th></tr></thead>
<tbody>{''.join(rows) if rows else '<tr><td colspan="6">No differences.</td></tr>'}</tbody>
</table>
</body>
</html>"""


def main() -> None:
    if len(sys.argv) < 3:
        print("usage: python3 diff_report.py diff.json report.html [title] [plan.json] [tx_log.json]", file=sys.stderr)
        raise SystemExit(2)
    diff = load(sys.argv[1])
    title = sys.argv[3] if len(sys.argv) > 3 else "Excel IR Diff Report"
    plan = load(sys.argv[4]) if len(sys.argv) > 4 else None
    tx_log = load(sys.argv[5]) if len(sys.argv) > 5 else None
    Path(sys.argv[2]).write_text(render(diff, title, plan, tx_log), encoding="utf-8")
    print(json.dumps({"ok": True, "output": sys.argv[2]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
