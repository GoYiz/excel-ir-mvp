from __future__ import annotations

import json
from pathlib import Path


def patch_ir(input_json: str, output_json: str) -> None:
    ir = json.loads(Path(input_json).read_text(encoding="utf-8"))
    sheet = ir["workbook"]["sheets"][0]

    # Example 1: modify title and metadata while preserving original layout/style.
    sheet["cells"]["A1"]["value"] = "华东大区销售日报（修订版）"
    sheet["cells"]["A2"]["value"] = "报表日期：2026-05-05"

    # Example 2: modify a body value; formulas remain formulas in IR.
    sheet["cells"]["C7"]["value"] = 188.8

    # Example 3: add a semantic note near logical table metadata.
    sheet.setdefault("logical", {}).setdefault("annotations", []).append({
        "type": "manual_patch",
        "target": "C7",
        "message": "将上海-直营本日销售额调整为 188.8；原样式不变。",
    })

    Path(output_json).write_text(json.dumps(ir, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    patch_ir(
        "/var/minis/workspace/excel_ir_mvp/sample_ir_v4.json",
        "/var/minis/workspace/excel_ir_mvp/patched_ir.json",
    )
    print("created patched_ir.json")
