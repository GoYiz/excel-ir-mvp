import json
ir=json.load(open('/var/minis/workspace/excel_ir_mvp/v04_patched_ir.json'))
s=ir['workbook']['sheets'][0]
print(s.get('logical',{}).get('patch_stats'))
