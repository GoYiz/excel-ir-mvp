from __future__ import annotations

import copy
import sys
import unittest
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / 'src'
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from tests.fixtures_loader import fixture_path, load_json_fixture


class ExcelIRPatchOpTests(unittest.TestCase):
    def test_patch_helpers(self):
        from excel_ir_mvp import ir_patch as ip
        ir = load_json_fixture('complex_ir_v07.json')
        sheet = ip.get_sheet(ir, name='经营驾驶舱')
        self.assertEqual(ip.coord_to_rc('C7'), (7, 3))
        self.assertEqual(ip.rc_to_coord(7, 3), 'C7')
        self.assertEqual(list(ip.iter_range('A1:B1'))[1][0], 'B1')
        self.assertEqual(ip.infer_data_type('=SUM(A1:A2)'), 'f')
        self.assertTrue(ip.compare_value(5, '>', 3))
        self.assertTrue(ip.compare_value('abc', 'contains', 'b'))
        self.assertEqual(ip.render_value_template('row-{row}', 9), 'row-9')
        self.assertEqual(ip.shift_range_ref('A1:A3', row_delta=1, row_at=2), 'A1:A4')
        c = ip.ensure_cell(sheet, 'Z99')
        self.assertEqual((c['row'], c['col']), (99, 26))

    def test_patch_apply_and_formula_tools(self):
        from excel_ir_mvp.formula_utils import extract_references, shift_formula_references, workbook_formula_dependencies
        from excel_ir_mvp.ir_patch import apply_patch_with_log, dry_run, validate_patch
        ir = load_json_fixture('complex_ir_v07.json')
        patch = load_json_fixture('v08_patch.json')
        self.assertFalse(validate_patch(ir, patch))
        plan = dry_run(ir, patch)
        self.assertTrue(plan['ok'])
        out, log = apply_patch_with_log(ir, patch)
        self.assertTrue(log['ok'])
        self.assertIn('impact_graph', log)
        self.assertIn('tables_changed', log['impact'])
        self.assertTrue(out['workbook'].get('patch_history'))
        refs = extract_references('=SUM(C7:C12)+"A1"')
        self.assertEqual(refs[0]['ref'], 'C7:C12')
        self.assertEqual(shift_formula_references('=SUM(C7:C12)', row_delta=1, row_at=10), '=SUM(C7:C13)')
        self.assertGreater(len(workbook_formula_dependencies(ir)), 0)

    def test_table_row_ops(self):
        from excel_ir_mvp.ir_patch import apply_patch_with_log
        ir = load_json_fixture('complex_ir_v07.json')
        patch = {'name': 'row_ops', 'actions': [
            {'op': 'update_rows_where', 'sheet': '经营驾驶舱', 'table': 'KPI_Table', 'header_rows': 2, 'where': {'col': '业务线', 'op': '==', 'value': '硬件'}, 'updates': {'状态/评级': 'A+'}},
            {'op': 'append_table_row', 'sheet': '经营驾驶舱', 'table': 'KPI_Table', 'header_rows': 2, 'values': {'业务线': '新增线', '区域': '北区', '收入/本月': 1}},
            {'op': 'recompute_totals', 'sheet': '经营驾驶舱', 'table': 'KPI_Table', 'columns': ['收入/本月']},
        ]}
        out, log = apply_patch_with_log(ir, patch)
        self.assertTrue(log['ok'])
        sheet = next(s for s in out['workbook']['sheets'] if s['name'] == '经营驾驶舱')
        self.assertTrue(any(c.get('value') == '新增线' for c in sheet['cells'].values()))

    def test_native_table_warning_is_suppressed_for_multilevel_headers(self):
        from excel_ir_mvp.excel_ir_plus import parse_workbook_plus, rebuild_workbook_plus, diff_workbooks_plus
        ir = parse_workbook_plus(str(fixture_path('complex_report.xlsx')))
        table = ir['workbook']['sheets'][0]['extra']['tables'][0]
        self.assertFalse(table.get('native_table_supported'))
        self.assertEqual(table.get('table_kind'), 'semantic')
        self.assertEqual(table.get('native_table_skip_reason'), 'merged_cells_intersect_table')
        out = ROOT / 'semantic_table_rebuilt.xlsx'
        with warnings.catch_warnings(record=True) as seen:
            warnings.simplefilter('always')
            rebuild_workbook_plus(copy.deepcopy(ir), str(out))
        messages = [str(w.message) for w in seen]
        self.assertFalse(any('column headings must be strings' in m for m in messages))
        diff = diff_workbooks_plus(str(fixture_path('complex_report.xlsx')), str(out))
        self.assertEqual(diff['diff_count'], 0)


if __name__ == '__main__':
    unittest.main()
