from __future__ import annotations

import json
import os
import subprocess
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / 'src'
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from tests.fixtures_loader import fixture_path


class ExcelIRMVPTests(unittest.TestCase):
    def run_cmd(self, args, env=None):
        p = subprocess.run(args, cwd=ROOT, env=env, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        self.assertEqual(p.returncode, 0, msg=p.stderr[-1200:] + p.stdout[-1200:])
        return p

    def module_env(self):
        env = os.environ.copy()
        env['PYTHONPATH'] = str(SRC) + (os.pathsep + env['PYTHONPATH'] if env.get('PYTHONPATH') else '')
        return env

    def test_package_import(self):
        import excel_ir_mvp
        self.assertEqual(excel_ir_mvp.__version__, '2.0.0a8')
        self.assertTrue(callable(excel_ir_mvp.parse_workbook_plus))

    def test_module_cli_smoke(self):
        self.run_cmd(['python3', '-m', 'excel_ir_mvp.excel_ir_cli', 'doctor'], env=self.module_env())

    def test_package_module_main_smoke(self):
        self.run_cmd(['python3', '-m', 'excel_ir_mvp', 'doctor'], env=self.module_env())

    def test_api_parse_diff_validate(self):
        from excel_ir_mvp.excel_ir_plus import parse_workbook_plus, rebuild_workbook_plus, diff_workbooks_plus
        from excel_ir_mvp.validate_ir import validate_basic_types, validate_json_schema
        ir = parse_workbook_plus(str(fixture_path('complex_report.xlsx')))
        self.assertFalse(validate_json_schema(ir, 'ir.schema.json'))
        self.assertFalse(validate_basic_types(ir))
        out = ROOT / 'api_rebuilt.xlsx'
        rebuild_workbook_plus(ir, str(out))
        diff = diff_workbooks_plus(str(fixture_path('complex_report.xlsx')), str(out))
        self.assertEqual(diff['diff_count'], 0)

    def test_excel_ir_plus_small_helpers(self):
        from excel_ir_mvp import excel_ir_plus as xp
        self.assertEqual(xp.formula_ref_to_range("'Sheet 1'!$A$1:$B$2")['ref'], 'A1:B2')
        self.assertEqual(xp.formula_ref_to_range('$C$3')['min_row'], 3)
        self.assertIsNone(xp.formula_ref_to_range(None))
        self.assertEqual(xp.chart_title_text(None), None)
        self.assertEqual(xp.table_style_to_dict(None), {})
        self.assertIsNone(xp.table_style_from_dict({}))
        self.assertEqual(xp._strip_none({'a': 1, 'b': None, 'c': [1, None]}), {'a': 1, 'c': [1]})

    def test_models_validation_errors(self):
        from excel_ir_mvp.models import CellIR, SheetIR, TableIR, WorkbookIR, validate_basic_types
        self.assertTrue(CellIR(row=0, col=1).validate('A1'))
        self.assertTrue(TableIR(name='', ref='bad', field_map={'x': '1'}).validate())
        self.assertTrue(SheetIR(name='', cells={'bad': CellIR(row=1, col=1)}).validate())
        self.assertTrue(WorkbookIR(schema_version='', sheets=[]).validate())
        self.assertTrue(validate_basic_types({'schema_version': 'x', 'workbook': {'sheets': [{'name': '', 'cells': {'BAD': {'row': 0, 'col': 0}}}]}}))

    def test_reports_and_review_render(self):
        from excel_ir_mvp import audit_report, diff_report
        from excel_ir_mvp.field_map_review_app import main as review_main
        diff = {'diff_count': 1, 'truncated': False, 'diffs': [{'sheet': 'S', 'coord': 'A1', 'type': 'value', 'a': 1, 'b': 2}]}
        plan = {'validation': [], 'actions': [{'op': 'set_cell', 'sheet': 'S', 'coord': 'A1', 'value': 2}]}
        tx = {'ok': True, 'impact': {'cells_changed': 1}, 'impact_graph': {'nodes': {}}, 'actions': [{'index': 1, 'op': 'set_cell', 'cell_diffs_sample': [{'sheet': 'S', 'coord': 'A1'}]}]}
        html = diff_report.render(diff, 'Diff Title', plan, tx)
        self.assertIn('Diff Title', html)
        self.assertIn('Patch Dry-run Plan', html)
        self.assertIn('Transaction Apply Log', html)
        audit_html = audit_report.render(tx, 'Audit Title')
        self.assertIn('Audit Title', audit_html)
        out = ROOT / 'review_api.html'
        review_main([str(fixture_path('complex_ir_v07.json')), str(out)])
        self.assertIn('Field Map Review App', out.read_text(encoding='utf-8'))

    def test_corpus_runner_api(self):
        import corpus_runner
        config = corpus_runner.load_config('tests/fixtures/corpus_config.json')
        self.assertGreaterEqual(len(config['samples']), 4)
        config = dict(config)
        config['output_dir'] = 'corpus_results_api'
        summary = corpus_runner.run_corpus(config)
        self.assertTrue(summary['ok'])
        self.assertTrue((ROOT / summary['report_html']).exists())
        self.assertIn('synthetic_complex', summary['categories'])
        self.assertEqual(summary['results'][0]['diff_count'], 0)

    def test_bench_budget_api(self):
        import bench
        summary = bench.run_bench(max_total_seconds=180.0)
        self.assertTrue(summary['ok'])
        self.assertLessEqual(summary['total_seconds'], 180.0)

    def test_validate_load_and_schema_errors(self):
        from excel_ir_mvp.validate_ir import load, validate_json_schema
        self.assertIn('required', load('ir.schema.json'))
        errors = validate_json_schema({'actions': []}, 'patch.schema.json')
        self.assertEqual(errors, [])
        errors = validate_json_schema({}, 'patch.schema.json')
        self.assertTrue(any(e.get('level') == 'error' for e in errors))

    def test_roundtrip_diff_zero(self):
        self.run_cmd(['python3', 'excel_ir_cli.py', 'parse', str(fixture_path('complex_report.xlsx')), 'unittest_ir.json'])
        self.run_cmd(['python3', 'excel_ir_cli.py', 'rebuild', 'unittest_ir.json', 'unittest_rebuilt.xlsx'])
        self.run_cmd(['python3', 'excel_ir_cli.py', 'diff', str(fixture_path('complex_report.xlsx')), 'unittest_rebuilt.xlsx', 'unittest_diff.json'])
        diff = json.loads((ROOT / 'unittest_diff.json').read_text(encoding='utf-8'))
        self.assertEqual(diff['diff_count'], 0)

    def test_patch_log_has_impact(self):
        self.run_cmd(['python3', 'excel_ir_cli.py', 'patch', str(fixture_path('complex_ir_v07.json')), str(fixture_path('v08_patch.json')), 'unittest_patched_ir.json', '--log', 'unittest_tx.json'])
        log = json.loads((ROOT / 'unittest_tx.json').read_text(encoding='utf-8'))
        self.assertTrue(log['ok'])
        self.assertIn('impact_graph', log)
        self.assertTrue(any(a.get('cell_diffs_sample') for a in log['actions']))

    def test_validate(self):
        self.run_cmd(['python3', 'excel_ir_cli.py', 'validate', 'ir', str(fixture_path('complex_ir_v07.json'))])
        self.run_cmd(['python3', 'excel_ir_cli.py', 'validate', 'patch', str(fixture_path('v08_patch.json'))])
    def test_package_corpus_runner_helpers(self):
        from excel_ir_mvp import corpus_runner as pkg_corpus
        results = [
            {'category': 'synthetic_complex', 'parse_ok': True, 'rebuild_ok': True, 'diff_count': 0},
            {'category': 'semantic_table', 'parse_ok': False, 'rebuild_ok': False, 'diff_count': 2},
        ]
        cats = pkg_corpus._category_summary(results)
        self.assertEqual(cats['synthetic_complex']['ok'], 1)
        self.assertEqual(cats['semantic_table']['failed'], 1)
        self.assertTrue(pkg_corpus.load_config('corpus_config.json')['samples'])
        listed = pkg_corpus.list_samples(pkg_corpus.load_config('corpus_config.json'))
        self.assertIn('metadata_roundtrip', listed['categories'])
        self.assertIn('native_table', listed['categories'])
        self.assertIn('semantic_table', listed['categories'])
        html = pkg_corpus.render_summary_html({'ok': True, 'categories': cats, 'results': results})
        self.assertIn('Excel IR Corpus Report', html)

    def test_validate_ir_error_branches(self):
        from excel_ir_mvp import validate_ir
        self.assertTrue(any(e['level'] == 'error' for e in validate_ir.validate_basic_types({'schema_version': 'x'})))
        self.assertTrue(any(e['level'] == 'error' for e in validate_ir.validate_json_schema({'bad': True}, 'patch.schema.json')))

    def test_cli_main_metadata_diff_inprocess(self):
        import contextlib, io
        from excel_ir_mvp.excel_ir_cli import main as cli_main
        from excel_ir_mvp.excel_ir_plus import collect_semantic_metadata, parse_workbook_plus, save_json, load_json
        ir = parse_workbook_plus(str(fixture_path('complex_report.xlsx')))
        a = ROOT / 'cli_main_meta_a.json'
        b = ROOT / 'cli_main_meta_b.json'
        out = ROOT / 'cli_main_meta_diff.json'
        save_json(collect_semantic_metadata(ir), str(a))
        save_json(collect_semantic_metadata(ir), str(b))
        saved_argv = sys.argv[:]
        buf = io.StringIO()
        try:
            sys.argv = ['excel-ir', 'metadata', 'diff', str(a), str(b), str(out)]
            with contextlib.redirect_stdout(buf):
                cli_main()
        finally:
            sys.argv = saved_argv
        self.assertEqual(load_json(str(out))['diff_count'], 0)
    def test_cli_main_corpus_list_report_inprocess(self):
        import contextlib, io
        from excel_ir_mvp.excel_ir_cli import main as cli_main
        saved_argv = sys.argv[:]
        buf = io.StringIO()
        try:
            sys.argv = ['excel-ir', 'corpus', 'list', '--config', 'corpus_config.json']
            with contextlib.redirect_stdout(buf):
                cli_main()
        finally:
            sys.argv = saved_argv
        self.assertIn('synthetic_complex', json.loads(buf.getvalue())['categories'])

        summary = {'ok': True, 'categories': {'synthetic_complex': {'count': 1, 'ok': 1, 'failed': 0, 'diff_count_total': 0}}, 'results': [{'name': 's', 'category': 'synthetic_complex', 'parse_ok': True, 'rebuild_ok': True, 'diff_count': 0}]}
        summary_path = ROOT / 'cli_corpus_summary.json'
        html_path = ROOT / 'cli_corpus_report.html'
        summary_path.write_text(json.dumps(summary), encoding='utf-8')
        saved_argv = sys.argv[:]
        try:
            sys.argv = ['excel-ir', 'corpus', 'report', str(summary_path), str(html_path)]
            with contextlib.redirect_stdout(io.StringIO()):
                cli_main()
        finally:
            sys.argv = saved_argv
        self.assertIn('Excel IR Corpus Report', html_path.read_text(encoding='utf-8'))


if __name__ == '__main__':
    unittest.main()
