from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent

SOURCE_COMMANDS = [
    ['python3', '-m', 'py_compile', 'excel_ir_cli.py', 'src/excel_ir_mvp/excel_ir_cli.py', 'src/excel_ir_mvp/excel_ir_plus.py', 'src/excel_ir_mvp/ir_patch.py', 'src/excel_ir_mvp/formula_utils.py', 'src/excel_ir_mvp/validate_ir.py', 'src/excel_ir_mvp/models.py', 'src/excel_ir_mvp/backends.py', 'src/excel_ir_mvp/__main__.py'],
    ['python3', '-m', 'unittest', '-v', 'tests.test_excel_ir_mvp', 'tests.test_patch_ops', 'tests.test_native_tables', 'tests.test_metadata'],
    ['python3', '-m', 'coverage', 'erase'],
    ['python3', '-m', 'coverage', 'run', '--parallel-mode', '--source=src/excel_ir_mvp', '-m', 'unittest',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_package_import',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_module_cli_smoke',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_package_module_main_smoke',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_api_parse_diff_validate',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_excel_ir_plus_small_helpers',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_models_validation_errors',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_reports_and_review_render',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_validate_load_and_schema_errors',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_validate_ir_error_branches',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_cli_main_metadata_diff_inprocess',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_cli_main_corpus_list_report_inprocess',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_inprocess_auxiliary_entrypoints_for_coverage'],
    ['python3', '-m', 'coverage', 'run', '--parallel-mode', '--source=src/excel_ir_mvp', '-m', 'unittest',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_corpus_runner_api',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_bench_budget_api',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_roundtrip_diff_zero',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_patch_log_has_impact',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_validate',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_package_corpus_runner_helpers',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_compare_ir_cli_and_metadata_strip',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_alpha10_anonymize_status_and_compare_modes',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_alpha12_stream_edit_preview_all_and_offsets',
     'tests.test_excel_ir_mvp.ExcelIRMVPTests.test_alpha13_backend_registry_and_engine_cli'],
    ['python3', '-m', 'coverage', 'run', '--parallel-mode', '--source=src/excel_ir_mvp', '-m', 'unittest', 'tests.test_patch_ops'],
    ['python3', '-m', 'coverage', 'run', '--parallel-mode', '--source=src/excel_ir_mvp', '-m', 'unittest', 'tests.test_native_tables', 'tests.test_metadata'],
    ['python3', '-m', 'coverage', 'combine'],
    ['python3', '-m', 'coverage', 'report', '--show-missing', '--fail-under=70'],
    ['python3', 'excel_ir_cli.py', 'validate', 'ir', 'tests/fixtures/complex_ir_v07.json'],
    ['python3', 'excel_ir_cli.py', 'validate', 'patch', 'tests/fixtures/v08_patch.json'],
    ['python3', 'golden_tests.py'],
    ['python3', 'corpus_runner.py'],
    ['python3', 'corpus_runner.py', 'report', 'corpus_results/summary.json', 'corpus_results/report.html'],
    ['python3', 'excel_ir_cli.py', 'inspect', 'tests/fixtures/complex_report.xlsx', '--out', 'ci_inspect.json', '--engine', 'openpyxl'],
    ['python3', 'excel_ir_cli.py', 'engines'],
    ['python3', 'excel_ir_cli.py', 'compare-ir', 'tests/fixtures/complex_ir_v07.json', 'tests/fixtures/complex_ir_v07.json', 'ci_compare_ir.json'],
    ['python3', 'excel_ir_cli.py', 'compare-ir', '--semantic-only', 'tests/fixtures/complex_ir_v07.json', 'tests/fixtures/complex_ir_v07.json', 'ci_compare_ir_semantic.json'],
    ['python3', 'excel_ir_cli.py', 'anonymize', 'tests/fixtures/complex_report.xlsx', 'ci_anonymized.xlsx'],
    ['python3', 'excel_ir_cli.py', 'stream-edit', 'tests/fixtures/complex_report.xlsx', 'ci_stream_edit.xlsx', '--match', '总计', '--value', '合计'],
    ['python3', 'excel_ir_cli.py', 'stream-edit', 'tests/fixtures/complex_report.xlsx', 'ci_stream_edit_right.xlsx', '--match', '备注', '--value', '说明', '--start', 'right'],
    ['python3', 'excel_ir_cli.py', 'stream-edit', 'tests/fixtures/complex_report.xlsx', 'ci_stream_edit_preview.xlsx', '--match', '业务线', '--value', '收入本月', '--offset-row', '1', '--offset-col', '2', '--preview'],
    ['python3', 'excel_ir_cli.py', 'stream-edit', 'tests/fixtures/complex_report.xlsx', 'ci_stream_edit_all.xlsx', '--match', '云业务', '--value', '云事业部', '--all'],
]

INSTALLED_COMMANDS = [
    ['excel-ir', 'doctor'],
    ['excel-ir', 'engines'],
    ['python3', '-m', 'excel_ir_mvp', 'doctor'],
    ['excel-ir', 'validate', 'ir', 'tests/fixtures/complex_ir_v07.json'],
    ['excel-ir', 'validate', 'patch', 'tests/fixtures/v08_patch.json'],
    ['excel-ir', 'field-map-review', 'tests/fixtures/complex_ir_v07.json', 'ci_installed_field_map.html'],
    ['excel-ir', 'metadata', 'export', 'tests/fixtures/complex_ir_v07.json', 'ci_installed_metadata.json'],
    ['excel-ir', 'metadata', 'diff', 'ci_installed_metadata.json', 'ci_installed_metadata.json', 'ci_installed_metadata_diff.json'],
    ['excel-ir', 'metadata', 'verify', 'ci_installed_metadata.json'],
    ['excel-ir', 'metadata', 'repair', 'ci_installed_repaired.xlsx', '--from-xlsx', 'tests/fixtures/complex_report.xlsx'],
    ['excel-ir', 'metadata', 'verify', '--from-xlsx', 'ci_installed_repaired.xlsx'],
    ['excel-ir', 'inspect', 'tests/fixtures/complex_report.xlsx', '--out', 'ci_installed_inspect.json', '--engine', 'openpyxl'],
    ['excel-ir', 'compare-ir', 'tests/fixtures/complex_ir_v07.json', 'tests/fixtures/complex_ir_v07.json', 'ci_installed_compare_ir.json'],
    ['excel-ir', 'metadata', 'strip', 'ci_installed_stripped.xlsx', '--from-xlsx', 'ci_installed_repaired.xlsx'],
    ['excel-ir', 'metadata', 'status', 'ci_installed_repaired.xlsx'],
    ['excel-ir', 'compare-ir', '--structural-only', 'tests/fixtures/complex_ir_v07.json', 'tests/fixtures/complex_ir_v07.json', 'ci_installed_compare_ir_structural.json'],
    ['excel-ir', 'anonymize', 'tests/fixtures/complex_report.xlsx', 'ci_installed_anonymized.xlsx'],
    ['excel-ir', 'stream-edit', 'tests/fixtures/complex_report.xlsx', 'ci_installed_stream_edit.xlsx', '--match', '总计', '--value', '合计'],
    ['excel-ir', 'stream-edit', 'tests/fixtures/complex_report.xlsx', 'ci_installed_stream_edit_preview.xlsx', '--match', '业务线', '--value', '收入本月', '--offset-row', '1', '--offset-col', '2', '--preview'],
]


def run_commands(commands):
    results = []
    for cmd in commands:
        p = subprocess.run(cmd, cwd=ROOT, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        results.append({'cmd': cmd, 'returncode': p.returncode, 'stdout_tail': p.stdout[-1000:], 'stderr_tail': p.stderr[-1000:]})
    return results


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    installed = '--installed' in argv
    commands = INSTALLED_COMMANDS if installed else SOURCE_COMMANDS
    results = run_commands(commands)
    ok = all(r['returncode'] == 0 for r in results)
    out = {'ok': ok, 'mode': 'installed' if installed else 'source', 'results': results}
    out_name = 'ci_results_installed.json' if installed else 'ci_results.json'
    (ROOT / out_name).write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(out, ensure_ascii=False, indent=2))
    if not ok:
        raise SystemExit(1)

if __name__ == '__main__':
    main()
