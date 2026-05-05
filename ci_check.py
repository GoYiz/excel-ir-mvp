from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent

SOURCE_COMMANDS = [
    ['python3', '-m', 'py_compile', 'excel_ir_cli.py', 'src/excel_ir_mvp/excel_ir_cli.py', 'src/excel_ir_mvp/excel_ir_plus.py', 'src/excel_ir_mvp/ir_patch.py', 'src/excel_ir_mvp/formula_utils.py', 'src/excel_ir_mvp/validate_ir.py', 'src/excel_ir_mvp/models.py', 'src/excel_ir_mvp/__main__.py'],
    ['python3', '-m', 'unittest', '-v', 'tests.test_excel_ir_mvp', 'tests.test_patch_ops', 'tests.test_native_tables', 'tests.test_metadata'],
    ['python3', '-m', 'coverage', 'erase'],
    ['python3', '-m', 'coverage', 'run', '--parallel-mode', '--source=src/excel_ir_mvp', '-m', 'unittest', 'tests.test_excel_ir_mvp', 'tests.test_patch_ops'],
    ['python3', '-m', 'coverage', 'run', '--parallel-mode', '--source=src/excel_ir_mvp', '-m', 'unittest', 'tests.test_native_tables', 'tests.test_metadata'],
    ['python3', '-m', 'coverage', 'combine'],
    ['python3', '-m', 'coverage', 'report', '--show-missing', '--fail-under=65'],
    ['python3', 'excel_ir_cli.py', 'validate', 'ir', 'tests/fixtures/complex_ir_v07.json'],
    ['python3', 'excel_ir_cli.py', 'validate', 'patch', 'tests/fixtures/v08_patch.json'],
    ['python3', 'golden_tests.py'],
    ['python3', 'corpus_runner.py'],
]

INSTALLED_COMMANDS = [
    ['excel-ir', 'doctor'],
    ['python3', '-m', 'excel_ir_mvp', 'doctor'],
    ['excel-ir', 'validate', 'ir', 'tests/fixtures/complex_ir_v07.json'],
    ['excel-ir', 'validate', 'patch', 'tests/fixtures/v08_patch.json'],
    ['excel-ir', 'field-map-review', 'tests/fixtures/complex_ir_v07.json', 'ci_installed_field_map.html'],
    ['excel-ir', 'metadata', 'export', 'tests/fixtures/complex_ir_v07.json', 'ci_installed_metadata.json'],
    ['excel-ir', 'metadata', 'diff', 'ci_installed_metadata.json', 'ci_installed_metadata.json', 'ci_installed_metadata_diff.json'],
    ['excel-ir', 'metadata', 'verify', 'ci_installed_metadata.json'],
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
