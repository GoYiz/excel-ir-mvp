from __future__ import annotations

import argparse
import json
from pathlib import Path

try:
    from . import excel_ir_plus, ir_patch, corpus_runner, diff_report, audit_report, validate_ir, bench, field_map_review_app
except ImportError:  # flat-source dev fallback
    import excel_ir_plus
    import ir_patch
    import corpus_runner
    import diff_report
    import audit_report
    import validate_ir
    import bench
    import field_map_review_app


def main():
    ap = argparse.ArgumentParser(prog='excel-ir', description='Excel IR MVP v1.0 unified CLI')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('parse')
    p.add_argument('xlsx'); p.add_argument('ir_json')

    p = sub.add_parser('rebuild')
    p.add_argument('ir_json'); p.add_argument('xlsx')

    p = sub.add_parser('diff')
    p.add_argument('a'); p.add_argument('b'); p.add_argument('diff_json')

    p = sub.add_parser('patch')
    p.add_argument('ir_json'); p.add_argument('patch_json'); p.add_argument('out_ir', nargs='?')
    p.add_argument('--dry-run', action='store_true'); p.add_argument('--plan'); p.add_argument('--log')

    p = sub.add_parser('report')
    p.add_argument('diff_json'); p.add_argument('html'); p.add_argument('--title', default='Excel IR Report'); p.add_argument('--plan'); p.add_argument('--log')

    p = sub.add_parser('audit')
    p.add_argument('tx_log'); p.add_argument('html'); p.add_argument('--title', default='Excel IR Audit Report')

    p = sub.add_parser('corpus')
    p.add_argument('--config')

    p = sub.add_parser('validate')
    p.add_argument('kind', choices=['ir', 'patch']); p.add_argument('json_file')

    p = sub.add_parser('field-map-review')
    p.add_argument('ir_json'); p.add_argument('html')

    p = sub.add_parser('metadata')
    meta_sub = p.add_subparsers(dest='metadata_cmd', required=True)
    p = meta_sub.add_parser('export')
    p.add_argument('ir_json'); p.add_argument('metadata_json')
    p = meta_sub.add_parser('import')
    p.add_argument('ir_json'); p.add_argument('metadata_json'); p.add_argument('out_ir')
    p = meta_sub.add_parser('diff')
    p.add_argument('a_metadata_json'); p.add_argument('b_metadata_json'); p.add_argument('diff_json', nargs='?')
    p = meta_sub.add_parser('verify')
    p.add_argument('metadata_json')

    sub.add_parser('doctor')
    sub.add_parser('bench')

    args = ap.parse_args()
    if args.cmd == 'parse':
        excel_ir_plus.save_json(excel_ir_plus.parse_workbook_plus(args.xlsx), args.ir_json)
        print(json.dumps({'ok': True, 'output': args.ir_json}, ensure_ascii=False))
    elif args.cmd == 'rebuild':
        excel_ir_plus.rebuild_workbook_plus(excel_ir_plus.load_json(args.ir_json), args.xlsx)
        print(json.dumps({'ok': True, 'output': args.xlsx}, ensure_ascii=False))
    elif args.cmd == 'diff':
        d = excel_ir_plus.diff_workbooks_plus(args.a, args.b)
        excel_ir_plus.save_json(d, args.diff_json)
        print(json.dumps(d, ensure_ascii=False, indent=2))
    elif args.cmd == 'patch':
        ir = ir_patch.load_json(args.ir_json); patch = ir_patch.load_json(args.patch_json)
        plan = ir_patch.dry_run(ir, patch)
        if args.dry_run:
            if args.plan: ir_patch.save_json(plan, args.plan)
            print(json.dumps(plan, ensure_ascii=False, indent=2)); return
        if not args.out_ir: raise SystemExit('out_ir required unless --dry-run')
        out, log = ir_patch.apply_patch_with_log(ir, patch)
        ir_patch.save_json(out, args.out_ir)
        if args.plan: ir_patch.save_json(plan, args.plan)
        if args.log: ir_patch.save_json(log, args.log)
        print(json.dumps({'ok': True, 'output': args.out_ir, 'plan': args.plan, 'log': args.log}, ensure_ascii=False))
    elif args.cmd == 'report':
        diff = diff_report.load(args.diff_json)
        plan = diff_report.load(args.plan) if args.plan else None
        tx = diff_report.load(args.log) if args.log else None
        Path(args.html).write_text(diff_report.render(diff, args.title, plan, tx), encoding='utf-8')
        print(json.dumps({'ok': True, 'output': args.html}, ensure_ascii=False))
    elif args.cmd == 'audit':
        log = json.loads(Path(args.tx_log).read_text(encoding='utf-8'))
        Path(args.html).write_text(audit_report.render(log, args.title), encoding='utf-8')
        print(json.dumps({'ok': True, 'output': args.html}, ensure_ascii=False))
    elif args.cmd == 'corpus':
        summary = corpus_runner.run_corpus(corpus_runner.load_config(args.config))
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    elif args.cmd == 'validate':
        data = json.loads(Path(args.json_file).read_text(encoding='utf-8'))
        schema = 'ir.schema.json' if args.kind == 'ir' else 'patch.schema.json'
        errors = validate_ir.validate_json_schema(data, schema)
        if args.kind == 'ir':
            try:
                from .formula_utils import workbook_formula_dependencies
            except ImportError:
                from formula_utils import workbook_formula_dependencies
            result = {'ok': not errors, 'errors': errors, 'formula_dependencies': len(workbook_formula_dependencies(data))}
        else:
            result = {'ok': not errors, 'errors': errors}
        print(json.dumps(result, ensure_ascii=False, indent=2))
        if not result['ok']: raise SystemExit(1)
    elif args.cmd == 'field-map-review':
        field_map_review_app.main([args.ir_json, args.html])
    elif args.cmd == 'metadata':
        if args.metadata_cmd == 'export':
            ir = excel_ir_plus.load_json(args.ir_json)
            metadata = excel_ir_plus.export_semantic_metadata_from_ir(ir, args.metadata_json)
            print(json.dumps({'ok': True, 'output': args.metadata_json, 'tables': sum(len(s.get('tables', [])) for s in metadata.get('sheets', []))}, ensure_ascii=False))
        elif args.metadata_cmd == 'import':
            ir = excel_ir_plus.load_json(args.ir_json)
            metadata = excel_ir_plus.load_json(args.metadata_json)
            out = excel_ir_plus.import_semantic_metadata_to_ir(ir, metadata)
            excel_ir_plus.save_json(out, args.out_ir)
            print(json.dumps({'ok': True, 'output': args.out_ir}, ensure_ascii=False))
        elif args.metadata_cmd == 'diff':
            d = excel_ir_plus.semantic_metadata_diff_files(args.a_metadata_json, args.b_metadata_json)
            if args.diff_json:
                excel_ir_plus.save_json(d, args.diff_json)
            print(json.dumps(d, ensure_ascii=False, indent=2))
        elif args.metadata_cmd == 'verify':
            result = excel_ir_plus.verify_semantic_metadata_file(args.metadata_json)
            print(json.dumps(result, ensure_ascii=False, indent=2))
            if not result.get('ok'):
                raise SystemExit(1)
    elif args.cmd == 'doctor':
        import sys, openpyxl
        result = {'ok': True, 'python': sys.version, 'openpyxl': openpyxl.__version__, 'project_root': str(Path(__file__).resolve().parent)}
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif args.cmd == 'bench':
        bench.main()

if __name__ == '__main__':
    main()
