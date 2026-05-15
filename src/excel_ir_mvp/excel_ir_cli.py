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


def _parse_header_path(text):
    try:
        data = json.loads(text)
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return [x for x in str(text).split('/') if x != '']


def _parse_axis_range(text, *, numeric=True):
    parts = str(text).split(':', 1)
    if len(parts) == 1:
        start = end = parts[0]
    else:
        start, end = parts[0], parts[1]
    if numeric:
        start, end = int(start), int(end)
        if start < 1 or end < start:
            raise ValueError('range must be START:END')
    else:
        start, end = _parse_col_arg(start), _parse_col_arg(end)
        if start < 1 or end < start:
            raise ValueError('column range must be START:END')
    return start, end


def _parse_row_range(text):
    return _parse_axis_range(text, numeric=True)


def _parse_col_arg(value):
    if value is None:
        return None
    text = str(value).strip()
    if text.isdigit():
        return int(text)
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(text.upper())


def main():
    ap = argparse.ArgumentParser(prog='excel-ir', description='Excel IR MVP v1.0 unified CLI')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('parse')
    p.add_argument('xlsx'); p.add_argument('ir_json')
    p.add_argument('--engine', default='openpyxl', choices=['openpyxl', 'wolfxl', 'auto'])
    p.add_argument('--sheet', action='append', dest='sheets', help='Sheet name to parse; repeat for multiple sheets')
    p.add_argument('--profile', choices=['full', 'fast'], default='full')
    p.add_argument('--fast', action='store_true', help='Shortcut for --profile fast')
    p.add_argument('--no-formula-cache', action='store_true')
    p.add_argument('--no-extra', action='store_true')
    p.add_argument('--no-images', action='store_true')
    p.add_argument('--no-charts', action='store_true')
    p.add_argument('--no-binary', action='store_true')
    p.add_argument('--read-only', action='store_true', help='Use openpyxl read-only streaming mode where possible')
    p.add_argument('--dense', action='store_true', help='Use rectangular worksheet iteration instead of sparse cells')

    p = sub.add_parser('rebuild')
    p.add_argument('ir_json'); p.add_argument('xlsx')
    p.add_argument('--engine', default='openpyxl', choices=['openpyxl', 'wolfxl', 'auto'])
    p.add_argument('--sheet', action='append', dest='sheets', help='Sheet name to rebuild; repeat for multiple sheets')

    p = sub.add_parser('diff')
    p.add_argument('a'); p.add_argument('b'); p.add_argument('diff_json')
    p.add_argument('--engine', default='openpyxl', choices=['openpyxl', 'wolfxl', 'auto'])

    p = sub.add_parser('compare-ir')
    p.add_argument('a_ir'); p.add_argument('b_ir'); p.add_argument('diff_json', nargs='?')
    mode = p.add_mutually_exclusive_group()
    mode.add_argument('--semantic-only', action='store_true')
    mode.add_argument('--structural-only', action='store_true')

    p = sub.add_parser('anonymize')
    p.add_argument('xlsx'); p.add_argument('out_xlsx')
    p.add_argument('--rewrite-formulas', action='store_true')
    p.add_argument('--engine', default='openpyxl', choices=['openpyxl', 'wolfxl', 'auto'])

    p = sub.add_parser('inspect')
    p.add_argument('xlsx'); p.add_argument('--out')
    p.add_argument('--engine', default='openpyxl', choices=['openpyxl', 'wolfxl', 'auto'])

    p = sub.add_parser('stream-edit')
    p.add_argument('xlsx'); p.add_argument('out_xlsx')
    p.add_argument('--match', required=True)
    p.add_argument('--value', required=True)
    p.add_argument('--sheet')
    p.add_argument('--start', choices=['left', 'right'], default='left')
    p.add_argument('--contains', action='store_true')
    p.add_argument('--case-sensitive', action='store_true')
    p.add_argument('--max-cells', type=int)
    p.add_argument('--offset-row', type=int, default=0)
    p.add_argument('--offset-col', type=int, default=0)
    p.add_argument('--preview', action='store_true')
    p.add_argument('--all', action='store_true', dest='update_all')
    p.add_argument('--engine', default='openpyxl', choices=['openpyxl', 'wolfxl', 'auto'])
    p.add_argument('--as-number', action='store_true')

    p = sub.add_parser('header-edit')
    p.add_argument('xlsx'); p.add_argument('out_xlsx')
    p.add_argument('--headers', required=True, help='JSON array or slash-separated path, e.g. ["2026","5","8"] or 2026/5/8')
    p.add_argument('--value', required=True)
    p.add_argument('--sheet')
    p.add_argument('--header-rows', default='1:3')
    p.add_argument('--orientation', choices=['horizontal', 'vertical'], default='horizontal')
    p.add_argument('--header-cols', default='1:3')
    p.add_argument('--row', type=int)
    p.add_argument('--row-match')
    p.add_argument('--row-match-col', default='1')
    p.add_argument('--col')
    p.add_argument('--col-match')
    p.add_argument('--col-match-row', type=int, default=1)
    p.add_argument('--data-start-row', type=int)
    p.add_argument('--data-start-col')
    p.add_argument('--min-col', default='1')
    p.add_argument('--max-col')
    p.add_argument('--min-row', type=int, default=1)
    p.add_argument('--max-row', type=int)
    p.add_argument('--contains', action='store_true')
    p.add_argument('--match-mode', choices=['exact', 'contains', 'wildcard', 'regex', 're', 'glob'], default='exact')
    p.add_argument('--case-sensitive', action='store_true')
    p.add_argument('--header-match-index', type=int, default=1)
    p.add_argument('--preview', action='store_true')
    p.add_argument('--as-number', action='store_true')
    p.add_argument('--engine', default='openpyxl', choices=['openpyxl', 'wolfxl', 'auto'])

    p = sub.add_parser('patch')
    p.add_argument('ir_json'); p.add_argument('patch_json'); p.add_argument('out_ir', nargs='?')
    p.add_argument('--dry-run', action='store_true'); p.add_argument('--plan'); p.add_argument('--log')

    p = sub.add_parser('report')
    p.add_argument('diff_json'); p.add_argument('html'); p.add_argument('--title', default='Excel IR Report'); p.add_argument('--plan'); p.add_argument('--log')

    p = sub.add_parser('audit')
    p.add_argument('tx_log'); p.add_argument('html'); p.add_argument('--title', default='Excel IR Audit Report')

    p = sub.add_parser('corpus')
    corpus_sub = p.add_subparsers(dest='corpus_cmd')
    p = corpus_sub.add_parser('run')
    p.add_argument('--config')
    p = corpus_sub.add_parser('list')
    p.add_argument('--config')
    p = corpus_sub.add_parser('report')
    p.add_argument('summary_json'); p.add_argument('html'); p.add_argument('--title', default='Excel IR Corpus Report')
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
    p = meta_sub.add_parser('extract')
    p.add_argument('metadata_json')
    p.add_argument('--from-xlsx', dest='from_xlsx', required=True)
    p.add_argument('--engine', default='openpyxl', choices=['openpyxl', 'wolfxl', 'auto'])
    p = meta_sub.add_parser('repair')
    p.add_argument('out_xlsx')
    p.add_argument('--from-xlsx', dest='from_xlsx', required=True)
    p.add_argument('--engine', default='openpyxl', choices=['openpyxl', 'wolfxl', 'auto'])
    p = meta_sub.add_parser('strip')
    p.add_argument('out_xlsx')
    p.add_argument('--from-xlsx', dest='from_xlsx', required=True)
    p.add_argument('--engine', default='openpyxl', choices=['openpyxl', 'wolfxl', 'auto'])
    p = meta_sub.add_parser('status')
    p.add_argument('xlsx')
    p.add_argument('--engine', default='openpyxl', choices=['openpyxl', 'wolfxl', 'auto'])
    p = meta_sub.add_parser('diff')
    p.add_argument('a_metadata_json'); p.add_argument('b_metadata_json'); p.add_argument('diff_json', nargs='?')
    p = meta_sub.add_parser('verify')
    p.add_argument('metadata_json', nargs='?')
    p.add_argument('--from-xlsx', dest='from_xlsx')
    p.add_argument('--engine', default='openpyxl', choices=['openpyxl', 'wolfxl', 'auto'])

    sub.add_parser('engines')
    sub.add_parser('doctor')
    sub.add_parser('bench')

    args = ap.parse_args()
    if args.cmd == 'parse':
        profile = 'fast' if args.fast else args.profile
        excel_ir_plus.save_json(excel_ir_plus.parse_workbook_plus(
            args.xlsx,
            engine=args.engine,
            sheet_names=args.sheets,
            include_formula_cache=not args.no_formula_cache,
            include_extra=not args.no_extra,
            include_images=not args.no_images,
            include_charts=not args.no_charts,
            include_binary=not args.no_binary,
            read_only=args.read_only,
            sparse=not args.dense,
            profile=profile,
        ), args.ir_json)
        print(json.dumps({'ok': True, 'output': args.ir_json}, ensure_ascii=False))
    elif args.cmd == 'rebuild':
        excel_ir_plus.rebuild_workbook_plus(excel_ir_plus.load_json(args.ir_json), args.xlsx, engine=args.engine, sheet_names=args.sheets)
        print(json.dumps({'ok': True, 'output': args.xlsx}, ensure_ascii=False))
    elif args.cmd == 'diff':
        d = excel_ir_plus.diff_workbooks_plus(args.a, args.b, engine=args.engine)
        excel_ir_plus.save_json(d, args.diff_json)
        print(json.dumps(d, ensure_ascii=False, indent=2))
    elif args.cmd == 'compare-ir':
        mode = 'semantic' if args.semantic_only else 'structural' if args.structural_only else 'full'
        d = excel_ir_plus.compare_ir_files(args.a_ir, args.b_ir, mode=mode)
        if args.diff_json:
            excel_ir_plus.save_json(d, args.diff_json)
        print(json.dumps(d, ensure_ascii=False, indent=2))
        if not d.get('ok'):
            raise SystemExit(1)
    elif args.cmd == 'anonymize':
        result = excel_ir_plus.anonymize_workbook_xlsx(args.xlsx, args.out_xlsx, keep_formulas=not args.rewrite_formulas, engine=args.engine)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif args.cmd == 'inspect':
        result = excel_ir_plus.inspect_workbook(args.xlsx, engine=args.engine)
        if args.out:
            excel_ir_plus.save_json(result, args.out)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif args.cmd == 'stream-edit':
        value = args.value
        if args.as_number:
            try:
                value = float(value) if ('.' in value) else int(value)
            except ValueError:
                raise SystemExit('--as-number requires a numeric --value')
        result = excel_ir_plus.stream_update_first_match_xlsx(
            args.xlsx, args.out_xlsx, args.match, value,
            sheet=args.sheet, start=args.start, contains=args.contains,
            case_sensitive=args.case_sensitive, max_cells=args.max_cells,
            offset_row=args.offset_row, offset_col=args.offset_col,
            preview=args.preview, update_all=args.update_all, engine=args.engine,
        )
        print(json.dumps(result, ensure_ascii=False, indent=2))
        if not result.get('found'):
            raise SystemExit(1)
    elif args.cmd == 'header-edit':
        value = args.value
        if args.as_number:
            try:
                value = float(value) if ('.' in value) else int(value)
            except ValueError:
                raise SystemExit('--as-number requires a numeric --value')
        try:
            header_start, header_end = _parse_row_range(args.header_rows)
            header_start_col, header_end_col = _parse_axis_range(args.header_cols, numeric=False)
            headers = _parse_header_path(args.headers)
            result = excel_ir_plus.update_cell_by_multi_header_xlsx(
                args.xlsx, args.out_xlsx, headers, value,
                sheet=args.sheet,
                header_start_row=header_start, header_end_row=header_end,
                orientation=args.orientation,
                header_start_col=header_start_col, header_end_col=header_end_col,
                row=args.row, row_match=args.row_match, row_match_col=args.row_match_col,
                col=args.col, col_match=args.col_match, col_match_row=args.col_match_row,
                data_start_row=args.data_start_row, data_start_col=args.data_start_col,
                min_col=_parse_col_arg(args.min_col) or 1,
                max_col=_parse_col_arg(args.max_col),
                min_row=args.min_row, max_row=args.max_row,
                contains=args.contains, case_sensitive=args.case_sensitive,
                match_mode=args.match_mode,
                header_match_index=args.header_match_index,
                preview=args.preview, engine=args.engine,
            )
        except ValueError as exc:
            raise SystemExit(str(exc))
        print(json.dumps(result, ensure_ascii=False, indent=2))
        if not result.get('found'):
            raise SystemExit(1)
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
        cmd = args.corpus_cmd or 'run'
        if cmd == 'list':
            print(json.dumps(corpus_runner.list_samples(corpus_runner.load_config(args.config)), ensure_ascii=False, indent=2))
        elif cmd == 'report':
            summary = json.loads(Path(args.summary_json).read_text(encoding='utf-8'))
            corpus_runner.write_report(summary, args.html, args.title)
            print(json.dumps({'ok': True, 'output': args.html}, ensure_ascii=False))
        else:
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
        elif args.metadata_cmd == 'extract':
            metadata = excel_ir_plus.extract_semantic_metadata_from_xlsx(args.from_xlsx, engine=args.engine)
            excel_ir_plus.save_json(metadata, args.metadata_json)
            print(json.dumps({'ok': True, 'output': args.metadata_json, 'tables': sum(len(s.get('tables', [])) for s in metadata.get('sheets', []))}, ensure_ascii=False))
        elif args.metadata_cmd == 'repair':
            result = excel_ir_plus.repair_semantic_metadata_xlsx(args.from_xlsx, args.out_xlsx, engine=args.engine)
            print(json.dumps(result, ensure_ascii=False, indent=2))
            if not result.get('ok'):
                raise SystemExit(1)
        elif args.metadata_cmd == 'strip':
            result = excel_ir_plus.strip_semantic_metadata_xlsx(args.from_xlsx, args.out_xlsx, engine=args.engine)
            print(json.dumps(result, ensure_ascii=False, indent=2))
            if not result.get('ok'):
                raise SystemExit(1)
        elif args.metadata_cmd == 'status':
            print(json.dumps(excel_ir_plus.metadata_status_xlsx(args.xlsx, engine=args.engine), ensure_ascii=False, indent=2))
        elif args.metadata_cmd == 'diff':
            d = excel_ir_plus.semantic_metadata_diff_files(args.a_metadata_json, args.b_metadata_json)
            if args.diff_json:
                excel_ir_plus.save_json(d, args.diff_json)
            print(json.dumps(d, ensure_ascii=False, indent=2))
        elif args.metadata_cmd == 'verify':
            if args.from_xlsx:
                result = excel_ir_plus.verify_semantic_metadata_xlsx(args.from_xlsx, engine=args.engine)
            else:
                if not args.metadata_json:
                    raise SystemExit('metadata_json required unless --from-xlsx is used')
                result = excel_ir_plus.verify_semantic_metadata_file(args.metadata_json)
            print(json.dumps(result, ensure_ascii=False, indent=2))
            if not result.get('ok'):
                raise SystemExit(1)
    elif args.cmd == 'engines':
        print(json.dumps({'ok': True, 'default': 'openpyxl', 'available': excel_ir_plus.available_engines(), 'engines': excel_ir_plus.engine_status()}, ensure_ascii=False, indent=2))
    elif args.cmd == 'doctor':
        import sys, openpyxl
        result = {'ok': True, 'python': sys.version, 'openpyxl': openpyxl.__version__, 'engines': excel_ir_plus.engine_status(), 'project_root': str(Path(__file__).resolve().parent)}
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif args.cmd == 'bench':
        bench.main()

if __name__ == '__main__':
    main()
