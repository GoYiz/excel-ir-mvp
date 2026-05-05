from __future__ import annotations

import argparse
import json
from pathlib import Path
try:
    from importlib import resources as importlib_resources
except Exception:  # pragma: no cover
    importlib_resources = None

try:
    import jsonschema
except Exception:
    jsonschema = None

try:
    from . import ir_patch, excel_ir_plus
    from .formula_utils import workbook_formula_dependencies
    from .models import validate_basic_types
except ImportError:  # flat-source dev fallback
    import ir_patch
    import excel_ir_plus
    from formula_utils import workbook_formula_dependencies
    from models import validate_basic_types

ROOT = Path(__file__).resolve().parent

EMBEDDED_SCHEMAS = {
    'ir.schema.json': {
        '$schema': 'https://json-schema.org/draft/2020-12/schema',
        'title': 'Excel IR MVP Workbook IR Schema',
        'type': 'object',
        'required': ['schema_version', 'workbook'],
        'properties': {
            'schema_version': {'type': 'string'},
            'workbook': {
                'type': 'object',
                'required': ['sheets', 'styles'],
                'properties': {
                    'styles': {'type': 'object'},
                    'patch_history': {'type': 'array'},
                    'sheets': {'type': 'array'}
                },
                'additionalProperties': True
            }
        },
        'additionalProperties': True
    },
    'patch.schema.json': {
        '$schema': 'https://json-schema.org/draft/2020-12/schema',
        'title': 'Excel IR MVP Patch Schema',
        'type': 'object',
        'required': ['actions'],
        'properties': {
            'name': {'type': 'string'},
            'actions': {'type': 'array', 'items': {'type': 'object', 'required': ['op'], 'additionalProperties': True}}
        },
        'additionalProperties': True
    }
}


def load(path: str):
    p = Path(path)
    if p.exists():
        return json.loads(p.read_text(encoding='utf-8'))
    p = ROOT / path
    if p.exists():
        return json.loads(p.read_text(encoding='utf-8'))
    if importlib_resources is not None and path in ('ir.schema.json', 'patch.schema.json', 'corpus_config.json'):
        try:
            return json.loads(importlib_resources.files(__package__).joinpath(path).read_text(encoding='utf-8'))
        except Exception:
            pass
    if path in EMBEDDED_SCHEMAS:
        return json.loads(json.dumps(EMBEDDED_SCHEMAS[path]))
    raise FileNotFoundError(path)


def validate_json_schema(data, schema_path: str):
    schema = load(schema_path)
    if jsonschema is not None:
        jsonschema.validate(data, schema)
        return []
    # Lightweight fallback: only required top-level keys.
    missing = [k for k in schema.get('required', []) if k not in data]
    return [{'level': 'error', 'message': f'missing top-level key {k}'} for k in missing]


def main():
    ap = argparse.ArgumentParser(description='Validate Excel IR / patch JSON')
    ap.add_argument('kind', choices=['ir', 'patch'])
    ap.add_argument('json_file')
    args = ap.parse_args()
    data = load(args.json_file)
    errors = []
    if args.kind == 'ir':
        errors.extend(validate_json_schema(data, 'ir.schema.json'))
        errors.extend(validate_basic_types(data))
        deps = workbook_formula_dependencies(data)
        result = {'ok': not errors, 'errors': errors, 'formula_dependencies': len(deps), 'formula_dependencies_sample': deps[:20]}
    else:
        errors.extend(validate_json_schema(data, 'patch.schema.json'))
        # If a default IR exists, also run semantic validation.
        default_ir = ROOT / 'complex_ir_v07.json'
        if default_ir.exists():
            errors.extend(ir_patch.validate_patch(json.loads(default_ir.read_text(encoding='utf-8')), data))
        result = {'ok': not any(e.get('level') == 'error' for e in errors), 'errors': errors}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if not result['ok']:
        raise SystemExit(1)

if __name__ == '__main__':
    main()
