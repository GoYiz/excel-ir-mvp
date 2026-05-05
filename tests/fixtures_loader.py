from __future__ import annotations

import json
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = Path(__file__).resolve().parent / 'fixtures'


def fixture_path(name: str) -> Path:
    return FIXTURES / name


def load_json_fixture(name: str) -> dict[str, Any]:
    return json.loads(fixture_path(name).read_text(encoding='utf-8'))
