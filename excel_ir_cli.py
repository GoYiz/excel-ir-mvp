from __future__ import annotations

import sys
from pathlib import Path

_SRC = Path(__file__).resolve().parent / "src"
if _SRC.exists():
    sys.path.insert(0, str(_SRC))

from excel_ir_mvp.excel_ir_cli import main

if __name__ == "__main__":
    main()
