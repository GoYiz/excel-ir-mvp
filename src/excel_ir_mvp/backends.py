from __future__ import annotations

import importlib
from dataclasses import dataclass
from typing import Any, Dict, List


class BackendUnavailableError(RuntimeError):
    """Raised when a requested Excel backend cannot be imported."""


@dataclass(frozen=True)
class ExcelBackend:
    name: str
    module_name: str
    description: str
    optional: bool = False

    def available(self) -> bool:
        try:
            importlib.import_module(self.module_name)
            return True
        except Exception:
            return False

    def module(self) -> Any:
        try:
            return importlib.import_module(self.module_name)
        except Exception as exc:
            raise BackendUnavailableError(
                f"Excel backend '{self.name}' is not available. Install/import '{self.module_name}' first."
            ) from exc


_BACKENDS: Dict[str, ExcelBackend] = {
    "openpyxl": ExcelBackend("openpyxl", "openpyxl", "Default full-fidelity Python engine"),
    "wolfxl": ExcelBackend("wolfxl", "wolfxl", "Optional Rust-backed openpyxl-compatible engine", optional=True),
}


def available_engines() -> List[str]:
    return [name for name, backend in _BACKENDS.items() if backend.available()]


def engine_status() -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    for name, backend in _BACKENDS.items():
        version = None
        available = backend.available()
        if available:
            try:
                mod = importlib.import_module(backend.module_name)
                version = getattr(mod, "__version__", None)
            except Exception:
                available = False
        out[name] = {
            "available": available,
            "module": backend.module_name,
            "version": version,
            "optional": backend.optional,
            "description": backend.description,
        }
    return out


def resolve_engine(engine: str | None = None) -> ExcelBackend:
    choice = (engine or "openpyxl").lower()
    if choice == "auto":
        choice = "wolfxl" if _BACKENDS["wolfxl"].available() else "openpyxl"
    if choice not in _BACKENDS:
        raise ValueError(f"unknown engine '{engine}'. Expected one of: auto, {', '.join(sorted(_BACKENDS))}")
    backend = _BACKENDS[choice]
    if not backend.available():
        raise BackendUnavailableError(
            f"Excel backend '{choice}' is not available. Use --engine openpyxl or install {backend.module_name}."
        )
    return backend


def load_workbook(engine: str | None = None, *args: Any, **kwargs: Any) -> Any:
    backend = resolve_engine(engine)
    return backend.module().load_workbook(*args, **kwargs)


def workbook_class(engine: str | None = None) -> Any:
    backend = resolve_engine(engine)
    return getattr(backend.module(), "Workbook")
