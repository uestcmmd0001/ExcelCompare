from __future__ import annotations

"""JSON export helpers."""

import json
from dataclasses import asdict, is_dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any


def save_json(path: str | Path, payload: Any) -> Path:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=_default)
    return output_path


def save_yaml(path: str | Path, payload: Any) -> Path:
    try:
        import yaml
    except ImportError as exc:  # pragma: no cover - dependency failure path
        raise RuntimeError("PyYAML is required to export YAML") from exc

    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as f:
        yaml.safe_dump(payload, f, allow_unicode=True, sort_keys=False)
    return output_path


def _default(value: Any) -> Any:
    if is_dataclass(value):
        return asdict(value)
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value)
