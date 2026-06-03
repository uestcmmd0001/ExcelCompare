from __future__ import annotations

"""Configuration loader for Excel comparison."""

from pathlib import Path
from typing import Any

from compare_excel_to_excel.models.schemas import SheetRule
from compare_excel_to_excel.profilers.workbook_profiler import infer_header_row


def load_config(config_path: str | Path | None) -> dict[str, Any]:
    if not config_path:
        return {}

    path = Path(config_path)
    if not path.exists():
        return {}

    try:
        import yaml
    except ImportError as exc:  # pragma: no cover - dependency failure path
        raise RuntimeError("PyYAML is required to load config.yaml") from exc

    with path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def get_sheet_rule(config: dict[str, Any], sheet_name: str, role: str | None = None) -> SheetRule:
    default_cfg = config.get("default", {}) or {}
    sheet_cfg = (config.get("sheets", {}) or {}).get(sheet_name, {}) or {}

    def _get(name: str, default: Any = None) -> Any:
        return sheet_cfg.get(name, default_cfg.get(name, default))

    header_row = _get("header_row", "auto")
    if role in ("baseline", "revised"):
        header_row = _get(f"header_row_{role}", header_row)
    key_columns_cfg = _get("key_columns", [])
    if key_columns_cfg == "auto":
        key_columns = ["auto"]
    else:
        key_columns = list(key_columns_cfg or [])

    return SheetRule(
        sheet_name=sheet_name,
        header_row=header_row,
        key_columns=key_columns,
        ignore_columns=list(_get("ignore_columns", []) or []),
        numeric_columns=dict(_get("numeric_columns", {}) or {}),
        date_columns=list(_get("date_columns", []) or []),
        case_insensitive_columns=list(_get("case_insensitive_columns", []) or []),
        compare_formulas=bool(_get("compare_formulas", True)),
        blank_values=list(_get("blank_values", ["", "NULL", "N/A", "NA", "-"]) or []),
    )


def should_compare_sheet(config: dict[str, Any], sheet_name: str) -> bool:
    if "include_sheets" in config and config.get("include_sheets") is not None:
        include_sheets = config.get("include_sheets") or []
        return sheet_name in include_sheets

    exclude_sheets = set(config.get("exclude_sheets", []) or [])
    return sheet_name not in exclude_sheets


def resolve_auto_rule(rule: SheetRule, preview_rows: list[list[Any]], config: dict[str, Any] | None = None) -> SheetRule:
    if rule.header_row == "auto":
        rule.header_row = infer_header_row(preview_rows, config or {})
    else:
        rule.header_row = int(rule.header_row)
    if rule.key_columns == ["auto"]:
        rule.key_columns = []
    return rule
