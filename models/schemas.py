from __future__ import annotations

"""Shared data models for structured Excel comparison."""

from dataclasses import dataclass, field
from typing import Any


@dataclass
class SheetRule:
    """Per-sheet comparison rule loaded from config."""

    sheet_name: str
    header_row: int | str = "auto"
    key_columns: list[str] = field(default_factory=list)
    ignore_columns: list[str] = field(default_factory=list)
    numeric_columns: dict[str, int] = field(default_factory=dict)
    date_columns: list[str] = field(default_factory=list)
    case_insensitive_columns: list[str] = field(default_factory=list)
    compare_formulas: bool = True
    blank_values: list[str] = field(default_factory=lambda: ["", "NULL", "N/A", "NA", "-"])


@dataclass
class ExcelCell:
    """A parsed cell keyed by logical column name."""

    column_name: str
    coordinate: str
    value: Any
    formula: str | None = None
    normalized_value: str = ""


@dataclass
class ExcelRow:
    """A parsed business row."""

    row_number: int
    key: str
    key_source: str
    cells: dict[str, ExcelCell]


@dataclass
class ExcelSheet:
    """Structured sheet data."""

    name: str
    header_row: int
    columns: list[str]
    rows: list[ExcelRow]
    hidden_columns: list[str] = field(default_factory=list)
    merged_ranges: list[str] = field(default_factory=list)


@dataclass
class ExcelWorkbook:
    """Structured workbook data."""

    path: str
    sheets: dict[str, ExcelSheet]


@dataclass
class DiffItem:
    """One auditable difference item."""

    diff_id: str
    diff_type: str
    severity: str
    result_layer: str
    sheet_name: str
    business_key: str | None
    row_before: int | None
    row_after: int | None
    column_name: str | None
    old_value: str
    new_value: str
    summary: str
    need_human_review: bool = False
    rule_source: str = ""
    details: dict[str, Any] = field(default_factory=dict)
