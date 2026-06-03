from __future__ import annotations

"""Value normalization before comparison."""

import re
import unicodedata
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any

from compare_excel_to_excel.models.schemas import SheetRule


_SPACE_RE = re.compile(r"\s+")


def normalize_value(value: Any, column_name: str, rule: SheetRule) -> str:
    """Normalize one cell value according to configured column rules."""

    if value is None:
        return ""

    if isinstance(value, datetime):
        return value.date().isoformat() if column_name in rule.date_columns else value.isoformat(sep=" ")

    if isinstance(value, date):
        return value.isoformat()

    text = str(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = "\n".join(_SPACE_RE.sub(" ", part).strip() for part in text.split("\n"))
    text = text.strip()

    if text in set(rule.blank_values):
        return ""

    if column_name in rule.numeric_columns:
        return _normalize_number(text, rule.numeric_columns[column_name])

    if column_name in rule.date_columns:
        parsed = _parse_date(text)
        return parsed.isoformat() if parsed else text

    if column_name in rule.case_insensitive_columns:
        return text.lower()

    return text


def normalize_formula(formula: str | None) -> str:
    if not formula:
        return ""
    return _SPACE_RE.sub("", str(formula)).upper()


def _normalize_number(text: str, precision: int) -> str:
    cleaned = text.replace(",", "")
    if cleaned.endswith("%"):
        cleaned = cleaned[:-1]
    try:
        value = Decimal(cleaned)
    except (InvalidOperation, ValueError):
        return text

    quant = Decimal("1") if precision <= 0 else Decimal("1").scaleb(-precision)
    return str(value.quantize(quant, rounding=ROUND_HALF_UP))


def _parse_date(text: str) -> date | None:
    candidates = [
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y.%m.%d",
        "%Y年%m月%d日",
        "%Y年%m月",
        "%Y/%m",
        "%Y-%m",
    ]
    for fmt in candidates:
        try:
            dt = datetime.strptime(text, fmt)
            return dt.date()
        except ValueError:
            continue
    return None

