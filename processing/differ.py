from __future__ import annotations

"""Structured workbook diff engine."""

from collections import Counter
from dataclasses import asdict
from typing import Any

from compare_excel_to_excel.config import get_sheet_rule
from compare_excel_to_excel.models.schemas import DiffItem, ExcelRow, ExcelWorkbook
from compare_excel_to_excel.processing.normalizer import normalize_formula


def diff_workbooks(
    baseline: ExcelWorkbook,
    revised: ExcelWorkbook,
    config: dict[str, Any],
) -> list[DiffItem]:
    diffs: list[DiffItem] = []
    counter = _DiffCounter()

    baseline_sheets = set(baseline.sheets)
    revised_sheets = set(revised.sheets)

    for sheet_name in sorted(revised_sheets - baseline_sheets):
        sheet = revised.sheets[sheet_name]
        diffs.append(counter.make(
            diff_type="sheet_added",
            severity="high",
            result_layer="confirmed_diff",
            sheet_name=sheet_name,
            business_key=None,
            row_before=None,
            row_after=None,
            column_name=None,
            old_value="",
            new_value=f"{len(sheet.rows)} rows",
            summary=f"revised 新增 sheet：{sheet_name}",
            rule_source="sheet_presence",
        ))

    for sheet_name in sorted(baseline_sheets - revised_sheets):
        sheet = baseline.sheets[sheet_name]
        diffs.append(counter.make(
            diff_type="sheet_deleted",
            severity="high",
            result_layer="confirmed_diff",
            sheet_name=sheet_name,
            business_key=None,
            row_before=None,
            row_after=None,
            column_name=None,
            old_value=f"{len(sheet.rows)} rows",
            new_value="",
            summary=f"revised 删除 sheet：{sheet_name}",
            rule_source="sheet_presence",
        ))

    for sheet_name in sorted(baseline_sheets & revised_sheets):
        diffs.extend(_diff_sheet(counter, sheet_name, baseline, revised, config))

    return diffs


def _diff_sheet(
    counter: "_DiffCounter",
    sheet_name: str,
    baseline: ExcelWorkbook,
    revised: ExcelWorkbook,
    config: dict[str, Any],
) -> list[DiffItem]:
    rule = get_sheet_rule(config, sheet_name)
    diffs: list[DiffItem] = []
    base_sheet = baseline.sheets[sheet_name]
    rev_sheet = revised.sheets[sheet_name]

    ignored = set(rule.ignore_columns)

    # 列存在性检测在“完整列集合”上进行（忽略列同样要报出，不能静默吞掉）。
    # ignore 只决定某列是否参与“单元格逐格比对”，不决定该列的“新增/删除”是否显示。
    base_all_set = set(base_sheet.columns)
    rev_all_set = set(rev_sheet.columns)

    for column in sorted(rev_all_set - base_all_set):
        diffs.append(_make_column_presence_diff(
            counter, sheet_name, column, added=True, is_ignored=column in ignored, config=config,
        ))

    for column in sorted(base_all_set - rev_all_set):
        diffs.append(_make_column_presence_diff(
            counter, sheet_name, column, added=False, is_ignored=column in ignored, config=config,
        ))

    # 单元格比对仅在“去掉忽略列”后的列集合上进行。
    base_columns = [col for col in base_sheet.columns if col not in ignored]
    rev_columns = [col for col in rev_sheet.columns if col not in ignored]
    rev_col_set = set(rev_columns)

    base_rows = _rows_by_key(base_sheet.rows)
    rev_rows = _rows_by_key(rev_sheet.rows)

    for duplicate_key in _duplicate_keys(base_sheet.rows) | _duplicate_keys(rev_sheet.rows):
        diffs.append(counter.make(
            diff_type="duplicate_key",
            severity="medium",
            result_layer="match_risk",
            sheet_name=sheet_name,
            business_key=duplicate_key,
            row_before=_first_row_number(base_rows.get(duplicate_key)),
            row_after=_first_row_number(rev_rows.get(duplicate_key)),
            column_name=None,
            old_value="",
            new_value="",
            summary=f"业务主键重复，属于行匹配风险而非内容差异：{duplicate_key}",
            need_human_review=True,
            rule_source="key_quality",
        ))

    base_keys = set(base_rows)
    rev_keys = set(rev_rows)

    for key in sorted(rev_keys - base_keys):
        row = rev_rows[key][0]
        diffs.append(counter.make(
            diff_type="row_added",
            severity="medium" if row.key_source == "business_key" else "low",
            result_layer="confirmed_diff" if row.key_source == "business_key" else "review_item",
            sheet_name=sheet_name,
            business_key=key,
            row_before=None,
            row_after=row.row_number,
            column_name=None,
            old_value="",
            new_value=_row_preview(row, rev_columns),
            summary=f"revised 新增行：{key}",
            need_human_review=row.key_source != "business_key",
            rule_source="row_presence",
            details={"key_source": row.key_source},
        ))

    for key in sorted(base_keys - rev_keys):
        row = base_rows[key][0]
        diffs.append(counter.make(
            diff_type="row_deleted",
            severity="medium" if row.key_source == "business_key" else "low",
            result_layer="confirmed_diff" if row.key_source == "business_key" else "review_item",
            sheet_name=sheet_name,
            business_key=key,
            row_before=row.row_number,
            row_after=None,
            column_name=None,
            old_value=_row_preview(row, base_columns),
            new_value="",
            summary=f"revised 删除行：{key}",
            need_human_review=row.key_source != "business_key",
            rule_source="row_presence",
            details={"key_source": row.key_source},
        ))

    shared_columns = [col for col in base_columns if col in rev_col_set]
    header_offset = (rev_sheet.header_row or 0) - (base_sheet.header_row or 0)
    for key in sorted(base_keys & rev_keys):
        base_row = base_rows[key][0]
        rev_row = rev_rows[key][0]

        row_delta = rev_row.row_number - base_row.row_number
        if (
            row_delta != 0
            and row_delta != header_offset
            and base_row.key_source == "business_key"
        ):
            diffs.append(counter.make(
                diff_type="row_moved",
                severity="low",
                result_layer="review_item",
                sheet_name=sheet_name,
                business_key=key,
                row_before=base_row.row_number,
                row_after=rev_row.row_number,
                column_name=None,
                old_value=str(base_row.row_number),
                new_value=str(rev_row.row_number),
                summary=f"同一业务主键行号变化：{base_row.row_number} -> {rev_row.row_number}",
                need_human_review=False,
                rule_source="row_position",
            ))

        for column in shared_columns:
            base_cell = base_row.cells.get(column)
            rev_cell = rev_row.cells.get(column)
            if not base_cell or not rev_cell:
                continue

            if base_cell.normalized_value != rev_cell.normalized_value:
                diffs.append(counter.make(
                    diff_type="cell_modified",
                    severity=_column_severity(column, config),
                    result_layer="confirmed_diff",
                    sheet_name=sheet_name,
                    business_key=key,
                    row_before=base_row.row_number,
                    row_after=rev_row.row_number,
                    column_name=column,
                    old_value=_stringify(base_cell.value),
                    new_value=_stringify(rev_cell.value),
                    summary=f"{column} 从“{_stringify(base_cell.value)}”变为“{_stringify(rev_cell.value)}”",
                    rule_source="cell_value",
                    details={
                        "old_coordinate": base_cell.coordinate,
                        "new_coordinate": rev_cell.coordinate,
                        "old_normalized": base_cell.normalized_value,
                        "new_normalized": rev_cell.normalized_value,
                    },
                ))

            if rule.compare_formulas:
                old_formula = normalize_formula(base_cell.formula)
                new_formula = normalize_formula(rev_cell.formula)
                # Old .xls files often do not expose formula text through xlrd.
                # Treat one-sided formula availability as parser capability mismatch,
                # unless the displayed value changed and was already reported above.
                if old_formula and new_formula and old_formula != new_formula:
                    diffs.append(counter.make(
                        diff_type="formula_modified",
                        severity=_column_severity(column, config),
                        result_layer="confirmed_diff",
                        sheet_name=sheet_name,
                        business_key=key,
                        row_before=base_row.row_number,
                        row_after=rev_row.row_number,
                        column_name=column,
                        old_value=base_cell.formula or "",
                        new_value=rev_cell.formula or "",
                        summary=f"{column} 公式变化",
                        rule_source="cell_formula",
                        details={
                            "old_coordinate": base_cell.coordinate,
                            "new_coordinate": rev_cell.coordinate,
                        },
                    ))

    return diffs


def diff_summary(diff_items: list[DiffItem]) -> dict[str, Any]:
    by_type = Counter(item.diff_type for item in diff_items)
    by_layer = Counter(item.result_layer for item in diff_items)
    by_severity = Counter(item.severity for item in diff_items)
    return {
        "total": len(diff_items),
        "by_type": dict(sorted(by_type.items())),
        "by_layer": dict(sorted(by_layer.items())),
        "by_severity": dict(sorted(by_severity.items())),
    }


def serialize_diff_items(diff_items: list[DiffItem]) -> list[dict[str, Any]]:
    return [asdict(item) for item in diff_items]


class _DiffCounter:
    def __init__(self) -> None:
        self.value = 0

    def make(self, **kwargs: Any) -> DiffItem:
        self.value += 1
        return DiffItem(diff_id=f"diff_{self.value:04d}", **kwargs)


def _make_column_presence_diff(
    counter: "_DiffCounter",
    sheet_name: str,
    column: str,
    *,
    added: bool,
    is_ignored: bool,
    config: dict[str, Any],
) -> DiffItem:
    """构造列“新增/删除”差异项。

    忽略列（profiler 判定为辅助列/工具生成列）同样要在结果中体现，
    但降级到 review_item / low，并注明“未参与单元格比对”，便于人工确认与否决。
    真实业务列则进入 confirmed_diff，按列名高危度定 severity。
    """
    action = "新增" if added else "删除"
    diff_type = "column_added" if added else "column_deleted"
    result_layer = "structure_change"
    severity = _column_severity(column, config)
    need_review = True
    if is_ignored:
        summary = (
            f"revised {action}列：{column}（疑似辅助列/工具生成列，"
            f"结构变化已留痕，默认不参与单元格逐格比对，请确认是否属于业务字段）"
        )
        rule_source = "column_presence_auxiliary"
    else:
        summary = f"revised {action}列：{column}（结构变化，需确认是否符合业务预期）"
        rule_source = "column_presence"

    return counter.make(
        diff_type=diff_type,
        severity=severity,
        result_layer=result_layer,
        sheet_name=sheet_name,
        business_key=None,
        row_before=None,
        row_after=None,
        column_name=column,
        old_value="" if added else column,
        new_value=column if added else "",
        summary=summary,
        need_human_review=need_review,
        rule_source=rule_source,
        details={"ignored_for_cell_compare": is_ignored},
    )


def _rows_by_key(rows: list[ExcelRow]) -> dict[str, list[ExcelRow]]:
    result: dict[str, list[ExcelRow]] = {}
    for row in rows:
        result.setdefault(row.key, []).append(row)
    return result


def _duplicate_keys(rows: list[ExcelRow]) -> set[str]:
    counts = Counter(row.key for row in rows)
    return {key for key, count in counts.items() if count > 1}


def _first_row_number(rows: list[ExcelRow] | None) -> int | None:
    if not rows:
        return None
    return rows[0].row_number


def _row_preview(row: ExcelRow, columns: list[str]) -> str:
    parts = []
    for column in columns[:8]:
        cell = row.cells.get(column)
        if cell and cell.normalized_value:
            parts.append(f"{column}={_stringify(cell.value)}")
    return "；".join(parts)


def _column_severity(column_name: str | None, config: dict[str, Any]) -> str:
    if not column_name:
        return "medium"
    lowered = column_name.lower()
    keywords = config.get("high_risk_keywords", []) or []
    return "high" if any(str(keyword).lower() in lowered for keyword in keywords) else "medium"


def _stringify(value: Any) -> str:
    if value is None:
        return ""
    return str(value)
