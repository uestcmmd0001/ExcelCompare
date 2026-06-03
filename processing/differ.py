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
    sheet_pairs = _executable_sheet_pairs(config, baseline_sheets, revised_sheets)
    paired_baseline_sheets = {pair[0] for pair in sheet_pairs}
    paired_revised_sheets = {pair[1] for pair in sheet_pairs}

    for sheet_name in sorted(revised_sheets - baseline_sheets - paired_revised_sheets):
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

    for sheet_name in sorted(baseline_sheets - revised_sheets - paired_baseline_sheets):
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

    for baseline_sheet, revised_sheet in sheet_pairs:
        diffs.extend(_diff_sheet(counter, baseline_sheet, revised_sheet, baseline, revised, config))

    return diffs


def _diff_sheet(
    counter: "_DiffCounter",
    baseline_sheet_name: str,
    revised_sheet_name: str,
    baseline: ExcelWorkbook,
    revised: ExcelWorkbook,
    config: dict[str, Any],
) -> list[DiffItem]:
    rule = get_sheet_rule(config, baseline_sheet_name)
    diffs: list[DiffItem] = []
    base_sheet = baseline.sheets[baseline_sheet_name]
    rev_sheet = revised.sheets[revised_sheet_name]
    sheet_label = _sheet_label(baseline_sheet_name, revised_sheet_name)

    ignored = set(rule.ignore_columns)
    structure_only = set(rule.structure_only_columns)
    skip_cell_compare = ignored | structure_only
    column_mapping = _effective_column_mapping(rule.column_mapping, base_sheet.columns, rev_sheet.columns)
    mapped_base_columns = set(column_mapping)
    mapped_rev_columns = set(column_mapping.values())

    # 列存在性检测在完整列集合上进行，任何新增/删除列都不能静默吞掉。
    # ignore / structure-only 只决定是否参与逐格比对，不决定结构变化是否显示。
    base_all_set = set(base_sheet.columns)
    rev_all_set = set(rev_sheet.columns)

    for column in sorted(rev_all_set - mapped_rev_columns):
        diffs.append(_make_column_presence_diff(
            counter,
            sheet_label,
            column,
            added=True,
            is_ignored=column in ignored,
            is_structure_only=column in structure_only,
            config=config,
        ))

    for column in sorted(base_all_set - mapped_base_columns):
        diffs.append(_make_column_presence_diff(
            counter,
            sheet_label,
            column,
            added=False,
            is_ignored=column in ignored,
            is_structure_only=column in structure_only,
            config=config,
        ))

    # 单元格比对仅在“去掉真正忽略列 + 只作结构留痕列”后的列集合上进行。
    base_columns = [col for col in base_sheet.columns if col not in skip_cell_compare]
    rev_preview_columns = [col for col in rev_sheet.columns if col not in skip_cell_compare]
    comparable_columns = [
        base_col for base_col in base_columns
        if base_col in column_mapping and column_mapping[base_col] not in skip_cell_compare
    ]

    base_rows = _rows_by_key(base_sheet.rows)
    rev_rows = _rows_by_key(rev_sheet.rows)

    duplicate_keys = _duplicate_keys(base_sheet.rows) | _duplicate_keys(rev_sheet.rows)
    for duplicate_key in duplicate_keys:
        diffs.append(counter.make(
            diff_type="duplicate_key",
            severity="medium",
            result_layer="match_risk",
            sheet_name=sheet_label,
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
        if key in duplicate_keys:
            continue
        row = rev_rows[key][0]
        diffs.append(counter.make(
            diff_type="row_added",
            severity="medium" if row.key_source == "business_key" else "low",
            result_layer="confirmed_diff" if row.key_source == "business_key" else "review_item",
            sheet_name=sheet_label,
            business_key=key,
            row_before=None,
            row_after=row.row_number,
            column_name=None,
            old_value="",
            new_value=_row_preview(row, rev_preview_columns),
            summary=f"revised 新增行：{key}",
            need_human_review=row.key_source != "business_key",
            rule_source="row_presence",
            details={"key_source": row.key_source},
        ))

    for key in sorted(base_keys - rev_keys):
        if key in duplicate_keys:
            continue
        row = base_rows[key][0]
        diffs.append(counter.make(
            diff_type="row_deleted",
            severity="medium" if row.key_source == "business_key" else "low",
            result_layer="confirmed_diff" if row.key_source == "business_key" else "review_item",
            sheet_name=sheet_label,
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

    header_offset = (rev_sheet.header_row or 0) - (base_sheet.header_row or 0)
    for key in sorted(base_keys & rev_keys):
        if key in duplicate_keys:
            diffs.append(counter.make(
                diff_type="duplicate_key_rows_skipped",
                severity="medium",
                result_layer="match_risk",
                sheet_name=sheet_label,
                business_key=key,
                row_before=_first_row_number(base_rows.get(key)),
                row_after=_first_row_number(rev_rows.get(key)),
                column_name=None,
                old_value=_rows_preview(base_rows.get(key) or [], base_columns),
                new_value=_rows_preview(rev_rows.get(key) or [], rev_preview_columns),
                summary=f"业务主键重复，跳过该 key 下的逐格确认差异：{key}",
                need_human_review=True,
                rule_source="key_quality",
            ))
            continue
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
                sheet_name=sheet_label,
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

        for column in comparable_columns:
            revised_column = column_mapping[column]
            base_cell = base_row.cells.get(column)
            rev_cell = rev_row.cells.get(revised_column)
            if not base_cell or not rev_cell:
                continue

            old_formula = normalize_formula(base_cell.formula)
            new_formula = normalize_formula(rev_cell.formula)
            formulas_changed = bool(old_formula and new_formula and old_formula != new_formula)
            values_are_formula_text = bool(
                base_cell.formula
                and rev_cell.formula
                and base_cell.value == base_cell.formula
                and rev_cell.value == rev_cell.formula
            )

            if base_cell.normalized_value != rev_cell.normalized_value and not (
                formulas_changed and values_are_formula_text
            ):
                column_label = _column_label(column, revised_column)
                diffs.append(counter.make(
                    diff_type="cell_modified",
                    severity=_column_severity(column, config),
                    result_layer="confirmed_diff",
                    sheet_name=sheet_label,
                    business_key=key,
                    row_before=base_row.row_number,
                    row_after=rev_row.row_number,
                    column_name=column_label,
                    old_value=_stringify(base_cell.value),
                    new_value=_stringify(rev_cell.value),
                    summary=f"{column_label} 从“{_stringify(base_cell.value)}”变为“{_stringify(rev_cell.value)}”",
                    rule_source="cell_value",
                    details={
                        "old_coordinate": base_cell.coordinate,
                        "new_coordinate": rev_cell.coordinate,
                        "baseline_sheet": baseline_sheet_name,
                        "revised_sheet": revised_sheet_name,
                        "canonical_column": column,
                        "baseline_column": column,
                        "revised_column": revised_column,
                        "old_normalized": base_cell.normalized_value,
                        "new_normalized": rev_cell.normalized_value,
                    },
                ))

            if rule.compare_formulas:
                # Old .xls files often do not expose formula text through xlrd.
                # Treat one-sided formula availability as parser capability mismatch,
                # unless the displayed value changed and was already reported above.
                if formulas_changed:
                    column_label = _column_label(column, revised_column)
                    diffs.append(counter.make(
                        diff_type="formula_modified",
                        severity=_column_severity(column, config),
                        result_layer="confirmed_diff",
                        sheet_name=sheet_label,
                        business_key=key,
                        row_before=base_row.row_number,
                        row_after=rev_row.row_number,
                        column_name=column_label,
                        old_value=base_cell.formula or "",
                        new_value=rev_cell.formula or "",
                        summary=f"{column_label} 公式变化",
                        rule_source="cell_formula",
                        details={
                            "old_coordinate": base_cell.coordinate,
                            "new_coordinate": rev_cell.coordinate,
                            "baseline_sheet": baseline_sheet_name,
                            "revised_sheet": revised_sheet_name,
                            "canonical_column": column,
                            "baseline_column": column,
                            "revised_column": revised_column,
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


def _executable_sheet_pairs(
    config: dict[str, Any],
    baseline_sheets: set[str],
    revised_sheets: set[str],
) -> list[tuple[str, str]]:
    configured_pairs = config.get("sheet_pairs")
    if configured_pairs is not None:
        result: list[tuple[str, str]] = []
        seen: set[tuple[str, str]] = set()
        for pair in configured_pairs or []:
            if not isinstance(pair, dict):
                continue
            baseline_sheet = str(pair.get("baseline_sheet") or "")
            revised_sheet = str(pair.get("revised_sheet") or "")
            item = (baseline_sheet, revised_sheet)
            if baseline_sheet in baseline_sheets and revised_sheet in revised_sheets and item not in seen:
                result.append(item)
                seen.add(item)
        return result

    common_sheets = baseline_sheets & revised_sheets
    if "include_sheets" in config and config.get("include_sheets") is not None:
        return [(sheet_name, sheet_name) for sheet_name in sorted(set(config.get("include_sheets") or []) & common_sheets)]
    exclude_sheets = set(config.get("exclude_sheets", []) or [])
    return [(sheet_name, sheet_name) for sheet_name in sorted(common_sheets - exclude_sheets)]


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
    is_structure_only: bool,
    config: dict[str, Any],
) -> DiffItem:
    """构造列“新增/删除”差异项。

    真正忽略列和 structure-only 列同样要在结果中体现，并注明未参与单元格比对，
    便于人工确认与否决。真实业务列按普通结构变化输出。
    """
    action = "新增" if added else "删除"
    diff_type = "column_added" if added else "column_deleted"
    result_layer = "structure_change"
    severity = _column_severity(column, config)
    need_review = True
    if is_ignored:
        summary = (
            f"revised {action}列：{column}（配置为完全忽略列，但列存在性变化仍留痕，"
            f"请确认该忽略规则是否符合业务预期）"
        )
        rule_source = "column_presence_ignored"
    elif is_structure_only:
        summary = (
            f"revised {action}列：{column}（疑似辅助列/工具生成列，"
            f"结构变化已留痕，默认不参与单元格逐格比对，请确认是否属于业务字段）"
        )
        rule_source = "column_presence_structure_only"
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
        details={
            "ignored_for_cell_compare": is_ignored or is_structure_only,
            "ignore_policy": "ignored" if is_ignored else "structure_only" if is_structure_only else "compared",
        },
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


def _rows_preview(rows: list[ExcelRow], columns: list[str]) -> str:
    return " || ".join(_row_preview(row, columns) for row in rows[:3])


def _effective_column_mapping(
    configured_mapping: dict[str, str],
    base_columns: list[str],
    rev_columns: list[str],
) -> dict[str, str]:
    base_set = set(base_columns)
    rev_set = set(rev_columns)
    mapping: dict[str, str] = {}
    for base_col, rev_col in configured_mapping.items():
        if base_col in base_set and rev_col in rev_set:
            mapping[base_col] = rev_col
    for column in base_columns:
        if column in mapping:
            continue
        if column in rev_set:
            mapping[column] = column
    return mapping


def _column_label(base_column: str, revised_column: str) -> str:
    if base_column == revised_column:
        return base_column
    return f"{base_column} -> {revised_column}"


def _sheet_label(baseline_sheet: str, revised_sheet: str) -> str:
    if baseline_sheet == revised_sheet:
        return baseline_sheet
    return f"{baseline_sheet} -> {revised_sheet}"


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
