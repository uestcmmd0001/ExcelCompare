from __future__ import annotations

"""Validate automatic comparison plans and convert them into executable config."""

import unicodedata
from statistics import mean
from typing import Any

from compare_excel_to_excel.planning.plan_schema import (
    ComparisonPlan,
    PlanValidationResult,
    SheetPlan,
    SheetValidation,
)


def validate_plan(
    plan: ComparisonPlan,
    profile: dict[str, Any],
    config: dict[str, Any],
) -> PlanValidationResult:
    baseline_sheets = {sheet["name"]: sheet for sheet in profile.get("baseline", {}).get("sheets", [])}
    revised_sheets = {sheet["name"]: sheet for sheet in profile.get("revised", {}).get("sheets", [])}
    baseline_sheet_names = _names_by_normalized_key(baseline_sheets)
    revised_sheet_names = _names_by_normalized_key(revised_sheets)

    sheet_validations: list[SheetValidation] = []
    executable_sheets: dict[str, Any] = {}
    include_sheets: list[str] = []
    risks: list[str] = []

    for sheet_plan in plan.primary_plan.sheet_pairs:
        original_baseline_name = sheet_plan.baseline_sheet
        original_revised_name = sheet_plan.revised_sheet
        baseline_name = _resolve_name(sheet_plan.baseline_sheet, baseline_sheets, baseline_sheet_names)
        revised_name = _resolve_name(sheet_plan.revised_sheet, revised_sheets, revised_sheet_names)
        if baseline_name:
            sheet_plan.baseline_sheet = baseline_name
        if revised_name:
            sheet_plan.revised_sheet = revised_name

        baseline_sheet = baseline_sheets.get(sheet_plan.baseline_sheet)
        revised_sheet = revised_sheets.get(sheet_plan.revised_sheet)
        if not baseline_sheet or not revised_sheet:
            risks.append(
                f"planned sheet pair `{original_baseline_name}` / `{original_revised_name}` is not available"
            )
            continue
        if original_baseline_name != sheet_plan.baseline_sheet or original_revised_name != sheet_plan.revised_sheet:
            risks.append(
                f"normalized planned sheet pair `{original_baseline_name}` / `{original_revised_name}` "
                f"to `{sheet_plan.baseline_sheet}` / `{sheet_plan.revised_sheet}`"
            )

        validation = _validate_sheet_plan(sheet_plan, baseline_sheet, revised_sheet, config)
        sheet_validations.append(validation)
        risks.extend(f"{sheet_plan.baseline_sheet}: {risk}" for risk in validation.risks)

        # The current parser/diff engine supports same-name sheet execution.
        # Cross-name matching is still reported as plan risk until the parser grows aliases.
        if sheet_plan.baseline_sheet != sheet_plan.revised_sheet:
            continue
        if _should_skip_execution(validation, config):
            risks.append(
                f"{sheet_plan.baseline_sheet}: skipped from executable diff because plan validation did not pass"
            )
            continue

        include_sheets.append(sheet_plan.baseline_sheet)
        executable_sheets[sheet_plan.baseline_sheet] = {
            "header_row": sheet_plan.header_row_baseline,
            "header_row_baseline": sheet_plan.header_row_baseline,
            "header_row_revised": sheet_plan.header_row_revised,
            "key_columns": validation.selected_key_columns,
            "ignore_columns": list(sheet_plan.ignore_columns),
            "numeric_columns": dict(sheet_plan.numeric_columns),
            "date_columns": list(sheet_plan.date_columns),
            "case_insensitive_columns": [],
            "compare_formulas": sheet_plan.compare_formulas,
            "plan_confidence": validation.overall_confidence,
            "plan_risks": validation.risks,
        }

    effective_config = _base_effective_config(config)
    if include_sheets:
        effective_config["include_sheets"] = include_sheets
    effective_config["sheets"] = executable_sheets

    overall = round(mean([item.overall_confidence for item in sheet_validations]), 4) if sheet_validations else 0.0
    return PlanValidationResult(
        overall_confidence=overall,
        sheet_validations=sheet_validations,
        risks=risks,
        effective_config=effective_config,
    )


def _validate_sheet_plan(
    sheet_plan: SheetPlan,
    baseline_sheet: dict[str, Any],
    revised_sheet: dict[str, Any],
    config: dict[str, Any],
) -> SheetValidation:
    risks: list[str] = []
    notes: list[str] = []

    sheet_match_score = 1.0 if sheet_plan.baseline_sheet == sheet_plan.revised_sheet else 0.55
    if sheet_match_score < 1.0:
        risks.append("sheet names differ; current executable diff cannot alias sheets automatically")

    header_score = _header_score(sheet_plan, baseline_sheet, revised_sheet)
    if header_score < 0.8:
        risks.append("header row differs or is weakly detected")

    base_columns = set(baseline_sheet.get("columns", []))
    rev_columns = set(revised_sheet.get("columns", []))
    base_column_names = _names_by_normalized_key({column: column for column in base_columns})
    rev_column_names = _names_by_normalized_key({column: column for column in rev_columns})
    mapped_columns = _resolve_column_mapping(sheet_plan, base_columns, rev_columns, base_column_names, rev_column_names)
    sheet_plan.column_mapping = dict(mapped_columns)
    column_mapping_score = round(len(mapped_columns) / max(len(base_columns | rev_columns), 1), 4)
    if column_mapping_score < 0.6:
        risks.append("low column overlap between baseline and revised")

    duplicate_header_ratio = max(
        float(baseline_sheet.get("duplicate_header_ratio", 0)),
        float(revised_sheet.get("duplicate_header_ratio", 0)),
    )
    if duplicate_header_ratio > float(config.get("auto_include_max_duplicate_header_ratio", 0.35)):
        risks.append(
            f"high duplicate-header ratio={duplicate_header_ratio:.2f}; sheet may be a matrix/reference sheet"
        )

    key_columns = []
    for column in sheet_plan.key_columns:
        resolved = _resolve_name(column, {name: name for name in mapped_columns}, base_column_names)
        if resolved and resolved in mapped_columns:
            key_columns.append(resolved)
    sheet_plan.key_columns = key_columns
    key_coverage_score, key_uniqueness_score = _key_scores(key_columns, baseline_sheet, revised_sheet)
    selected_key_columns = key_columns
    min_key_uniqueness = float(config.get("planner_min_key_uniqueness", 0.7))
    min_key_coverage = float(config.get("planner_min_key_coverage", 0.6))

    better_key, better_coverage, better_uniqueness = _select_better_key_candidate(
        key_columns,
        baseline_sheet,
        revised_sheet,
        mapped_columns,
        min_key_coverage,
        min_key_uniqueness,
    )
    current_key_is_weak = key_coverage_score < min_key_coverage or key_uniqueness_score < min_key_uniqueness
    if better_key and current_key_is_weak and (better_uniqueness, better_coverage, -len(better_key)) > (
        key_uniqueness_score,
        key_coverage_score,
        -len(selected_key_columns or []),
    ):
        selected_key_columns = better_key
        sheet_plan.key_columns = better_key
        notes.append(
            f"auto-selected stronger key: {' + '.join(better_key)} "
            f"(coverage={better_coverage:.2f}, uniqueness={better_uniqueness:.2f})"
        )
        key_coverage_score = better_coverage
        key_uniqueness_score = better_uniqueness

    if not selected_key_columns:
        risks.append("no business key selected; row-number fallback may produce more review items")
    elif key_coverage_score < min_key_coverage or key_uniqueness_score < min_key_uniqueness:
        risks.append(
            f"selected key is weak: coverage={key_coverage_score:.2f}, uniqueness={key_uniqueness_score:.2f}"
        )
        notes.append("weak keys are still used but related row results should be reviewed")

    type_consistency_score = _type_consistency_score(mapped_columns, baseline_sheet, revised_sheet)
    if type_consistency_score < 0.75:
        risks.append("mapped columns have inconsistent observed value types")

    added_columns = sorted(rev_columns - set(mapped_columns.values()))
    deleted_columns = sorted(base_columns - set(mapped_columns))
    if added_columns:
        notes.append(f"revised has {len(added_columns)} added columns; they will be reported as structure changes")
    if deleted_columns:
        notes.append(f"revised has {len(deleted_columns)} deleted columns; they will be reported as structure changes")

    overall = round(
        sheet_match_score * 0.18
        + header_score * 0.12
        + key_coverage_score * 0.18
        + key_uniqueness_score * 0.22
        + column_mapping_score * 0.18
        + type_consistency_score * 0.12,
        4,
    )

    return SheetValidation(
        baseline_sheet=sheet_plan.baseline_sheet,
        revised_sheet=sheet_plan.revised_sheet,
        sheet_match_score=round(sheet_match_score, 4),
        header_score=round(header_score, 4),
        key_coverage_score=round(key_coverage_score, 4),
        key_uniqueness_score=round(key_uniqueness_score, 4),
        column_mapping_score=round(column_mapping_score, 4),
        type_consistency_score=round(type_consistency_score, 4),
        overall_confidence=overall,
        selected_key_columns=selected_key_columns,
        risks=risks,
        notes=notes,
    )


def _header_score(sheet_plan: SheetPlan, baseline_sheet: dict[str, Any], revised_sheet: dict[str, Any]) -> float:
    base_detected = int(baseline_sheet.get("detected_header_row") or 1)
    rev_detected = int(revised_sheet.get("detected_header_row") or 1)
    if sheet_plan.header_row_baseline == base_detected and sheet_plan.header_row_revised == rev_detected:
        return 1.0
    if sheet_plan.header_row_baseline == sheet_plan.header_row_revised:
        return 0.85
    return 0.6


def _key_scores(
    key_columns: list[str],
    baseline_sheet: dict[str, Any],
    revised_sheet: dict[str, Any],
) -> tuple[float, float]:
    if not key_columns:
        return 0.25, 0.2

    base_profiles = _profiles_by_name(baseline_sheet)
    rev_profiles = _profiles_by_name(revised_sheet)
    coverages = []
    uniqueness = []
    for column in key_columns:
        for sheet, profiles in ((baseline_sheet, base_profiles), (revised_sheet, rev_profiles)):
            profile = profiles.get(column)
            if not profile:
                coverages.append(0.0)
                uniqueness.append(0.0)
                continue
            data_rows = max(int(sheet.get("data_rows") or 0), 1)
            coverages.append(min(float(profile.get("non_blank_count", 0)) / data_rows, 1.0))
            uniqueness.append(float(profile.get("unique_ratio", 0)))

    return min(coverages or [0.0]), min(uniqueness or [0.0])


def _select_better_key_candidate(
    current_columns: list[str],
    baseline_sheet: dict[str, Any],
    revised_sheet: dict[str, Any],
    mapped_columns: dict[str, str],
    min_coverage: float,
    min_uniqueness: float,
) -> tuple[list[str], float, float]:
    candidates = []
    seen: set[tuple[str, ...]] = set()
    for item in baseline_sheet.get("key_candidates", []) + revised_sheet.get("key_candidates", []):
        columns = tuple(column for column in item.get("columns", []) if column in mapped_columns)
        if not columns or columns in seen or len(columns) != len(item.get("columns", [])):
            continue
        if _candidate_has_measure_like_column(columns, baseline_sheet, revised_sheet):
            continue
        seen.add(columns)
        coverage, uniqueness = _sample_pair_key_quality(list(columns), baseline_sheet, revised_sheet)
        candidates.append((list(columns), coverage, uniqueness))

    for columns in _generated_composite_candidates(mapped_columns, baseline_sheet, revised_sheet):
        key = tuple(columns)
        if key in seen:
            continue
        seen.add(key)
        coverage, uniqueness = _sample_pair_key_quality(columns, baseline_sheet, revised_sheet)
        candidates.append((columns, coverage, uniqueness))

    if current_columns:
        current_coverage, current_uniqueness = _sample_pair_key_quality(current_columns, baseline_sheet, revised_sheet)
        candidates.append((current_columns, current_coverage, current_uniqueness))

    if not candidates:
        return [], 0.0, 0.0

    candidates.sort(
        key=lambda item: (
            item[1] >= min_coverage,
            item[2] >= min_uniqueness,
            item[2],
            item[1],
            -len(item[0]),
        ),
        reverse=True,
    )
    best_columns, best_coverage, best_uniqueness = candidates[0]
    if best_uniqueness >= min_uniqueness and best_coverage >= min_coverage:
        return best_columns, best_coverage, best_uniqueness
    return best_columns, best_coverage, best_uniqueness


def _sample_pair_key_quality(
    columns: list[str],
    baseline_sheet: dict[str, Any],
    revised_sheet: dict[str, Any],
) -> tuple[float, float]:
    base_coverage, base_unique = _sample_key_quality(columns, baseline_sheet.get("row_samples", []) or [])
    rev_coverage, rev_unique = _sample_key_quality(columns, revised_sheet.get("row_samples", []) or [])
    return min(base_coverage, rev_coverage), min(base_unique, rev_unique)


def _sample_key_quality(columns: list[str], samples: list[dict[str, str]]) -> tuple[float, float]:
    if not columns or not samples:
        return 0.0, 0.0
    keys = []
    non_blank = 0
    for row in samples:
        parts = [str(row.get(column, "")).strip() for column in columns]
        if all(parts):
            non_blank += 1
            keys.append("|".join(parts))
    coverage = non_blank / max(len(samples), 1)
    uniqueness = len(set(keys)) / max(len(keys), 1) if keys else 0.0
    return coverage, uniqueness


def _generated_composite_candidates(
    mapped_columns: dict[str, str],
    baseline_sheet: dict[str, Any],
    revised_sheet: dict[str, Any],
) -> list[list[str]]:
    profiles = _profiles_by_name(baseline_sheet) | _profiles_by_name(revised_sheet)
    text_columns = [
        column for column in mapped_columns
        if profiles.get(column, {}).get("non_blank_count", 0) > 0
        and profiles.get(column, {}).get("numeric_ratio", 0) <= 0.95
    ]
    text_columns.sort(
        key=lambda column: (
            _sample_pair_key_quality([column], baseline_sheet, revised_sheet)[0],
            _sample_pair_key_quality([column], baseline_sheet, revised_sheet)[1],
        ),
        reverse=True,
    )
    text_columns = text_columns[:8]

    candidates: list[list[str]] = []
    for idx, left in enumerate(text_columns):
        for right in text_columns[idx + 1:]:
            candidates.append([left, right])
    return candidates


def _candidate_has_measure_like_column(
    columns: tuple[str, ...],
    baseline_sheet: dict[str, Any],
    revised_sheet: dict[str, Any],
) -> bool:
    profiles = _profiles_by_name(baseline_sheet) | _profiles_by_name(revised_sheet)
    return any(
        profiles.get(column, {}).get("numeric_ratio", 0) > 0.95
        for column in columns
    )


def _type_consistency_score(
    mapped_columns: dict[str, str],
    baseline_sheet: dict[str, Any],
    revised_sheet: dict[str, Any],
) -> float:
    if not mapped_columns:
        return 0.0

    base_profiles = _profiles_by_name(baseline_sheet)
    rev_profiles = _profiles_by_name(revised_sheet)
    scores = []
    for base_col, rev_col in mapped_columns.items():
        base_profile = base_profiles.get(base_col)
        rev_profile = rev_profiles.get(rev_col)
        if not base_profile or not rev_profile:
            continue
        numeric_gap = abs(float(base_profile.get("numeric_ratio", 0)) - float(rev_profile.get("numeric_ratio", 0)))
        date_gap = abs(float(base_profile.get("date_like_ratio", 0)) - float(rev_profile.get("date_like_ratio", 0)))
        scores.append(max(0.0, 1.0 - max(numeric_gap, date_gap)))

    return round(mean(scores), 4) if scores else 0.0


def _profiles_by_name(sheet: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {profile.get("name"): profile for profile in sheet.get("column_profiles", [])}


def _should_skip_execution(validation: SheetValidation, config: dict[str, Any]) -> bool:
    min_executable_confidence = float(config.get("planner_min_executable_confidence", 0.62))
    min_column_mapping_score = float(config.get("planner_min_column_mapping_score", 0.6))
    max_duplicate_header_ratio = float(config.get("auto_include_max_duplicate_header_ratio", 0.35))

    if validation.overall_confidence < min_executable_confidence:
        return True
    if validation.column_mapping_score < min_column_mapping_score:
        return True
    duplicate_header_ratio = _risk_ratio(validation.risks, "high duplicate-header ratio=")
    if duplicate_header_ratio is not None and duplicate_header_ratio > max_duplicate_header_ratio:
        return True
    return False


def _risk_ratio(risks: list[str], prefix: str) -> float | None:
    for risk in risks:
        if prefix not in risk:
            continue
        try:
            return float(risk.split(prefix, 1)[1].split(";", 1)[0])
        except (ValueError, IndexError):
            return None
    return None


def _resolve_column_mapping(
    sheet_plan: SheetPlan,
    base_columns: set[str],
    rev_columns: set[str],
    base_column_names: dict[str, str],
    rev_column_names: dict[str, str],
) -> dict[str, str]:
    result: dict[str, str] = {}
    for base_col, rev_col in (sheet_plan.column_mapping or {}).items():
        resolved_base = _resolve_name(str(base_col), {name: name for name in base_columns}, base_column_names)
        resolved_rev = _resolve_name(str(rev_col), {name: name for name in rev_columns}, rev_column_names)
        if resolved_base and resolved_rev:
            result[resolved_base] = resolved_rev

    for base_col in sorted(base_columns):
        if base_col in result:
            continue
        resolved_rev = _resolve_name(base_col, {name: name for name in rev_columns}, rev_column_names)
        if resolved_rev:
            result[base_col] = resolved_rev
    return result


def _resolve_name(name: str, exact_names: dict[str, Any], normalized_names: dict[str, str]) -> str | None:
    if name in exact_names:
        return name
    return normalized_names.get(_normalize_name_key(name))


def _names_by_normalized_key(items: dict[str, Any]) -> dict[str, str]:
    result: dict[str, str] = {}
    for name in items:
        key = _normalize_name_key(name)
        if key and key not in result:
            result[key] = name
    return result


def _normalize_name_key(name: str) -> str:
    text = unicodedata.normalize("NFKC", str(name))
    return "".join(
        ch for ch in text.lower()
        if ch.isalnum() or "\u4e00" <= ch <= "\u9fff"
    )


def _base_effective_config(config: dict[str, Any]) -> dict[str, Any]:
    return {
        "default": dict(config.get("default", {}) or {}),
        "exclude_sheets": list(config.get("exclude_sheets", []) or []),
        "profile_max_rows": config.get("profile_max_rows", 80),
        "key_hint_keywords": list(config.get("key_hint_keywords", []) or []),
        "id_like_keywords": list(config.get("id_like_keywords", []) or []),
        "date_like_keywords": list(config.get("date_like_keywords", []) or []),
        "high_risk_keywords": list(config.get("high_risk_keywords", []) or []),
        "sheets": {},
    }
