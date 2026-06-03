from __future__ import annotations

"""Deterministic fallback planner used when no LLM planner is configured."""

from typing import Any

from compare_excel_to_excel.planning.plan_schema import ComparisonPlan, PlanCandidate, SheetPlan


def build_fallback_plan(profile: dict[str, Any], config: dict[str, Any]) -> ComparisonPlan:
    """Build an executable comparison plan from workbook profiles.

    This planner intentionally uses only profile statistics and generic config,
    not any sample-specific sheet or column names.
    """

    suggested = profile.get("suggested_config", {}) or {}
    baseline_sheets = {sheet["name"]: sheet for sheet in profile.get("baseline", {}).get("sheets", [])}
    revised_sheets = {sheet["name"]: sheet for sheet in profile.get("revised", {}).get("sheets", [])}
    include_sheets = suggested.get("include_sheets") or [
        name for name in baseline_sheets if name in revised_sheets
    ]

    sheet_plans: list[SheetPlan] = []
    uncertainties: list[str] = []

    for sheet_name in include_sheets:
        baseline_sheet = baseline_sheets.get(sheet_name)
        revised_sheet = revised_sheets.get(sheet_name)
        if not baseline_sheet or not revised_sheet:
            uncertainties.append(f"sheet `{sheet_name}` is not present on both sides")
            continue

        sheet_cfg = (suggested.get("sheets", {}) or {}).get(sheet_name, {}) or {}
        common_columns = [
            column for column in baseline_sheet.get("columns", [])
            if column in set(revised_sheet.get("columns", []))
        ]
        column_mapping = {column: column for column in common_columns}
        key_columns = [
            column for column in sheet_cfg.get("key_columns", []) or []
            if column in column_mapping
        ]

        sheet_plans.append(SheetPlan(
            baseline_sheet=sheet_name,
            revised_sheet=sheet_name,
            header_row_baseline=int(baseline_sheet.get("detected_header_row") or 1),
            header_row_revised=int(revised_sheet.get("detected_header_row") or 1),
            key_columns=key_columns,
            column_mapping=column_mapping,
            ignore_columns=list(sheet_cfg.get("ignore_columns", []) or []),
            structure_only_columns=list(sheet_cfg.get("structure_only_columns", []) or []),
            numeric_columns=dict(sheet_cfg.get("numeric_columns", {}) or {}),
            date_columns=list(sheet_cfg.get("date_columns", []) or []),
            compare_formulas=bool(sheet_cfg.get("compare_formulas", True)),
            confidence=0.72 if key_columns else 0.48,
            reasons=[
                "same sheet name on both sides",
                "columns mapped by exact header text",
            ],
            uncertainties=[] if key_columns else ["no reliable business key was selected automatically"],
        ))

    primary = PlanCandidate(
        candidate_id="primary_auto_profile",
        strategy="profile_statistics_exact_sheet_and_column_match",
        confidence=_average([sheet.confidence for sheet in sheet_plans]),
        sheet_pairs=sheet_plans,
        reasons=[
            "generated from workbook profile statistics",
            "uses exact sheet-name and column-name matching before deterministic diff",
        ],
        uncertainties=uncertainties,
    )

    return ComparisonPlan(
        version="1.0",
        planner="fallback_profile_planner",
        primary_plan=primary,
        alternative_plans=[],
        uncertainties=uncertainties,
        source="fallback",
    )


def _average(values: list[float]) -> float:
    if not values:
        return 0.0
    return round(sum(values) / len(values), 4)
