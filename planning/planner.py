from __future__ import annotations

"""Planner entrypoint."""

from dataclasses import replace
from pathlib import Path
from typing import Any

from compare_excel_to_excel.planning.fallback_planner import build_fallback_plan
from compare_excel_to_excel.planning.llm_planner import build_llm_plan
from compare_excel_to_excel.planning.plan_schema import ComparisonPlan, PlanCandidate


def build_comparison_plan(
    profile: dict[str, Any],
    config: dict[str, Any],
    output_dir: str | Path | None = None,
) -> ComparisonPlan:
    planner_cfg = config.get("planner", {}) or {}
    mode = str(planner_cfg.get("mode", "auto")).lower()

    if mode in ("auto", "llm"):
        plan = build_llm_plan(profile, config, output_dir=output_dir)
        if plan is not None:
            fallback = build_fallback_plan(profile, config)
            plan.alternative_plans.append(_as_fallback_alternative(fallback))
            return plan
        if mode == "llm":
            # Keep the CLI useful even when a configured endpoint is down.
            fallback = build_fallback_plan(profile, config)
            fallback.uncertainties.append("LLM planner was requested but unavailable; used fallback planner")
            return fallback

    return build_fallback_plan(profile, config)


def _as_fallback_alternative(fallback_plan: ComparisonPlan) -> PlanCandidate:
    primary = fallback_plan.primary_plan
    return replace(
        primary,
        candidate_id="fallback_profile_plan",
        strategy=f"fallback::{primary.strategy}",
        reasons=[
            "deterministic profile planner retained as an audit comparison plan",
            *primary.reasons,
        ],
    )
