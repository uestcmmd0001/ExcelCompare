from __future__ import annotations

"""Planner entrypoint."""

from typing import Any

from compare_excel_to_excel.planning.fallback_planner import build_fallback_plan
from compare_excel_to_excel.planning.llm_planner import build_llm_plan
from compare_excel_to_excel.planning.plan_schema import ComparisonPlan


def build_comparison_plan(profile: dict[str, Any], config: dict[str, Any]) -> ComparisonPlan:
    planner_cfg = config.get("planner", {}) or {}
    mode = str(planner_cfg.get("mode", "auto")).lower()

    if mode in ("auto", "llm"):
        plan = build_llm_plan(profile, config)
        if plan is not None:
            return plan
        if mode == "llm":
            # Keep the CLI useful even when a configured endpoint is down.
            fallback = build_fallback_plan(profile, config)
            fallback.uncertainties.append("LLM planner was requested but unavailable; used fallback planner")
            return fallback

    return build_fallback_plan(profile, config)

