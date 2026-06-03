from __future__ import annotations

"""Optional OpenAI-compatible LLM planner for Excel comparison plans."""

import json
import logging
import os
from pathlib import Path
import urllib.error
import urllib.request
from typing import Any

from compare_excel_to_excel.planning.plan_schema import ComparisonPlan, PlanCandidate, SheetPlan


logger = logging.getLogger("excel_compare")


def build_llm_plan(
    profile: dict[str, Any],
    config: dict[str, Any],
    output_dir: str | Path | None = None,
) -> ComparisonPlan | None:
    """Try to build a comparison plan with an OpenAI-compatible chat endpoint.

    Returns None when the planner is not configured or the endpoint fails.
    """

    llm_cfg = ((config.get("planner") or {}).get("llm") or {})
    base_url = str(llm_cfg.get("base_url") or "").rstrip("/")
    model = str(llm_cfg.get("model") or "")
    if not base_url or not model:
        return None

    api_key = str(llm_cfg.get("api_key") or "")
    api_key_env = str(llm_cfg.get("api_key_env") or "")
    if api_key_env:
        api_key = os.environ.get(api_key_env, api_key)

    payload = {
        "model": model,
        "temperature": float(llm_cfg.get("temperature", 0.1)),
        "max_tokens": int(llm_cfg.get("max_tokens", 4096)),
        "response_format": {"type": "json_object"},
        "messages": [
            {"role": "system", "content": _system_prompt()},
            {"role": "user", "content": json.dumps(_trim_profile(profile, config), ensure_ascii=False)},
        ],
    }
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"

    _write_debug_json(output_dir, "llm_planner_request.json", _redact_payload_for_debug(payload, bool(api_key)))

    timeout = float(llm_cfg.get("timeout_seconds", 60))
    url = f"{base_url}/chat/completions" if not base_url.endswith("/chat/completions") else base_url
    request = urllib.request.Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            data = json.loads(response.read().decode("utf-8"))
    except (OSError, urllib.error.URLError, json.JSONDecodeError) as exc:
        _write_debug_text(output_dir, "llm_planner_parse_error.txt", f"request_or_response_error: {exc!r}")
        logger.warning("LLM planner unavailable, fallback planner will be used: %s", exc)
        return None

    _write_debug_json(output_dir, "llm_planner_raw_response.json", data)

    try:
        content = _extract_content(data)
        _write_debug_text(output_dir, "llm_planner_raw_content.txt", content)
        raw_plan = json.loads(content)
        return _plan_from_payload(raw_plan)
    except (KeyError, TypeError, json.JSONDecodeError, ValueError) as exc:
        _write_debug_text(output_dir, "llm_planner_parse_error.txt", f"invalid_plan_error: {exc!r}")
        logger.warning("LLM planner returned invalid plan, fallback planner will be used: %s", exc)
        return None


def _system_prompt() -> str:
    return (
        "You are an Excel comparison planner. Return only JSON. "
        "Do not decide cell-level differences. Generate a candidate comparison plan: "
        "sheet pairs, header rows, key columns, exact/renamed column mapping, structure-only columns, "
        "numeric/date columns, confidence, reasons, and uncertainties. "
        "Use generic evidence from the workbook profile only; do not invent domain-specific rules. "
        "Important: baseline_sheet, revised_sheet, key_columns, structure_only_columns, ignore_columns, numeric_columns, "
        "date_columns, and column_mapping keys/values must be copied exactly from the provided profile strings. "
        "Do not translate names, insert spaces, remove spaces, rewrite punctuation, or normalize newlines. "
        "New/deleted columns must be reported as structure changes. If an added/deleted column looks auxiliary, "
        "put it in structure_only_columns, not ignore_columns. Use ignore_columns only for columns that should be "
        "completely excluded by explicit policy; otherwise keep ignore_columns empty. "
        "If a sheet looks like a cover page, legend, reference/mapping table, matrix table, or has a high "
        "duplicate_header_ratio, put it in uncertainties instead of executable sheet_pairs."
    )


def _trim_profile(profile: dict[str, Any], config: dict[str, Any]) -> dict[str, Any]:
    max_sheets = int(((config.get("planner") or {}).get("max_profile_sheets", 12)))
    max_columns = int(((config.get("planner") or {}).get("max_profile_columns", 80)))

    def trim_workbook(workbook: dict[str, Any]) -> dict[str, Any]:
        sheets = []
        for sheet in (workbook.get("sheets", []) or [])[:max_sheets]:
            sheets.append({
                "name": sheet.get("name"),
                "total_rows": sheet.get("total_rows"),
                "total_columns": sheet.get("total_columns"),
                "detected_header_row": sheet.get("detected_header_row"),
                "columns": (sheet.get("columns", []) or [])[:max_columns],
                "data_rows": sheet.get("data_rows"),
                "duplicate_header_ratio": sheet.get("duplicate_header_ratio"),
                "key_candidates": sheet.get("key_candidates", []),
                "column_profiles": (sheet.get("column_profiles", []) or [])[:max_columns],
            })
        return {"path": workbook.get("path"), "sheets": sheets}

    return {
        "task": "plan_excel_to_excel_diff",
        "baseline": trim_workbook(profile.get("baseline", {}) or {}),
        "revised": trim_workbook(profile.get("revised", {}) or {}),
        "suggested_config": profile.get("suggested_config", {}),
        "required_output_schema": {
            "primary_plan": {
                "candidate_id": "string",
                "strategy": "string",
                "confidence": "number 0-1",
                "sheet_pairs": [
                    {
                        "baseline_sheet": "string",
                        "revised_sheet": "string",
                        "header_row_baseline": "integer",
                        "header_row_revised": "integer",
                        "key_columns": ["string"],
                        "column_mapping": {"baseline column": "revised column"},
                        "structure_only_columns": ["string"],
                        "ignore_columns": ["string"],
                        "numeric_columns": {"column": "integer precision"},
                        "date_columns": ["string"],
                        "compare_formulas": "boolean",
                        "confidence": "number 0-1",
                        "reasons": ["string"],
                        "uncertainties": ["string"],
                    }
                ],
                "reasons": ["string"],
                "uncertainties": ["string"],
            },
            "alternative_plans": [],
            "uncertainties": ["string"],
        },
    }


def _plan_from_payload(payload: dict[str, Any]) -> ComparisonPlan:
    primary_payload = payload.get("primary_plan") or {}
    primary = _candidate_from_payload(primary_payload, "primary_llm")
    alternatives = [
        _candidate_from_payload(item, f"alternative_{idx}")
        for idx, item in enumerate(payload.get("alternative_plans", []) or [], start=1)
    ]
    return ComparisonPlan(
        version=str(payload.get("version") or "1.0"),
        planner=str(payload.get("planner") or "llm_planner"),
        primary_plan=primary,
        alternative_plans=alternatives,
        uncertainties=list(payload.get("uncertainties", []) or []),
        source="llm",
    )


def _candidate_from_payload(payload: dict[str, Any], fallback_id: str) -> PlanCandidate:
    return PlanCandidate(
        candidate_id=str(payload.get("candidate_id") or fallback_id),
        strategy=str(payload.get("strategy") or "llm_generated_plan"),
        confidence=float(payload.get("confidence") or 0.0),
        sheet_pairs=[
            _sheet_plan_from_payload(item)
            for item in payload.get("sheet_pairs", []) or []
        ],
        reasons=list(payload.get("reasons", []) or []),
        uncertainties=list(payload.get("uncertainties", []) or []),
    )


def _sheet_plan_from_payload(payload: dict[str, Any]) -> SheetPlan:
    return SheetPlan(
        baseline_sheet=str(payload.get("baseline_sheet") or ""),
        revised_sheet=str(payload.get("revised_sheet") or ""),
        header_row_baseline=int(payload.get("header_row_baseline") or 1),
        header_row_revised=int(payload.get("header_row_revised") or 1),
        key_columns=list(payload.get("key_columns", []) or []),
        column_mapping=dict(payload.get("column_mapping", {}) or {}),
        ignore_columns=list(payload.get("ignore_columns", []) or []),
        structure_only_columns=list(payload.get("structure_only_columns", []) or []),
        numeric_columns={key: int(value) for key, value in (payload.get("numeric_columns", {}) or {}).items()},
        date_columns=list(payload.get("date_columns", []) or []),
        compare_formulas=bool(payload.get("compare_formulas", True)),
        confidence=float(payload.get("confidence") or 0.0),
        reasons=list(payload.get("reasons", []) or []),
        uncertainties=list(payload.get("uncertainties", []) or []),
    )


def _extract_content(data: dict[str, Any]) -> str:
    choice = data["choices"][0]
    message = choice.get("message") or {}
    content = message.get("content") or message.get("reasoning_content") or choice.get("text")
    if not content:
        raise ValueError("LLM response does not contain message content")
    return str(content)


def _redact_payload_for_debug(payload: dict[str, Any], has_api_key: bool) -> dict[str, Any]:
    debug_payload = dict(payload)
    if has_api_key:
        debug_payload["authorization"] = "Bearer ***"
    return debug_payload


def _write_debug_json(output_dir: str | Path | None, filename: str, payload: Any) -> None:
    if output_dir is None:
        return
    try:
        path = Path(output_dir)
        path.mkdir(parents=True, exist_ok=True)
        (path / filename).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError as exc:
        logger.debug("Unable to write LLM planner debug file %s: %s", filename, exc)


def _write_debug_text(output_dir: str | Path | None, filename: str, text: str) -> None:
    if output_dir is None:
        return
    try:
        path = Path(output_dir)
        path.mkdir(parents=True, exist_ok=True)
        (path / filename).write_text(text, encoding="utf-8")
    except OSError as exc:
        logger.debug("Unable to write LLM planner debug file %s: %s", filename, exc)
