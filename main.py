from __future__ import annotations

"""Structured Excel-to-Excel comparison CLI.

Usage:
  python3 -m compare_excel_to_excel.main \
    --baseline baseline.xlsx \
    --revised revised.xlsx \
    --run-dir runs/excel_test \
    --config compare_excel_to_excel/config.yaml
"""

import argparse
import logging
from pathlib import Path
from typing import Any

from compare_excel_to_excel.config import load_config
from compare_excel_to_excel.exporters.excel_exporter import export_excel
from compare_excel_to_excel.exporters.json_exporter import save_json, save_yaml
from compare_excel_to_excel.parsers.workbook_parser import parse_workbook
from compare_excel_to_excel.planning.plan_schema import to_dict
from compare_excel_to_excel.planning.plan_validator import validate_plan
from compare_excel_to_excel.planning.planner import build_comparison_plan
from compare_excel_to_excel.profilers.workbook_profiler import profile_workbook_pair
from compare_excel_to_excel.processing.differ import diff_summary, diff_workbooks, serialize_diff_items


logger = logging.getLogger("excel_compare")


def setup_logger(run_dir: str | Path) -> None:
    log_dir = Path(run_dir) / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "excel_compare.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(log_path, encoding="utf-8"),
        ],
        force=True,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="结构化 Excel 差异比对工具")
    parser.add_argument("--baseline", required=True, help="基准 Excel（我方初版/原稿）")
    parser.add_argument("--revised", required=True, help="待核 Excel（乙方返回版/修改版）")
    parser.add_argument("--run-dir", required=True, help="运行目录")
    parser.add_argument(
        "--config",
        default=str(Path(__file__).parent / "config.yaml"),
        help="配置文件路径",
    )
    parser.add_argument(
        "--profile-only",
        action="store_true",
        help="只生成画像、自动比对计划和校验结果，不执行差异比对",
    )
    parser.add_argument(
        "--include-sheets",
        nargs="+",
        help="仅比对指定 sheet；可用于调试或人工收窄范围",
    )
    args = parser.parse_args()

    run_dir = Path(args.run_dir)
    run_dir.mkdir(parents=True, exist_ok=True)
    setup_logger(run_dir)

    logger.info("=" * 60)
    logger.info("Excel 结构化比对启动")
    logger.info("baseline: %s", args.baseline)
    logger.info("revised: %s", args.revised)
    logger.info("run_dir: %s", run_dir)
    logger.info("=" * 60)

    config = load_config(args.config)
    if args.include_sheets:
        config["include_sheets"] = args.include_sheets
    output_dir = run_dir / "output"

    profile = profile_workbook_pair(args.baseline, args.revised, config)
    suggested_config = profile["suggested_config"]
    comparison_plan = build_comparison_plan(profile, config, output_dir=output_dir)
    plan_validation = validate_plan(comparison_plan, profile, config)
    effective_config = _merge_config(plan_validation.effective_config, config)

    save_json(output_dir / "workbook_profile.json", profile)
    save_yaml(output_dir / "suggested_config.yaml", suggested_config)
    save_json(output_dir / "comparison_plan.json", to_dict(comparison_plan))
    save_json(output_dir / "plan_validation.json", to_dict(plan_validation))
    save_yaml(output_dir / "effective_config.yaml", effective_config)
    logger.info("画像输出: %s", output_dir / "workbook_profile.json")
    logger.info("自动比对计划: %s", output_dir / "comparison_plan.json")
    logger.info("计划校验结果: %s", output_dir / "plan_validation.json")
    logger.info("有效配置: %s", output_dir / "effective_config.yaml")

    if args.profile_only:
        logger.info("profile-only 模式，已生成自动计划诊断，跳过差异比对")
        return

    baseline = parse_workbook(args.baseline, effective_config, role="baseline", include_all_sheets=True)
    revised = parse_workbook(args.revised, effective_config, role="revised", include_all_sheets=True)
    logger.info("baseline sheets: %s", ", ".join(baseline.sheets))
    logger.info("revised sheets: %s", ", ".join(revised.sheets))

    diff_items = diff_workbooks(baseline, revised, effective_config)
    summary = diff_summary(diff_items)

    save_json(output_dir / "diff_results.json", {
        "baseline": str(Path(args.baseline)),
        "revised": str(Path(args.revised)),
        "summary": summary,
        "comparison_plan": to_dict(comparison_plan),
        "plan_validation": to_dict(plan_validation),
        "effective_config": effective_config,
        "diff_results": serialize_diff_items(diff_items),
    })
    excel_path = export_excel(run_dir, diff_items, comparison_plan=comparison_plan, plan_validation=plan_validation)

    logger.info("差异总数: %s", summary["total"])
    logger.info("JSON 输出: %s", output_dir / "diff_results.json")
    logger.info("Excel 输出: %s", excel_path)


def _merge_config(auto_config: dict[str, Any], user_config: dict[str, Any]) -> dict[str, Any]:
    """Use auto-planned structure, while letting explicit user config win."""

    merged = _deep_merge(auto_config, user_config)
    if user_config.get("include_sheets"):
        merged["include_sheets"] = user_config["include_sheets"]
    if user_config.get("exclude_sheets"):
        merged["exclude_sheets"] = user_config["exclude_sheets"]
    return merged


def _deep_merge(base: Any, override: Any) -> Any:
    if isinstance(base, dict) and isinstance(override, dict):
        result = dict(base)
        for key, value in override.items():
            if value is None:
                continue
            if value == {} and isinstance(result.get(key), dict):
                continue
            if value == [] and isinstance(result.get(key), list):
                continue
            result[key] = _deep_merge(result.get(key), value)
        return result
    return override if override not in ({}, [], None) else base


if __name__ == "__main__":
    main()
