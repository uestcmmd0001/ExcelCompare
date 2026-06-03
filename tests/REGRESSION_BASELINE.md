# Excel Compare Regression Baseline

This document defines the generic regression cases that should be kept for this project. The cases must not use any sample-specific sheet or column names from a real business workbook.

## Goal

Every planner, validator, parser, differ, or exporter change should be checked against these facts:

- row insertions should not create noisy cell diffs when a business key is reliable;
- real added/deleted/modified rows and cells should be visible;
- structure changes should never be silently swallowed;
- auxiliary/tool columns should be reported as structure changes and skipped for cell-by-cell comparison;
- LLM plans should be validated and normalized before becoming executable config.

## Required Minimal Cases

1. `cell_value_modified`

   Baseline and revised share the same sheet, header row, and business key. One numeric/text cell changes.

   Expected:

   - one `cell_modified`;
   - result layer `confirmed_diff`;
   - old/new values preserved.

2. `row_added_deleted`

   Revised adds one business-key row and removes one baseline row.

   Expected:

   - one `row_added`;
   - one `row_deleted`;
   - no unrelated cell diffs.

3. `row_moved`

   Revised inserts a row or reorders rows while business-key rows remain unchanged.

   Expected:

   - stable matching by key;
   - optional `row_moved`;
   - no cascade of cell diffs.

4. `structure_only_column_added`

   Revised adds a tool/auxiliary column such as a generated source row id.

   Expected:

   - one `column_added`;
   - result layer `structure_change`;
   - `details.ignore_policy = structure_only`;
   - no cell-by-cell diffs for that column.

5. `business_column_added`

   Revised adds a real business column that is not marked structure-only.

   Expected:

   - one `column_added`;
   - result layer `structure_change`;
   - `details.ignore_policy = compared`;
   - needs human review.

6. `column_renamed_candidate`

   Revised renames a column while values and types align.

   Expected for the current MVP:

   - column added/deleted structure changes may appear;
   - report should make this auditable.

   Future expected behavior after column mapping support improves:

   - validator accepts a high-confidence mapping;
   - cell comparison uses the mapped canonical column.

7. `llm_plan_normalization`

   A synthetic LLM plan changes punctuation, spacing, or case in sheet/column names.

   Expected:

   - validator resolves names back to profile headers;
   - executable config uses only real sheet/column names from the workbook profile;
   - one-sided columns suggested as ignored are migrated to `structure_only_columns`.

## Output Files To Compare

For each case, keep these as the audit baseline:

- `output/diff_results.json`
- `output/comparison_plan.json`
- `output/plan_validation.json`
- `output/effective_config.yaml`

The Excel report is user-facing and should be spot-checked, but JSON/YAML files are the primary regression assertions.

## Recommended Command Shape

Use a separate run directory per case:

```bash
python3 -m compare_excel_to_excel.main \
  --baseline tests/fixtures/<case>/baseline.xlsx \
  --revised tests/fixtures/<case>/revised.xlsx \
  --run-dir runs/regression/<case> \
  --config compare_excel_to_excel/config.yaml
```

For deterministic CI, use a config with:

```yaml
planner:
  mode: fallback
```

For LLM-specific manual verification, use the normal config and inspect:

- `llm_planner_request.json`
- `llm_planner_raw_response.json`
- `llm_planner_raw_content.txt`
- `llm_planner_parse_error.txt` if present

