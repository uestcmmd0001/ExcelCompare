# Excel to Excel 结构化比对工具

这是一个参照 `pmod` 分层方式搭建的 Excel 结构化差异引擎。它把两份 Excel 解析成统一数据模型，自动生成比对计划，经过程序校验后，再按 sheet、业务主键、行、列、单元格和公式输出可审计差异报告。

设计原则：**不要把任何业务样表字段写死在代码里**。默认流程不是“用户先确认规则再比对”，而是：

```text
profile -> comparison_plan -> plan_validate -> effective_config -> diff -> audit_report
```

系统先自动给出一版可执行结果，审计员再根据报告里的计划依据、置信度和风险进行后置复核。

## 输入语义

- `baseline`：我方初版、基准版、原稿。
- `revised`：乙方返回版、修改版、待核版。

差异方向固定为：`revised` 相对 `baseline` 的新增、删除、修改。

## 当前 MVP 能力

- 支持 `.xlsx`，并可读取老格式 `.xls`。
- 支持多 sheet。
- 支持自动识别表头行，并允许配置覆盖。
- 支持自动建议业务主键列，并允许配置覆盖。
- 支持只作结构变化留痕的辅助列，不参与逐格单元格比对。
- 支持配置真正忽略列，但默认不建议使用。
- 支持数字精度、日期、大小写不敏感等基础标准化。
- 支持 sheet 新增/删除。
- 支持列新增/删除。
- 支持行新增/删除。
- 支持行移动识别。
- 支持单元格值修改。
- 支持公式修改。
- 支持自动生成 `comparison_plan.json`。
- 支持程序校验计划并输出 `plan_validation.json`。
- 支持报告中展示“比对计划说明”“计划对比”和“计划风险”。
- 输出 `diff_results.json` 和 `diff_results.xlsx`。

## 推荐运行方式

在 `/Users/duananduo/Desktop/实习-文本比对需求` 下运行：

```bash
python3 -m compare_excel_to_excel.main \
  --baseline baseline.xlsx \
  --revised revised.xlsx \
  --run-dir runs/excel_to_excel_test \
  --config compare_excel_to_excel/config.yaml
```

输出文件：

```text
runs/excel_to_excel_test/output/diff_results.json
runs/excel_to_excel_test/output/diff_results.xlsx
runs/excel_to_excel_test/output/workbook_profile.json
runs/excel_to_excel_test/output/comparison_plan.json
runs/excel_to_excel_test/output/plan_validation.json
runs/excel_to_excel_test/output/effective_config.yaml
runs/excel_to_excel_test/output/suggested_config.yaml
runs/excel_to_excel_test/logs/excel_compare.log
```

`diff_results.xlsx` 中除了差异明细，还会包含：

- `比对计划说明`：系统采用了哪些 sheet、表头行、主键列、只作结构留痕列、真正忽略列、数字/日期列，以及各项校验分。
- `计划对比`：展示 LLM 主计划、确定性 fallback 计划、最终校验计划之间的关键差异。
- `计划风险`：主键弱、列映射弱、类型不一致、sheet 无法可靠匹配等风险。

### 诊断模式

如果只想看画像和自动计划，不执行差异计算，可以跑：

```bash
python3 -m compare_excel_to_excel.main \
  --baseline baseline.xlsx \
  --revised revised.xlsx \
  --run-dir runs/excel_profile \
  --config compare_excel_to_excel/config.yaml \
  --profile-only
```

输出：

```text
runs/excel_profile/output/workbook_profile.json
runs/excel_profile/output/suggested_config.yaml
runs/excel_profile/output/comparison_plan.json
runs/excel_profile/output/plan_validation.json
runs/excel_profile/output/effective_config.yaml
```

### 调试时临时收窄 sheet

```bash
python3 -m compare_excel_to_excel.main \
  --baseline baseline.xlsx \
  --revised revised.xlsx \
  --run-dir runs/excel_to_excel_test \
  --config compare_excel_to_excel/config.yaml \
  --include-sheets "需要比对的sheet名称"
```

`--include-sheets` 是调试/收窄范围用的，不是通用系统的必需参数。

## 配置说明

核心配置在 `config.yaml`：

```yaml
default:
  header_row: auto
  key_columns: auto
  structure_only_columns: []
  ignore_columns: []
  numeric_columns: {}
  date_columns: []
  case_insensitive_columns: []
  compare_formulas: true

sheets: {}
```

默认配置不内置任何业务字段名。如果某类业务确实有稳定字段提示，可以在独立业务配置文件里补充 `key_hint_keywords / id_like_keywords / date_like_keywords / high_risk_keywords`，不要改代码。

`planner.mode` 支持：

- `auto`：优先尝试 LLM planner，未配置或失败时回退 fallback，默认模式。
- `fallback`：只使用确定性画像 planner。
- `llm`：请求 LLM planner；失败时仍回退 fallback，保证命令可产出结果。

LLM planner 只负责生成比对计划，不直接输出差异事实。单元格差异仍由确定性 diff 引擎计算。

### structure_only_columns 与 ignore_columns

审计语义上，两者不要混用：

- `structure_only_columns`：列新增/删除必须报告为结构变化，但该列不参与逐格单元格比对。适合原始行号、工具生成辅助列、导入批次列等。
- `ignore_columns`：真正完全忽略列。默认应为空，只在明确业务规则要求时配置。即便列存在性变化，系统仍会留痕，避免静默吞掉结构变化。

LLM planner 会优先使用 `structure_only_columns`。如果模型误把单侧新增列放进 `ignore_columns`，validator 会把它迁移到 `structure_only_columns`，以保证结构变化可审计。

## 报告结构

Excel 报告包含：

- `汇总`
- `确认差异`
- `结构变化`
- `比对计划说明`
- `计划对比`
- `计划风险`
- `匹配风险`
- `需人工复核`
- `新增行`
- `删除行`
- `修改单元格`
- `全部结果`

字段包含 sheet、业务主键、原行号、新行号、列名、原值、新值、说明、规则来源等，方便后续接 API 或前端。

## 测试

```bash
python3 -m unittest discover -s compare_excel_to_excel/tests -v
```

## 后期其他 Excel 怎么处理

不要改代码，也不要为某个样表在代码里写 sheet 名或列名。

默认做法是直接运行完整比对，先拿到系统自动结果；审计员再查看：

1. `diff_results.xlsx` 的差异明细。
2. `比对计划说明` 中采用的 sheet、表头行、主键、列映射和置信度。
3. `计划风险` 中标出的不确定点。

如果完全没有稳定业务主键，系统会退回更保守的结果层，把相关新增/删除/重复主键标为需复核。后续可继续增加“组合主键搜索 / 相似行匹配 / LLM 字段映射”能力。

## 后续扩展

- 格式差异：颜色、字体、边框、隐藏行列、合并单元格。
- 批注差异。
- `.csv` 支持。
- AI 辅助识别表头和主键。
- 差异确认 / 驳回流程。
- 带颜色标注的 revised 工作簿。
