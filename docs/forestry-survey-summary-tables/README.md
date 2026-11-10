# Forestry survey summary tables

[Skills index](../../README.md) | [简体中文](README.zh-CN.md)

Generate linked `现状调查表`, `A.1`, and `B.1`–`B.7` reports from finalized `内业处理表` data, including vegetation restoration estimates for permanent and temporary forestland use.

## Requirements and installation

Windows, desktop Microsoft Excel, and Python 3.10+. No additional Python packages are required. Install `forestry-survey-summary-tables` through CC Switch or copy the complete Skill folder, preserving its templates, scripts, and references. See the [repository installation guide](../../README.md).

## Prepare the input

- Provide the finalized survey workbook. Reports read only `内业处理表`; source and field-observation sheets cannot replace it.
- Use the current 36-column template with matching field names, units, and order. Headers are on row 3 and data starts on row 4. Older schemas are not migrated automatically.
- Confirm the project name, use nature (`永久` or `临时`), occupied area in hectares, classifications, and ecological areas. Report stock uses finalized `小班蓄积m3`; `每公顷蓄积m3` must not be used to derive or validate it.
- Bundled fee rules cover Hebei only. Other regions need confirmed standards. Temporary-use estimates are reference calculations only.

## Usage

Ask your Agent to use `$forestry-survey-summary-tables` to generate linked reports from the finalized workbook. Provide its path, confirmed use nature, and project information.

The Agent verifies the [project configuration](../../skills/forestry-survey-summary-tables/references/project-rules.md) and keeps run-specific JSON in the system temporary directory. Replace `<skill-root>` and `<temp>` with actual paths:

```powershell
python "<skill-root>/scripts/summary_tables.py" "D:/project/survey.xlsx" --config "<temp>/project-rules.json"
```

## Output and verification

The script stages a temporary candidate, checks structure, template formatting, formulas, and independent subtotals, then reopens it read-only in Excel before replacing the workbook in place. Existing report sheets with matching names are rebuilt; other sheets retain their input state. Confirm the intended update scope.

Processing stops if the source changes or is locked; it does not force Excel to close. Print areas, pagination, PDF output, and print layout are not checked. Data and formatting verification does not establish print readiness.

## Updates and detailed rules

Update the complete Skill through the original installer and preserve project workbooks. Older inputs need confirmed manual migration to the current template.

See the [execution rules](../../skills/forestry-survey-summary-tables/SKILL.md) for exact fields, classification boundaries, and acceptance requirements, and the [project configuration](../../skills/forestry-survey-summary-tables/references/project-rules.md) for configuration details.
