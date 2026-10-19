# Forestry field survey table

[Skills index](../../README.md) | [简体中文](README.zh-CN.md)

Land-use types follow the confirmed forest-purpose mapping, including energy forest land. Shrub categories use the general shrub label; existing public-benefit classifications are retained without additional authority splits. Unknown codes remain unchanged with an empty derived type. Source stock density is reference only. Post-survey summaries and fees are outside scope. Update through the original installer and preserve project workbooks.

Prepare independent 35-column source and 37-column field-observation sheets from `.xls` or `.xlsx`. The clean template is bundled. This skill covers pre-survey preparation only; post-survey status, A.1 and B.1–B.7 sheets are outside scope.

## Requirements and installation

Windows, licensed desktop Microsoft Excel, and Python 3.10+. No pip dependencies. Finish Excel first-run setup. Install `forestry-field-survey-table` from this repository with CC Switch, or copy the complete skill folder into your agent client’s skill directory, including assets and scripts.

```powershell
python "<skill-root>/scripts/forestry_field_survey_table.py" "D:/project/使用林地小班因子调查表_初表.xls"
```

Use `--sheet` for multi-sheet inputs; `--id-field` and `--area-field` select confirmed parcel identifiers and occupied area in hectares. Defaults are `小班编` and `面积`. Use `--use-nature` for explicitly confirmed 临时 or 永久; otherwise leave it blank. Optional confirmed location names: `--city`, `--county`, `--township`. `--check-only` uses Excel and temporary files without creating a deliverable.

The default output is next to the source, replacing trailing `_初表` with `_外调.xlsx`. Existing output is never overwritten; select another `--output`. Source data stays unchanged. Preserve source identifiers, real zeroes, template row heights and formatting; no background fills. Column widths and row heights follow the latest bundled master. Paired observation columns stay blank; neither sheet writes back to the other. No project-only old identifier is included. Automatic pagination is deferred. Slopes remain numeric degrees. Unmapped codes remain unchanged. Missing survey/design inputs stay blank. Forest-type code 242 maps to 能源林 (energy forest), based on confirmed applicability; the exact source-standard version remains to be documented. Numeric display precision follows the template without rounding stored values.

Read the [field mapping](../../skills/forestry-field-survey-table/references/field-mapping.md) before processing: bundled mappings are limited working conventions, not a universal official dictionary.

## Validation and upgrades

```shell
python -m unittest discover -s tests -p test_forestry_field_survey_table.py -v
```

Portable tests check transformation boundaries and the clean template. CI does not provide desktop Excel. Runtime verification saves and reopens in Excel and compares values and key styles before publishing; CI alone cannot prove Excel compatibility. Test an actual input to a temporary output after installation or template changes. Do not replace native verification with raw XML construction. Update using the original installation tool; preserve user outputs and custom mapping changes.

The bundled `references/code-dictionary.json` is loaded directly by the script: 291 field/code mappings, including 123 tree-species entries. No personal note service, account or network access is needed. Sources, historical versions and display aliases travel with the skill. Unknown extensions remain exact. Soil names are bundled but gated: use `--soil-dictionary revision-reference` only after confirming that source convention. `dictionary_coverage` separates recognized, confirmation-required and unknown values.
