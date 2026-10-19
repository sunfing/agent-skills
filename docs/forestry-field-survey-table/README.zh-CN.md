# 使用林地小班因子外调表整理

[Skills 总目录](../../README.zh-CN.md) | [English](README.md)

将 `.xls`、`.xlsx` 初表整理为独立的35列原始因子表与37列外业调查表。内置干净模板，不需要额外指定模板。仅整理外业前底表，外业后现状调查表、A.1及B.1—B.7由汇总Skill处理。

林地类型按已确认林种归并（水源林/水保林→防护林林地，能源林→能源林林地）；灌木类统一一般灌木林地，不追加公益林事权细分。未知林种保留代码且类型留空。源库蓄积仅作参考；本skill不处理外业后汇总或费用。更新通过原安装工具，保留项目成果。

## 环境与安装

Windows + 桌面版 Microsoft Excel + Python 3.10 或更新版本，无需 pip 安装依赖。Excel 必须已完成首次运行和许可激活。通过 CC Switch 从本仓库安装 `forestry-field-survey-table`，或将 `skills/forestry-field-survey-table` 完整复制到客户端技能目录。保持 assets、scripts、references 一同安装。

## 使用

```powershell
python "<skill-root>/scripts/forestry_field_survey_table.py" "D:/项目/使用林地小班因子调查表_初表.xls"
```

自然语言：“使用 $forestry-field-survey-table 将这份小班初表整理成外调表。”

默认源表列为“小班编”和“面积”。必须确认面积是以公顷计的占用面积。多工作表用 `--sheet` 指定。不同列名用 `--id-field`、`--area-field`；确认后的地名可用 `--city`、`--county`、`--township`。`--check-only` 会启动 Excel、使用临时文件，但不生成成果。

默认同目录 `_初表.xls` → `_外调.xlsx`。已有输出拒绝覆盖，可明确指定另一个 `--output`。源文件不改动。输出两张独立静态表，同名成对栏后列留空实测，不含项目旧编号，不自动回写内业表。使用性质通过 `--use-nature` 明确指定，否则留空。

列宽沿用最新模板，暂不自动分页；保留模板行高、字体、边框和合并方式，无背景填充。坡度显示角度，林种242转换为能源林，未知代码直接保留。模板显示精度可能比原数值少，底层数值不舍入。

有限工作代码映射见 [字段对应](../../skills/forestry-field-survey-table/references/field-mapping.md)，使用前核对版本，不将其当作通用官方字典。

## 验证

```powershell
python -m unittest discover -s tests -p test_forestry_field_survey_table.py -v
```

自动测试覆盖字段对应、数据边界和模板净化；跨平台 CI 不具备桌面 Excel。实际生成必须在 Windows Excel 中保存并重新打开，通过数据和样式检查才发布。不得把 CI 通过等同于原生 Excel 验证。

首次安装和模板更新后，用临时输出跑一份实际初表并检查。没有 Excel 时报告环境不满足，不自动降级拼接工作簿 XML。升级时使用原安装工具，保留用户成果，勿覆盖本地自定义代码映射。
