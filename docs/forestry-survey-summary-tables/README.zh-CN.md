# 使用林地联动报告表

[Skills 总目录](../../README.zh-CN.md) | [English](README.md)

从核定后的 `内业处理表` 生成联动的 `现状调查表`、`A.1` 和 `B.1`—`B.7`，包括永久、临时使用林地的植被恢复测算。

## 环境与安装

需要 Windows、桌面版 Microsoft Excel 和 Python 3.10+，无需安装额外 Python 包。通过 CC Switch 安装 `forestry-survey-summary-tables`，或完整复制对应 Skill 文件夹，保留其中的模板、脚本和参考文件。安装方式见 [仓库首页](../../README.zh-CN.md)。

## 准备资料

- 提供已经核定的调查工作簿。报告只读取 `内业处理表`，不能用原始因子表或外业调查表代替。
- 输入必须符合当前模板的 36 列字段、单位和顺序；表头在第 3 行，数据从第 4 行开始。旧版字段不自动迁移。
- 确认项目名称、永久／临时性质、占用面积（公顷）、分类及生态区域等项目口径。蓄积使用内业核定的 `小班蓄积m3`，不由 `每公顷蓄积m3` 反算或校验。
- 费用测算仅内置河北口径；其他地区需要另行确认标准。临时项目只作参考测算。

## 使用

向 Agent 提出：“使用 $forestry-survey-summary-tables，根据这份已核定的内业处理表生成联动报告表，使用性质为临时。”同时提供工作簿路径和已确认的项目资料。

Agent 按 [项目配置](../../skills/forestry-survey-summary-tables/references/project-rules.md) 核实口径，并将本次运行的 JSON 配置放在系统临时目录。命令入口如下，`<skill-root>` 和 `<temp>` 需替换为实际路径：

```powershell
python "<skill-root>/scripts/summary_tables.py" "D:/项目/调查表.xlsx" --config "<temp>/project-rules.json"
```

## 输出与核对

脚本先生成临时候选，核对结构、模板格式、公式及独立小计，再通过 Excel 只读重开验证，成功后原地更新工作簿。同名报告表会重建，其他工作表按输入实际状态保留；请明确本次更新范围。

源文件变化或被占用时停止，不强关 Excel。不执行打印区域、分页、PDF 或打印版式检查；数据和格式验证不代表已完成打印验收。

## 更新与详细规则

通过原安装工具更新完整 Skill 文件夹，保留项目工作簿。旧版输入须先核实并整理到当前模板，不能靠升级自动迁移。

完整字段、分类边界和验收要求见 [执行规则](../../skills/forestry-survey-summary-tables/SKILL.md)；配置字段见 [项目配置](../../skills/forestry-survey-summary-tables/references/project-rules.md)。
