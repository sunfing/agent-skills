# 林木权属证明

[Skills 总目录](../../README.zh-CN.md) | [English](README.md)

从用户指定的完善后调查表核对地块数据，按林木权利人制作永久或临时林木权属证明，附签章对接首页。

## 环境与安装

需要 Windows、桌面版 Microsoft Excel／Word 和 Python 3.10+。脚本使用 Python 标准库，通过 PowerShell 调用 Office，无需安装额外 Python 包。通过 CC Switch 安装 `forestry-tree-ownership-certificate`，或完整复制对应 Skill 文件夹。安装方式见 [仓库首页](../../README.zh-CN.md)。

## 准备资料

- 提供完善后的调查表，默认读取 `A.1`。缺少该表或必需字段时停止，不能自动切换其他表补齐。
- 核实正式工程名称、建设单位、永久／临时性质、行政区划、实际林木权利人及无争议依据。林地权属代码不能直接证明林木所有权。
- 明确占用面积（公顷）、核定地块总蓄积（立方米）和总株数。缺失不能当作零，缺少株数仅在明确授权后省略。
- 提供项目根目录；源表位于标准 `05_报告表格与图件` 目录下时可按该结构定位。已有证明需要更新时，说明更新范围和应保留的人工内容。

## 使用

向 Agent 提出：“使用 $forestry-tree-ownership-certificate，根据这份调查表制作临时林木权属证明。”同时提供源文件路径及已核实的项目资料，无需自行填写 JSON 或列字母。

Agent 核对实际表头、行范围和分组后，将配置存放在系统临时目录。以下命令先生成临时候选，不直接写入正式目录；`<skill-root>`、`<temp>` 和 `<stage-dir>` 需替换为实际路径：

```powershell
python "<skill-root>/scripts/ownership_certificate.py" --config "<temp>/project.json"
```

候选需核对来源、分组、数量和全部页面，再在已有制作／更新授权范围内发布：

```powershell
python "<skill-root>/scripts/ownership_certificate.py" --publish-stage "<stage-dir>/stage.json"
```

## 输出与核对

默认每项目、每种使用性质一份 DOCX，保存到 `项目根目录/06_公示查验与报批/永久或临时/林木权属证明/`。首页为对接备注，各村／林场证明分别独立起页；日期、签字和印章留空。配置、核对信息及预览只留在系统临时目录。

发布前再次检查来源和目标是否变化。更新会重新生成，需事先核对并保留应保留的人工文字；不能直接覆盖已完善的内容。Word 重开和结构校验不能代替逐页视觉检查，生成完成也不代表法律确权或已满足报批条件。

## 更新与详细规则

通过原安装工具更新 Skill，保留项目成果。具体分组、取整、签章和更新边界见 [执行规则](../../skills/forestry-tree-ownership-certificate/SKILL.md)，字段映射及配置见 [项目配置](../../skills/forestry-tree-ownership-certificate/references/project-config.md)。
