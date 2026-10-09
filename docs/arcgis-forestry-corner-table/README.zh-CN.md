# ArcGIS Forestry Corner Table Agent Skill

[Skills 总目录](../../README.zh-CN.md) | [English](README.md) | 简体中文

把 ArcGIS 已完成“要素折点转点（ALL）+ 添加 XY 坐标”后导出的林勘/使用林地拐点初表（`.xls`、`.xlsx`、`.csv`）整理成报批用 `.xlsx`。

本 Skill 不修改原始文件，也不猜测、定义或修复未知坐标系。输入中的 `POINT_X` 必须已经是带号的 CGCS2000 3°高斯-克吕格 Easting，例如 Zone 39 的 `39xxxxxx.xxx` 或 Zone 40 的 `40xxxxxx.xxx`。

## 输出内容

自动生成：

- `拐点坐标表`
- `报批文本`
- `核查辅助`
- `地块汇总`
- `附表B.8三栏版`

坐标方向按林勘/测量成果习惯：

```text
纵坐标 X = POINT_Y = Northing
横坐标 Y = POINT_X = Easting
```

坐标保留 3 位小数。每个独立闭合面环按原始顶点顺序生成 `J1 → J2 → … → Jn → J1`；同一地块有多个面环时，每个面环重新从 J1 开始。

## ArcGIS 前置流程

1. 确认最终拟使用林地面。
2. 确认源坐标系定义正确。
3. 报批需要带号坐标时，用 **Project（投影）** 转到 `CGCS2000 3 Degree GK Zone N`。
4. 执行 **要素折点转点**，点类型选 `ALL`。
5. 执行 **添加 XY 坐标**。
6. 导出属性表。

不要把 **Define Projection / 定义投影** 当成 **Project / 投影**。若 `POINT_X` 仍是 `432xxx`、`603xxx` 等不带带号坐标，Skill 会拒绝生成正式成果。

## 安装依赖

```powershell
python -m pip install -r "$env:USERPROFILE\.codex\skills\arcgis-forestry-corner-table\requirements.txt"
```

## 使用

```text
$arcgis-forestry-corner-table
处理 D:\项目\林地红线临时Zone39_初表.xls，生成报批拐点坐标 Excel。
```

也可以自然语言要求：

```text
把这个 ArcGIS 导出的临时使用林地拐点初表整理成报批坐标表。
```

直接运行：

```powershell
python "<skill-root>\scripts\arcgis_corner_table.py" "D:\项目\林地红线临时Zone39_初表.xls"
```

常用参数：`--expected-zone`、`--parcel-field`、`--small-area-threshold`、`--closure-tolerance`、`--b8-rows`、`--check-only`、`--json`。

## 从本仓库安装

### CC Switch 3.19.2+

添加仓库：

```text
https://github.com/sunfing/agent-skills
```

Branch 选择 `main`，刷新后搜索 `arcgis-forestry-corner-table`。

### Codex Skill Installer

```powershell
python "$env:USERPROFILE\.codex\skills\.system\skill-installer\scripts\install-skill-from-github.py" `
  --repo sunfing/agent-skills `
  --ref main `
  --path skills/arcgis-forestry-corner-table `
  --dest "$env:USERPROFILE\.codex\skills"
```

## 自动核查

Skill 会检查必要字段、带号一致性、坐标量级、面环闭合和唯一顶点数量；保留 ArcGIS 原始顶点顺序；面积 `<= 0.0001 ha` 默认仅提示复核，不自动删除。

详细字段别名、森林类别和保护等级映射见 Skill 目录下的 `references/workflow.md`。
