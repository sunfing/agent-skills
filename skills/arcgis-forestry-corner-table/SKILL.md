---
name: arcgis-forestry-corner-table
description: 将 ArcGIS“要素折点转点 + 添加 XY 坐标”后导出的林勘/使用林地拐点 .xls、.xlsx 或 .csv 初表，整理成报批用拐点坐标 Excel。自动识别 CGCS2000 3°高斯-克吕格带号、地块与多个闭合面环，生成 J1/J2…J1 编号、报批文本、核查辅助、地块汇总和表B.8三栏版。用户提到林地红线拐点、ArcGIS 导出初表、POINT_X/POINT_Y、Zone39/Zone40、永久/临时使用林地坐标表时使用。
---

# 林勘 ArcGIS 拐点坐标表整理

用于把 ArcGIS 已经完成以下步骤后导出的属性表整理成正式 Excel：

拟使用林地面 → 投影到带号 Zone → 要素折点转点(ALL) → 添加 XY 坐标 → 导出初表

本 Skill 只负责从“初表”到“最终 Excel”。它不替用户猜测、定义或投影未知坐标系。

## 输入要求

- 首选 ArcGIS 导出的 .xls；也支持 .xlsx 和 .csv。
- 必须包含 POINT_X、POINT_Y。
- 必须存在地块字段，默认优先识别：序号、地块号、小班号、DKBH。
- POINT_X 必须已经是带号高斯-克吕格 Easting，例如：
  - Zone 39：约 39xxxxxx.xxx
  - Zone 40：约 40xxxxxx.xxx
- POINT_Y 是 Northing，一般为数百万米量级。

如果 POINT_X 仍是 432xxx、603xxx 这类不带带号的数值，停止处理，明确提示用户先在 ArcGIS 使用 Project（投影）转为 CGCS2000 3 Degree GK Zone N 后重新“添加 XY 坐标”并导出。不要通过给数字加前缀伪造带号，也不要自动执行 Define Projection。

## 执行流程

1. 找到用户指定的初表文件，不修改原文件。
2. 运行：

   python "<skill-root>/scripts/arcgis_corner_table.py" "<input.xls>"

   可选参数：

   --output "<output.xlsx>"
   --expected-zone 39
   --parcel-field 序号
   --small-area-threshold 0.0001
   --closure-tolerance 0.001
   --b8-rows 50
   --check-only
   --json

3. 如果提示缺少 Python 包，只安装本 Skill 根目录 requirements.txt 中列出的依赖，然后重试。
4. 检查脚本摘要：带号、地块数、面环数、拐点记录数、警告项。
5. 把生成的 .xlsx 返回给用户，并重点报告自动识别带号、地块数、独立闭合面环数、极小面积地块和多面环地块。

## 输出规则

正式报批坐标使用国内测量习惯：

- 纵坐标（X） = POINT_Y = Northing
- 横坐标（Y） = POINT_X = Easting
- 坐标保留 3 位小数（0.001 m）
- 每个独立面环按原始顶点顺序编号：J1, J2, ... , Jn, J1
- 同一地块存在多个独立面环时，每个面环均重新从 J1 开始；不能把不相连的面环串成一圈。

输出工作簿包含：

1. 拐点坐标表
2. 报批文本
3. 核查辅助
4. 地块汇总
5. 附表B.8三栏版

## 安全与校验

- 不根据文件名单独判断带号，以 POINT_X 数值为主。
- 一个初表中出现多个带号时停止并报错。
- 环未闭合、唯一顶点少于 3 个、POINT_X/POINT_Y 非数值时停止并报错。
- 面积 <= 0.0001 ha 默认标记“面积极小，请复核”，但不擅自删除。
- 保持 ArcGIS 导出的原始行顺序；不得按坐标大小重新排序顶点。
- 不覆盖输入文件。默认输出名把 _初表 替换为 _拐点坐标表，扩展名改为 .xlsx。

详细字段映射和原理见 references/workflow.md。
