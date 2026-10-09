# 工作流与字段规则

## 1. ArcGIS 前置流程

正式使用本 Skill 前，ArcGIS 侧应完成：

1. 确认拟使用林地范围面。
2. 确认源坐标系正确；不要把“定义投影”当作“投影”。
3. 如报批需要带号坐标，使用 Project 转到 CGCS2000 3 Degree GK Zone N。
4. Feature Vertices To Points，点类型选 ALL。
5. Add XY Coordinates。
6. 导出属性表 .xls、.xlsx 或 .csv。

## 2. 坐标方向

ArcGIS：

- POINT_X = Easting（东坐标）
- POINT_Y = Northing（北坐标）

林勘/测量报批表：

- x / 纵坐标(X) = Northing = POINT_Y
- y / 横坐标(Y) = Easting = POINT_X

## 3. 带号识别

带号 Easting 常见形态：

- 39603020.068 → Zone 39
- 40432271.870 → Zone 40

算法取 floor(POINT_X / 1,000,000)，并要求全表一致且处于合理 3°带范围。

不带带号的 432xxx、603xxx 等不能靠字符串拼接变成报批坐标，必须回 ArcGIS 正确投影。

## 4. 面环识别

ArcGIS 对 Polygon 使用 ALL 提取折点时，一个闭合环通常以起点重复作为最后一个记录：

A → B → C → D → A

转换为：

J1 → J2 → J3 → J4 → J1

如果同一地块后续继续出现另一组 E → F → G → E，这是第二个独立面环，应重新从 J1 编号。

## 5. 默认字段别名

- 地块：序号, 地块号, 小班号, DKBH
- 面积：面积, 小班面积, TBMJ, XBMJ, Shape_Area
- 地类：地类, DI_LEI, DLMC
- 森林类别：森林类, 森林类别, SEN_LIN_LB, SBN_LIN_LB
- 保护等级：保护等, 保护等级, BH_DJ
- 坐标：POINT_X, POINT_Y

## 6. 已知代码标准化

森林类别：

- 011 → 重点公益林
- 012 → 一般公益林
- 021 → 重点商品林
- 022 → 一般商品林

保护等级：

- 1 → Ⅰ
- 2 → Ⅱ
- 3 → Ⅲ
- 4 → Ⅳ

如果输入已经是文字或罗马数字，原样保留。
