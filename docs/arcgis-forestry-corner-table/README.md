# ArcGIS Forestry Corner Table Agent Skill

[Skills index](../../README.md) | English | [简体中文](README.zh-CN.md)

Convert forestry/land-use corner-point tables exported from ArcGIS after `Feature Vertices To Points (ALL)` and `Add XY Coordinates` into approval-ready `.xlsx` workbooks.

The Skill does not modify source files or guess/fix unknown projections. `POINT_X` must already be a band-prefixed CGCS2000 3-degree Gauss-Kruger Easting such as Zone 39 `39xxxxxx.xxx` or Zone 40 `40xxxxxx.xxx`.

## Output

The workbook contains:

- `拐点坐标表`
- `报批文本`
- `核查辅助`
- `地块汇总`
- `附表B.8三栏版`

Coordinate convention:

```text
X / northing = POINT_Y
Y / easting  = POINT_X
```

Coordinates are rounded to 0.001 m. Each independent ring is numbered `J1 → J2 → … → Jn → J1`; numbering restarts at J1 for another ring in the same parcel.

## ArcGIS prerequisite workflow

1. Confirm the final proposed forest-land polygon layer.
2. Confirm the source CRS definition is correct.
3. Use **Project** to convert to `CGCS2000 3 Degree GK Zone N` when band-prefixed approval coordinates are required.
4. Run **Feature Vertices To Points** with `ALL`.
5. Run **Add XY Coordinates**.
6. Export the attribute table to `.xls`, `.xlsx`, or `.csv`.

Do not confuse **Define Projection** with **Project**. Unbanded six-digit Eastings such as `432xxx` or `603xxx` are rejected.

## Requirements

- Python 3.10+
- Local command execution
- Packages in `requirements.txt`

Install:

```shell
python -m pip install -r <skill-root>/requirements.txt
```

## Usage

```text
$arcgis-forestry-corner-table
Process D:\project\forest-temp-Zone39_initial.xls and create the approval-ready corner-coordinate workbook.
```

Direct CLI:

```shell
python "<skill-root>/scripts/arcgis_corner_table.py" "/path/to/input.xls"
```

Useful options: `--expected-zone`, `--parcel-field`, `--small-area-threshold`, `--closure-tolerance`, `--b8-rows`, `--check-only`, and `--json`.

## Install from this repository

With CC Switch 3.19.2+, add `https://github.com/sunfing/agent-skills` on branch `main`, refresh Skills, and install `arcgis-forestry-corner-table`.

With the Codex Skill Installer:

```shell
python ~/.codex/skills/.system/skill-installer/scripts/install-skill-from-github.py \
  --repo sunfing/agent-skills \
  --ref main \
  --path skills/arcgis-forestry-corner-table \
  --dest ~/.codex/skills
```

See the Chinese guide for Windows PowerShell examples and detailed forestry workflow notes.
