# Tree ownership certificates

[Skills index](../../README.md) | [简体中文](README.zh-CN.md)

Verify parcel data from a designated finalized survey workbook and prepare permanent or temporary tree ownership certificates grouped by tree owner, with a signing coordination cover.

## Requirements and installation

Windows, desktop Microsoft Excel and Word, and Python 3.10+. The script uses the Python standard library and calls Office through PowerShell; no additional Python packages are required. Install `forestry-tree-ownership-certificate` through CC Switch or copy the complete Skill folder. See the [repository installation guide](../../README.md).

## Prepare the input

- Provide the finalized survey workbook. The default sheet is `A.1`. Missing sheets or required fields stop generation; other sheets are not used as automatic substitutes.
- Confirm the approved project name, applicant, use nature (`永久` or `临时`), administrative locations, actual tree owners, and evidence supporting undisputed ownership. Land ownership codes alone do not establish tree ownership.
- Confirm occupied area in hectares, total parcel stock in cubic metres, and total tree counts. Missing values are not zero; omitting unavailable tree counts requires explicit authorization.
- Supply the project root, or use the standard source path under `05_报告表格与图件` for root detection. For updates, identify the intended changes and manual content to preserve.

## Usage

Ask your Agent to use `$forestry-tree-ownership-certificate` with the workbook path, confirmed project information, and use nature. You do not need to write JSON or supply spreadsheet column letters yourself.

The Agent verifies actual headers, row ranges, and owner groups, then keeps the configuration in the system temporary directory. This command stages a candidate without writing to the final output folder. Replace `<skill-root>`, `<temp>`, and `<stage-dir>` with actual paths:

```powershell
python "<skill-root>/scripts/ownership_certificate.py" --config "<temp>/project.json"
```

After checking source data, groups, quantities, and every page, publish within the existing creation or update authorization:

```powershell
python "<skill-root>/scripts/ownership_certificate.py" --publish-stage "<stage-dir>/stage.json"
```

## Output and verification

The default is one DOCX per project and use nature under `project-root/06_公示查验与报批/永久-or-临时/林木权属证明/`, using the actual Chinese use-nature value as the folder name. A coordination cover precedes separate village or forest-farm certificates. Dates, signatures, and seals remain blank. Configuration, audit information, and previews stay in the system temporary directory.

Publishing rechecks source and target changes. Updates regenerate the document, so required manual content must be reviewed and preserved beforehand. Word reopening and structural checks do not replace visual inspection of every page. Generated certificates do not establish legal ownership or approval readiness.

## Updates and detailed rules

Update through the original installer and preserve project outputs. See the [execution rules](../../skills/forestry-tree-ownership-certificate/SKILL.md) for grouping, rounding, signing, and update boundaries, and the [project configuration](../../skills/forestry-tree-ownership-certificate/references/project-config.md) for mappings and configuration fields.
