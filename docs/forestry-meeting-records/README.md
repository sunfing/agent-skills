# Forestland-use meeting records

[Skills index](../../README.md) | [简体中文](README.zh-CN.md)

Prepare meeting record drafts from designated tree ownership certificates. Separate permanent and temporary use, apply the appropriate village or forest-farm format, and leave meeting facts blank for completion at the actual meeting.

## Requirements and installation

Install `forestry-meeting-records` through CC Switch or copy the complete Skill folder, preserving `assets/`, `references/`, and `agents/`. See the [repository installation guide](../../README.md).

There is no fixed script entry point or requirement to add Python packages. Use existing document editing and rendering tools. When Microsoft Word is available, reopen in a separate instance, check pagination, and export a preview; otherwise use an existing suitable renderer.

## Prepare the input

- Designate the certificate versions, project root, and use nature (`永久` or `临时`). Survey workbooks and older certificates are not automatic substitutes.
- Confirm whether each meeting body is a village or forest farm. Preserve source project names, applicant names, locations, parcel identifiers, and area precision, including leading zeroes and non-contiguous identifiers.
- For existing records, identify the update scope. Completed meeting facts, manual text, and signing content must not be overwritten by refilling the template.

## Usage

Ask your Agent to use `$forestry-meeting-records` to prepare village and forest-farm meeting drafts from the specified certificates. Provide their paths and the project root.

After checking sources, the Agent fills the [bundled template](../../skills/forestry-meeting-records/assets/meeting-record-template.docx). Dates, meeting location, chair, recorder, and attendance counts stay blank. The body describes discussion topics without inventing voting results, meeting resolutions, or ownership conclusions.

## Output and verification

The default folder is `project-root/06_公示查验与报批/永久-or-临时/会议记录/`, using the actual Chinese use-nature value. Village records are combined per project and use nature into `会议记录（永久）.docx` or `会议记录（临时）.docx`, with each village starting on a separate page. Each forest farm normally gets its own `会议记录（林场全称）.docx`.

Check every record against its source and inspect every page for names, identifiers, areas, blank fields, signature space, and signing blocks. Each record defaults to one page; longer content needs confirmation before continuing onto another page. Deliver only final DOCX files after verification; extracted data, backups, and previews remain in the system temporary directory.

Recheck source and existing target files before writing. Drafts still require actual meeting information and on-site signatures and seals. Generation does not establish that a meeting occurred or that approval requirements are met.

## Updates and detailed rules

Update through the original installer without regenerating or overwriting completed project records. See the [execution rules](../../skills/forestry-meeting-records/SKILL.md) for scope and the [record rules](../../skills/forestry-meeting-records/references/meeting-record-rules.md) for fields, wording, and layout.
