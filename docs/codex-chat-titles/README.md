# Codex Chat Titles

[简体中文](README.zh-CN.md) | English

Normalize Codex conversation names to concise Chinese sidebar titles:

```text
0903｜优化｜批次文字显示
0902｜功能｜整合快捷键提示页
0901｜修复｜ChatGPT 启动失败
```

The Skill always previews an exact old-to-new table. By default, it waits for confirmation before renaming anything; when the initial request explicitly waives a second confirmation, it can preview, validate, and apply in the same turn. Dates always come from each thread's `createdAt`, converted to `Asia/Shanghai`.

## Safety and scope

- Uses the local Codex app-server methods `project/list`, `thread/list`, `thread/read`, and `thread/name/set`.
- Opts into the app-server experimental capability required by `project/list`; that method may change between Codex CLI versions.
- Never edits `session_index.jsonl`, rollout files, project names, messages, project assignment, ordering, pin state, sections, or archive state directly.
- Includes visible, non-archived conversations across all projects by default. Archived conversations are included only when explicitly requested.
- Uses the canonical project name when assigned, otherwise the thread `cwd` directory name, only to prevent repeating it in the topic.
- Skips a conversation when its type or topic cannot be determined reliably.
- Uses one ASCII space at every boundary between Chinese text and ASCII letters or digits while preserving identifiers and brand names internally.
- Binds the approved preview to its mapping file with SHA-256 and rechecks the old title and `createdAt` before the first update.

## Requirements

- Python 3.10+
- A local `codex` CLI on `PATH` with the required app-server methods
- An Agent client that supports `SKILL.md` and local command execution

No third-party Python packages, API keys, hooks, or background services are required. Set `CODEX_CLI_PATH` only when `codex` is not available on `PATH`.

## Install with CC Switch

CC Switch 3.19.2+ can discover this Skill from the repository:

1. Open **Skills** in CC Switch.
2. Open **Repository Manager**.
3. Add `https://github.com/sunfing/agent-skills` on branch `main`.
4. Refresh Skills, search for `codex-chat-titles`, then install and enable it for Codex.
5. Restart Codex or open a new thread if the Skill does not appear immediately.

## Install with Codex

Windows PowerShell:

```powershell
python "$env:USERPROFILE\.codex\skills\.system\skill-installer\scripts\install-skill-from-github.py" `
  --repo sunfing/agent-skills `
  --ref main `
  --path skills/codex-chat-titles `
  --dest "$env:USERPROFILE\.codex\skills"
```

macOS/Linux:

```shell
python3 ~/.codex/skills/.system/skill-installer/scripts/install-skill-from-github.py \
  --repo sunfing/agent-skills \
  --ref main \
  --path skills/codex-chat-titles \
  --dest ~/.codex/skills
```

## Use

Invoke the Skill explicitly:

```text
$codex-chat-titles Preview standardized titles for all my visible Codex conversations.
```

In the default confirmation mode, the first response is only the preview table:

```markdown
| 原名称 | 新名称 |
| --- | --- |
| 优化批次文字显示 | 0903｜优化｜批次文字显示 |
```

By default, no title changes occur until you explicitly confirm that table. After confirmation, the Skill applies only the approved mapping and reports changed, skipped, and failed counts.

To complete the operation in one turn, explicitly waive the second confirmation in the initial request:

```text
$codex-chat-titles Normalize all my visible Codex conversation titles and apply directly after preview without a second confirmation.
```

The waiver applies only to the mapping generated for that request and does not carry over to later runs. Both modes retain the SHA-256, old-title, and `createdAt` checks.

## Compatibility check

The implementation follows the app-server protocol shipped with the installed Codex CLI. Older CLI versions may not expose all required methods. To verify read-only access manually:

```powershell
python "<skill-directory>\scripts\chat_titles.py" export --out "$env:TEMP\codex-chat-titles-export.json"
```

```shell
python3 "<skill-directory>/scripts/chat_titles.py" export --out "/tmp/codex-chat-titles-export.json"
```

Choose a new output path because the CLI refuses to overwrite files.
