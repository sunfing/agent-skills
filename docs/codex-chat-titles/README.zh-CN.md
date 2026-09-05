# Codex 对话标题

简体中文 | [English](README.md)

将 Codex 对话名称统一为简洁、适合侧栏浏览的中文标题：

```text
0903｜优化｜批次文字显示
0902｜功能｜整合快捷键提示页
```

Skill 会先输出准确的“原名称 → 新名称”预览表，收到明确确认后才执行改名。日期固定取每条对话的 `createdAt`，并转换为 `Asia/Shanghai`。

## 范围与安全边界

- 使用本机 Codex app-server 的 `project/list`、`thread/list`、`thread/read` 和 `thread/name/set`。
- 显式启用 `project/list` 所需的 app-server 实验能力；该方法可能随 Codex CLI 版本变化。
- 不直接修改 `session_index.jsonl`、rollout 文件、项目名称、对话内容、项目归属、排序、置顶、分区或归档状态。
- 默认处理所有项目中当前可见且未归档的对话；只有用户明确要求时才纳入已归档对话。
- 已归属项目时使用 canonical project 名称，否则使用对话 `cwd` 的目录名；该信息仅用于避免主题重复项目名。
- 无法可靠判断类型或主题时保留原名，不猜测。
- 通过 SHA-256 将确认内容绑定到预览映射；首次写入前重新核对旧标题和 `createdAt`。

## 环境要求

- Python 3.10+
- `PATH` 中存在支持所需 app-server 方法的本机 `codex` CLI
- Agent 客户端支持 `SKILL.md` 和本地命令执行

不需要第三方 Python 包、API key、Hook 或后台服务。只有 `PATH` 中找不到 `codex` 时，才需要设置 `CODEX_CLI_PATH`。

## 使用 CC Switch 安装

CC Switch 3.19.2+ 可以从本仓库发现此 Skill：

1. 在 CC Switch 中打开 **Skills**。
2. 打开右上角 **仓库管理**。
3. 添加 `https://github.com/sunfing/agent-skills`，Branch 选择 `main`。
4. 刷新 Skills，搜索 `codex-chat-titles`，然后安装并为 Codex 启用。
5. 如果没有立即显示，完整重启 Codex 或新建对话。

## 使用 Codex 安装

Windows PowerShell：

```powershell
python "$env:USERPROFILE\.codex\skills\.system\skill-installer\scripts\install-skill-from-github.py" `
  --repo sunfing/agent-skills `
  --ref main `
  --path skills/codex-chat-titles `
  --dest "$env:USERPROFILE\.codex\skills"
```

macOS/Linux：

```shell
python3 ~/.codex/skills/.system/skill-installer/scripts/install-skill-from-github.py \
  --repo sunfing/agent-skills \
  --ref main \
  --path skills/codex-chat-titles \
  --dest ~/.codex/skills
```

## 使用方式

显式调用 Skill：

```text
$codex-chat-titles 预览当前 Codex 所有项目下可见对话的规范化标题。
```

第一次响应只会显示预览表：

```markdown
| 原名称 | 新名称 |
| --- | --- |
| 优化批次文字显示 | 0903｜优化｜批次文字显示 |
```

明确确认前不会修改标题。确认后只应用该预览对应的映射，并报告成功、跳过和失败数量。

## 兼容性检查

脚本使用已安装 Codex CLI 随附的 app-server 协议。较旧版本可能没有全部所需方法。可手动执行只读导出检查：

```powershell
python "<skill-directory>\scripts\chat_titles.py" export --out "$env:TEMP\codex-chat-titles-export.json"
```

```shell
python3 "<skill-directory>/scripts/chat_titles.py" export --out "/tmp/codex-chat-titles-export.json"
```

脚本拒绝覆盖已有文件，因此每次应使用新的输出路径。
