# Agent Skills

[English](README.md) | 简体中文

一个统一收纳非官方 Agent Skills 的仓库，适用于支持 `SKILL.md` 格式并能够执行本地命令的 Agent 客户端。

每个 Skill 位于 [`skills/`](skills/) 下，只包含 Agent 实际运行所需的说明和资源。面向用户的安装、配置和使用指南统一放在 [`docs/`](docs/) 下。

本仓库及其中的 Skills 与 OpenAI、Anthropic 或其他模型供应商无隶属关系，不提供 API 服务、额度或 API key。

## 可用 Skills

| Skill | 用途 | 文档 |
| --- | --- | --- |
| [`gpt-image`](skills/gpt-image/) | 通过用户配置的 OpenAI-compatible Image API，使用固定模型 `gpt-image-2` 生成或编辑位图。 | [English](docs/gpt-image/README.md) · [简体中文](docs/gpt-image/README.zh-CN.md) |

## 使用 CC Switch 安装

CC Switch 3.19.2+ 会递归发现仓库中的 `SKILL.md`。

1. 在 CC Switch 中打开 **Skills**。
2. 点击右上角 **仓库管理**。
3. 点击 **添加仓库**。
4. 输入 `https://github.com/sunfing/agent-skills`，Branch 选择 `main`。
5. 返回 Skills 发现页面并点击 **刷新**。
6. 搜索需要的 Skill 名称，然后安装并为目标 Agent 客户端启用。

Skills 发现页面的搜索框只过滤已配置仓库中发现的 Skill。把 GitHub URL 粘贴到搜索框不会添加仓库。

安装完成后，请继续阅读对应 Skill 的文档，完成依赖安装、环境变量配置和验证。

## 使用 Codex 安装单个 Skill

使用 Codex 自带的 Skill Installer，并指定需要安装的 Skill 目录。

Windows PowerShell：

```powershell
python "$env:USERPROFILE\.codex\skills\.system\skill-installer\scripts\install-skill-from-github.py" `
  --repo sunfing/agent-skills `
  --ref main `
  --path skills/gpt-image `
  --dest "$env:USERPROFILE\.codex\skills"
```

macOS/Linux：

```shell
python3 ~/.codex/skills/.system/skill-installer/scripts/install-skill-from-github.py \
  --repo sunfing/agent-skills \
  --ref main \
  --path skills/gpt-image \
  --dest ~/.codex/skills
```

目标 Skill 目录已经存在时，安装器会拒绝覆盖。应使用最初安装它的工具进行更新或卸载，不要混用安装管理方式。

## 手工安装

Clone 或下载本仓库，然后只将需要的 `skills/<skill-name>` 目录复制到 Agent 客户端规定的 Skills 目录，并保持 Skill 目录名不变。

Codex 的默认目标目录通常是 Windows 的 `%USERPROFILE%\.codex\skills\<skill-name>`，或 macOS/Linux 的 `~/.codex/skills/<skill-name>`。

## 仓库结构

```text
agent-skills/
├── skills/
│   └── gpt-image/
│       ├── SKILL.md
│       ├── agents/
│       ├── requirements.txt
│       └── scripts/
├── docs/
│   └── gpt-image/
│       ├── README.md
│       └── README.zh-CN.md
└── tests/
```

## 开发验证

当前测试使用 mock，不会调用真实或付费 API：

```shell
python -m unittest discover -s tests -v
```

各 Skill 的具体开发和验证要求以对应文档为准。

## 许可证

[MIT](LICENSE)
