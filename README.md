# Agent Skills

English | [简体中文](README.zh-CN.md)

A collection of unofficial, self-contained Skills for Agent clients that support the `SKILL.md` format and local command execution.

Each Skill lives under [`skills/`](skills/) and includes only the instructions and resources required by the Agent. User-facing installation, configuration, and usage guides live under [`docs/`](docs/).

This repository and its Skills are not affiliated with OpenAI, Anthropic, or any other model provider. They do not provide API services, credits, or API keys.

## Available Skills

| Skill | Purpose | Documentation |
| --- | --- | --- |
| [`gpt-image`](skills/gpt-image/) | Generate and edit bitmap images with the fixed `gpt-image-2` model through a user-configured OpenAI-compatible Image API. | [English](docs/gpt-image/README.md) · [简体中文](docs/gpt-image/README.zh-CN.md) |

## Install with CC Switch

CC Switch 3.19.2+ recursively discovers `SKILL.md` files in a repository.

1. Open **Skills** in CC Switch.
2. Click **Repository Manager** (`仓库管理`) in the upper-right corner.
3. Click **Add Repository** (`添加仓库`).
4. Enter `https://github.com/sunfing/agent-skills` and select branch `main`.
5. Return to the Skills discovery page and click **Refresh**.
6. Search for the Skill name, then install and enable it for the required Agent client.

The discovery-page search box only filters Skills from configured repositories. It does not add a repository when a GitHub URL is pasted into it.

After installation, follow the selected Skill's documentation for dependencies, environment variables, and verification.

## Install one Skill with Codex

Use Codex's bundled Skill Installer and pass the directory of the required Skill.

Windows PowerShell:

```powershell
python "$env:USERPROFILE\.codex\skills\.system\skill-installer\scripts\install-skill-from-github.py" `
  --repo sunfing/agent-skills `
  --ref main `
  --path skills/gpt-image `
  --dest "$env:USERPROFILE\.codex\skills"
```

macOS/Linux:

```shell
python3 ~/.codex/skills/.system/skill-installer/scripts/install-skill-from-github.py \
  --repo sunfing/agent-skills \
  --ref main \
  --path skills/gpt-image \
  --dest ~/.codex/skills
```

The installer refuses to overwrite an existing Skill directory. Update or remove an existing installation through the tool that originally installed it instead of mixing installation ownership.

## Manual installation

Clone or download this repository, then copy only the required `skills/<skill-name>` directory into the Skills directory documented by your Agent client. Preserve the Skill directory name.

For Codex, the destination is normally `%USERPROFILE%\.codex\skills\<skill-name>` on Windows or `~/.codex/skills/<skill-name>` on macOS/Linux.

## Repository layout

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

## Development

The current test suite uses mocks and does not call a live or paid API:

```shell
python -m unittest discover -s tests -v
```

See each Skill's documentation for its exact development and validation requirements.

## License

[MIT](LICENSE)
