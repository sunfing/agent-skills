# GPT Image Agent Skill

[Skills catalog](../../README.md) | English | [简体中文](README.zh-CN.md)

An unofficial Agent Skill that calls the OpenAI-compatible Image API directly with the fixed `gpt-image-2` model. It is intended for Agent clients that can run local commands but cannot expose a native hosted image tool when connected through a third-party relay.

This project is not affiliated with OpenAI. It does not provide an API service, credits, or API keys.

## What it supports

- Text-to-image generation
- Single-image and multi-reference editing
- PNG masks / inpainting
- `1..10` outputs per request
- Supported sizes accepted by `gpt-image-2`
- `auto`, `low`, `medium`, and `high` quality
- PNG, JPEG, and WebP output, including JPEG/WebP compression
- `auto` or `opaque` backgrounds and `auto` or `low` moderation
- Platform Pictures directory defaults with collision-safe filenames
- Exactly one paid Image API request per CLI invocation, with SDK retries disabled

Each edit accepts at most 16 PNG, JPEG, or WebP input images under 50 MB each. With a mask, the mask and first input must both be PNG files with matching dimensions, and the mask must contain an alpha channel. Prompts are limited to 32,000 characters.

`gpt-image-2` does not support transparent backgrounds. This Skill does not use Responses, partial-image streaming, Batch, video, audio, or automatic model fallback.

## Requirements

- Python 3.10+
- An Agent client that can read `SKILL.md` and execute local Python commands
- Network and local filesystem access
- An OpenAI-compatible relay implementing both `/v1/images/generations` and `/v1/images/edits`
- A relay API key
- A complete relay API root ending in `/v1`, for example `https://relay.example.com/v1`

The model is fixed to `gpt-image-2`; there is no model environment variable to configure.

## Quick start

1. Install the Skill with CC Switch, the Codex Skill Installer, or a manual copy.
2. Install the Python dependency from the installed Skill directory.
3. Set fresh `GPT_IMAGE_API_KEY` and `GPT_IMAGE_BASE_URL` environment variables.
4. Fully quit and restart the Agent client.
5. Run the no-cost `--help` check.
6. Start a new task and explicitly invoke `$gpt-image` for the first paid test.

Do not manage the same installation with both CC Switch and a manual installer. CC Switch maintains its own source-of-truth directory and may remove synchronized copies when a Skill is uninstalled there.

## Install

Choose one installation method.

### Option A: CC Switch 3.19.2+

CC Switch recursively discovers `SKILL.md`, so this repository's `skills/gpt-image/SKILL.md` layout is supported.

1. Open **Skills** in CC Switch.
2. Click **Repository Manager** (`仓库管理`) in the upper-right corner.
3. Click **Add Repository** (`添加仓库`).
4. Enter this repository URL:

   ```text
   https://github.com/sunfing/agent-skills
   ```

5. Set the branch to:

   ```text
   main
   ```

6. Add the repository, return to the Skills discovery page, and click **Refresh**.
7. Search for `gpt-image`, click **Install**, and enable it for Codex or another supported client.

Important: the search box on the Skills discovery page only filters Skills from repositories that are already configured. Pasting a GitHub URL into that search box does not add a repository.

After CC Switch enables the Skill for Codex, its application path is normally:

```text
~/.codex/skills/gpt-image
```

CC Switch may use a symlink or a copied directory backed by its own Skill storage. Continue with [Install the dependency](#install-the-dependency).

### Option B: Codex Skill Installer

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

The installer refuses to overwrite an existing `gpt-image` directory. Update or remove the existing installation through the tool that originally installed it instead of mixing ownership.

### Option C: Manual installation

1. Download or clone this repository.
2. Copy only the `skills/gpt-image` directory into the Skills directory documented by your Agent client.
3. Preserve the directory name `gpt-image`.

For Codex, the destination is normally `%USERPROFILE%\.codex\skills\gpt-image` on Windows or `~/.codex/skills/gpt-image` on macOS/Linux.

### Install the dependency

CC Switch and manual Skill installation copy files but do not install Python packages.

Windows PowerShell:

```powershell
python -m pip install -r "$env:USERPROFILE\.codex\skills\gpt-image\requirements.txt"
```

macOS/Linux:

```shell
python3 -m pip install -r ~/.codex/skills/gpt-image/requirements.txt
```

For another Agent client, replace the path with that client's actual installed Skill directory.

## Configure

The Skill reads credentials only from these environment variables:

| Variable | Required value |
| --- | --- |
| `GPT_IMAGE_API_KEY` | A valid key for your relay |
| `GPT_IMAGE_BASE_URL` | A complete HTTPS API root ending in `/v1` |

Never write credentials into this repository, `SKILL.md`, or a prompt.

### Windows PowerShell: persistent user variables

This prompt hides the key so it is not stored as plaintext in PowerShell command history:

```powershell
$gptImageSecureKey = Read-Host "Enter GPT_IMAGE_API_KEY" -AsSecureString
$gptImageKey = [System.Net.NetworkCredential]::new("", $gptImageSecureKey).Password
$gptImageBaseUrl = (Read-Host "Enter GPT_IMAGE_BASE_URL ending in /v1").Trim().TrimEnd("/")

if ($gptImageBaseUrl -notmatch '^https://.+/v1$') {
    throw "GPT_IMAGE_BASE_URL must be a complete HTTPS API root ending in /v1."
}

[Environment]::SetEnvironmentVariable("GPT_IMAGE_API_KEY", $gptImageKey, "User")
[Environment]::SetEnvironmentVariable("GPT_IMAGE_BASE_URL", $gptImageBaseUrl, "User")

Remove-Variable gptImageSecureKey, gptImageKey, gptImageBaseUrl
```

Confirm that both variables exist without displaying their values:

```powershell
[bool][Environment]::GetEnvironmentVariable("GPT_IMAGE_API_KEY", "User")
[bool][Environment]::GetEnvironmentVariable("GPT_IMAGE_BASE_URL", "User")
```

Both commands should return `True`.

### Bash: current shell

```shell
read -rsp "Enter GPT_IMAGE_API_KEY: " GPT_IMAGE_API_KEY
printf '\n'
read -rp "Enter GPT_IMAGE_BASE_URL ending in /v1: " GPT_IMAGE_BASE_URL
GPT_IMAGE_BASE_URL="${GPT_IMAGE_BASE_URL%/}"
export GPT_IMAGE_API_KEY GPT_IMAGE_BASE_URL
```

Launch the Agent from this shell so it inherits the variables. For persistent configuration, use the secure environment or secret-management mechanism recommended by your operating system and Agent client.

### Zsh: current shell

```shell
read -rs "GPT_IMAGE_API_KEY?Enter GPT_IMAGE_API_KEY: "
printf '\n'
read "GPT_IMAGE_BASE_URL?Enter GPT_IMAGE_BASE_URL ending in /v1: "
GPT_IMAGE_BASE_URL="${GPT_IMAGE_BASE_URL%/}"
export GPT_IMAGE_API_KEY GPT_IMAGE_BASE_URL
```

Launch the Agent from this shell so it inherits the variables. For persistent configuration, use the secure environment or secret-management mechanism recommended by your operating system and Agent client.

### Restart the Agent

Fully quit and restart the Agent client after setting persistent variables. Opening only a new conversation may not reload the parent process environment. Start a new task after the restart so the client also reloads Skill metadata.

## Optional: Make GPT Image the default in Codex

This Skill already permits implicit invocation: Codex can select it when a natural-language bitmap generation or editing request matches the Skill description, even when the user does not type `$gpt-image`. If Codex also has another image Skill or a native image tool and you want these requests to prefer `$gpt-image`, add the following guidance to the global Codex `AGENTS.md`.

- Windows: `%USERPROFILE%\.codex\AGENTS.md`
- macOS/Linux: `~/.codex/AGENTS.md`

```markdown
## Default bitmap generation route

- For natural-language bitmap generation or editing requests, use `$gpt-image` by default even when the user does not explicitly name it.
- Follow explicit requests for another Skill, provider, model, or SVG, HTML/CSS, Canvas, Three.js, vector, or other code-drawn output.
- Pass the user's prompt unchanged unless the user asks for rewriting.
```

This is an optional Codex-wide preference, not a requirement for implicit invocation. Explicit user choices still take precedence. If a non-empty `AGENTS.override.md` exists in the Codex home directory, Codex uses it instead of the global `AGENTS.md`. Restart Codex and start a new task after changing global guidance. See the official OpenAI documentation for [Skill invocation](https://learn.chatgpt.com/docs/build-skills) and [`AGENTS.md`](https://learn.chatgpt.com/docs/agent-configuration/agents-md).

## Verify without a paid request

Windows PowerShell:

```powershell
python "$env:USERPROFILE\.codex\skills\gpt-image\scripts\image_gen.py" --help
```

macOS/Linux:

```shell
python3 ~/.codex/skills/gpt-image/scripts/image_gen.py --help
```

The output should list the `generate` and `edit` commands. `--help` does not contact the relay.

## First paid test

In a new Agent task, invoke the Skill explicitly:

```text
$gpt-image Generate a pixel-art orange cat wearing an astronaut helmet on a solid blue background, with no text.
```

This sends one paid Image API request. On success, the Agent should embed the generated image and report its absolute output path.

## Use

### Natural-language generation

```text
$gpt-image Generate a product photo of a white ceramic mug on a gray studio background.
```

### Natural-language editing

Attach or identify an existing image, then ask:

```text
$gpt-image Keep the subject and composition unchanged, but replace the background with a rainy night street.
```

### Direct CLI generation

```shell
python "/path/to/gpt-image/scripts/image_gen.py" generate \
  --prompt "A pixel-art orange cat wearing an astronaut helmet" \
  --size 1024x1024 \
  --quality medium \
  --output-format png \
  --out output.png
```

### Direct CLI editing

```shell
python "/path/to/gpt-image/scripts/image_gen.py" edit \
  --prompt "Replace the background with a rainy night street" \
  --image input.png \
  --out edited.png
```

Repeat `--image` for multiple reference images:

```shell
python "/path/to/gpt-image/scripts/image_gen.py" edit \
  --prompt "Use the subject from the first image and the color palette from the second" \
  --image subject.png \
  --image palette.png \
  --out combined.png
```

Add `--mask mask.png` for inpainting. The first input and mask must be same-size PNG files, and the mask must contain an alpha channel.

If `--out` is omitted, the Skill saves to the platform Pictures directory with a `yyyyMMdd-HHmmss-fff-<uuid>` filename. Existing files are never overwritten.

## Troubleshooting

### CC Switch shows zero matching Skills

- Do not paste the repository URL into the Skills search box.
- Add the URL through **Repository Manager** (`仓库管理`).
- Confirm the repository is listed as `sunfing/agent-skills` with branch `main`.
- Return to Skills discovery and click **Refresh**.
- Search for `gpt-image`, not the complete URL.

### `error: GPT_IMAGE_API_KEY is not set.`

The Agent process did not receive `GPT_IMAGE_API_KEY`. Confirm that the variable exists, then fully restart the Agent application. A new task inside an already-running desktop process may still use the old environment.

### `error: GPT_IMAGE_BASE_URL must use HTTPS and be an absolute URL.`

Use an `https://` relay URL. Plain HTTP would expose the API key, prompt, and input images in transit.

### `error: GPT_IMAGE_BASE_URL must be a complete API root ending in /v1.`

Use the complete API root, such as `https://relay.example.com/v1`, not the website homepage and not the full `/images/generations` endpoint.

### The Skill is installed but does not trigger

- Confirm `<skills-directory>/gpt-image/SKILL.md` exists.
- Fully restart the Agent and start a new task.
- Invoke `$gpt-image` explicitly for the first test.
- If the client provides its own native image tool, its routing policy may take precedence over implicit Skill invocation.

### The relay rejects generation or editing

Confirm that the relay supports `gpt-image-2` on both `/v1/images/generations` and `/v1/images/edits`. Chat Completions or Responses compatibility alone does not prove Image API compatibility.

### Installation ownership conflicts

If CC Switch manages the Skill, update or uninstall it from CC Switch. If the Codex Skill Installer manages it, do not also import it into CC Switch without first deciding which tool owns the installation.

## Updating

- CC Switch: click **Refresh**, then use the Skill's **Update** action when shown.
- Codex Skill Installer/manual copy: back up any local modifications, remove the old `gpt-image` directory intentionally, and reinstall from `main`. API credentials are environment variables and should never be stored inside the Skill directory.

## Development

The test suite uses mocks and does not call a live API:

```shell
python -m unittest discover -s tests -v
```

API behavior follows the official [OpenAI GPT Image 2](https://developers.openai.com/api/docs/models/gpt-image-2) and [Image generation](https://developers.openai.com/api/docs/guides/image-generation) documentation. CC Switch installation behavior was checked against its [v3.19.2 Skills manual](https://github.com/farion1231/cc-switch/blob/v3.19.2/docs/user-manual/en/3-extensions/3.3-skills.md).
