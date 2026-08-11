# GPT Image Agent Skill

English | [简体中文](README.zh-CN.md)

An unofficial Agent Skill for generating and editing bitmap images with the fixed OpenAI `gpt-image-2` model through a user-configured OpenAI-compatible relay. It is intended for skills-compatible agents that can run local Python commands but cannot use a native hosted image tool.

## Requirements

- Python 3.10+
- Local command, network, and filesystem access
- An OpenAI-compatible relay implementing `/v1/images/generations` and `/v1/images/edits`
- `GPT_IMAGE_API_KEY` and a complete `GPT_IMAGE_BASE_URL` ending in `/v1`

## Install

Copy `skills/gpt-image` into the client-specific Skills directory, then install its dependency:

```shell
python -m pip install -r <skills-directory>/gpt-image/requirements.txt
```

For Codex, the default destination is `%USERPROFILE%\.codex\skills\gpt-image` on Windows or `~/.codex/skills/gpt-image` on macOS/Linux. Use the directory documented by your client for other agents, then start a new task so the client reloads the Skill metadata.

## Configure

PowerShell user-level environment:

```powershell
[Environment]::SetEnvironmentVariable("GPT_IMAGE_API_KEY", "your-key", "User")
[Environment]::SetEnvironmentVariable("GPT_IMAGE_BASE_URL", "https://relay.example.com/v1", "User")
```

Bash/Zsh current session:

```shell
export GPT_IMAGE_API_KEY="your-key"
export GPT_IMAGE_BASE_URL="https://relay.example.com/v1"
```

Restart the agent after setting persistent environment variables. Never put credentials in this repository or in `SKILL.md`.

## Use

Invoke the Skill naturally:

```text
$gpt-image Generate a pixel-art orange cat wearing an astronaut helmet on a solid background.
```

Or run the bundled CLI directly:

```shell
python skills/gpt-image/scripts/image_gen.py generate --prompt "A pixel-art orange cat wearing an astronaut helmet"
```

```shell
python skills/gpt-image/scripts/image_gen.py edit --prompt "Replace the background with a rainy street" --image input.png --out output.png
```

Run `python skills/gpt-image/scripts/image_gen.py --help` for the command surface. Each edit accepts at most 16 PNG, JPEG, or WebP inputs under 50 MB each. With a mask, both the mask and first input must be PNG with matching dimensions, and the mask must contain an alpha channel. The `--n` range is `1..10`, and prompts are limited to 32,000 characters. The CLI makes one paid Image API request, disables SDK retries, preserves prompts by default, refuses overwrites, and never stores credentials. If a compatible relay returns an image URL instead of base64 data, saving the artifact may require an additional ordinary download limited to 100 MiB and restricted to public hosts or the configured relay host.

## Test

The test suite uses mocks and does not call a live API:

```shell
python -m unittest discover -s tests -v
```

API behavior follows the official [OpenAI GPT Image 2](https://developers.openai.com/api/docs/models/gpt-image-2) and [Image generation](https://developers.openai.com/api/docs/guides/image-generation) documentation.

This project is not affiliated with OpenAI and does not provide an API service, credits, or keys.
