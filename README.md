# GPT Image Agent Skill

[English](#english)

一个面向 Agent Skills 兼容客户端的非官方社区 Skill。当客户端通过第三方 OpenAI-compatible 中转站而无法使用原生图片工具时，它会直接调用 OpenAI Image API，并使用固定模型 `gpt-image-2` 生成或编辑位图。

本项目与 OpenAI 无隶属关系。它不提供 API 服务、额度或密钥。

## 适用范围

客户端必须能够：

- 读取 Agent Skills 格式的 `SKILL.md`；
- 执行本地 Python 命令；
- 访问网络和本地文件系统；
- 连接实现 `/v1/images/generations` 和 `/v1/images/edits` 的第三方 OpenAI-compatible 中转站。

本 Skill 不适用于不支持本地命令的纯聊天客户端，也不支持其他供应商的生图模型。

## 功能

- 文本生成图片；
- 单图或多参考图编辑；
- PNG mask / inpainting；
- 一次请求生成多张图片；
- 自定义合规尺寸及 `auto/low/medium/high` 质量；
- PNG、JPEG、WebP，以及 JPEG/WebP 压缩；
- `auto/opaque` 背景和 `auto/low` moderation；
- 默认保存到系统 Pictures 目录，且绝不覆盖已有文件。

每次编辑最多接受 16 张 PNG、JPEG 或 WebP 输入图，每张小于 50 MB。使用 mask 时，mask 和第一张输入图都必须是 PNG、尺寸一致且 mask 包含 alpha channel；mask 也必须小于 50 MB。`--n` 的范围是 `1..10`，prompt 最长 32,000 个字符。

`gpt-image-2` 不支持透明背景。该 Skill 也不使用 Responses、streaming partial images、Batch、视频、音频或自动模型 fallback。

## 安装

安装 Python 3.10+，将 `skills/gpt-image` 复制到客户端规定的 Skills 目录，然后安装依赖：

```shell
python -m pip install -r <skills-directory>/gpt-image/requirements.txt
```

Codex 的默认目标目录是：

- Windows：`%USERPROFILE%\.codex\skills\gpt-image`
- macOS/Linux：`~/.codex/skills/gpt-image`

其他客户端应使用各自文档指定的 Skills 目录。安装后重新开启 Agent 任务，使 Skill metadata 生效。

## 配置

不要把密钥写进仓库或 `SKILL.md`。PowerShell 用户级配置：

```powershell
[Environment]::SetEnvironmentVariable("GPT_IMAGE_API_KEY", "your-key", "User")
[Environment]::SetEnvironmentVariable("GPT_IMAGE_BASE_URL", "https://relay.example.com/v1", "User")
```

Bash/Zsh 当前会话配置：

```shell
export GPT_IMAGE_API_KEY="your-key"
export GPT_IMAGE_BASE_URL="https://relay.example.com/v1"
```

`GPT_IMAGE_BASE_URL` 必须是完整 API 根地址并以 `/v1` 结尾。配置持久环境变量后，应重启 Agent 客户端。

## 使用

自然语言请求示例：

```text
$gpt-image 生成一只戴宇航员头盔的橘猫，像素插画风格，纯色背景。
```

```text
$gpt-image 以这两张图片为参考，保留主体服装，把背景改成雨夜街道。
```

Agent 会调用一次仓库自带 CLI。也可以直接运行：

```shell
python skills/gpt-image/scripts/image_gen.py generate --prompt "A pixel-art orange cat wearing an astronaut helmet"
```

```shell
python skills/gpt-image/scripts/image_gen.py edit --prompt "Replace the background with a rainy street" --image input.png --out output.png
```

使用 `--help` 查看所有参数。默认文件名为 `yyyyMMdd-HHmmss-fff-<uuid>.png`。传入 `--n` 大于 1 时，文件名会增加数字后缀。

CLI 每次只发送一次付费生成或编辑请求，显式关闭 SDK 自动重试。若兼容中转站返回图片 URL，保存图片可能额外执行普通下载请求；下载仅允许公网地址或与已配置 relay 相同的主机，并限制为 100 MiB。

## 开发验证

以下测试不会调用真实 API：

```shell
python -m unittest discover -s tests -v
```

API 行为依据 [OpenAI GPT Image 2](https://developers.openai.com/api/docs/models/gpt-image-2) 和 [Image generation](https://developers.openai.com/api/docs/guides/image-generation) 官方文档。

## English

An unofficial Agent Skill for generating and editing bitmap images with the fixed OpenAI `gpt-image-2` model through a user-configured OpenAI-compatible relay. It is intended for skills-compatible agents that can run local Python commands but cannot use a native hosted image tool.

### Requirements

- Python 3.10+
- Local command, network, and filesystem access
- An OpenAI-compatible relay implementing `/v1/images/generations` and `/v1/images/edits`
- `GPT_IMAGE_API_KEY` and a complete `GPT_IMAGE_BASE_URL` ending in `/v1`

### Install

Copy `skills/gpt-image` into the client-specific Skills directory, then install its dependency:

```shell
python -m pip install -r <skills-directory>/gpt-image/requirements.txt
```

For Codex, the default destination is `%USERPROFILE%\.codex\skills\gpt-image` on Windows or `~/.codex/skills/gpt-image` on macOS/Linux. Use the directory documented by your client for other agents, then start a new task so the client reloads the Skill metadata.

### Configure

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

### Use

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

### Test

The test suite uses mocks and does not call a live API:

```shell
python -m unittest discover -s tests -v
```

This project is not affiliated with OpenAI and does not provide an API service, credits, or keys.
