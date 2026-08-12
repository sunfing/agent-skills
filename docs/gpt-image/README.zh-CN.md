# GPT Image Agent Skill

[Skills 总目录](../../README.zh-CN.md) | [English](README.md) | 简体中文

一个非官方 Agent Skill，通过用户配置的 OpenAI-compatible Image API 中转站直接调用固定模型 `gpt-image-2`。它适用于能够执行本地命令，但在第三方中转环境中无法获得原生 Hosted 生图工具的 Agent 客户端。

本项目与 OpenAI 无隶属关系，不提供 API 服务、额度或 API key。

## 支持能力

- 文本生成图片
- 单图编辑与多参考图编辑
- PNG mask / inpainting
- 单次请求生成 `1..10` 张图片
- `gpt-image-2` 接受的合规尺寸
- `auto`、`low`、`medium`、`high` 质量
- PNG、JPEG、WebP 输出及 JPEG/WebP 压缩
- `auto`/`opaque` 背景与 `auto`/`low` moderation
- 默认保存到系统 Pictures 目录，并生成不冲突的文件名
- 每次 CLI 调用只发送一次付费 Image API 请求，并关闭 SDK 自动重试

每次编辑最多接受 16 张 PNG、JPEG 或 WebP 输入图，每张小于 50 MB。使用 mask 时，mask 和第一张输入图都必须是尺寸一致的 PNG，且 mask 必须包含 alpha channel。prompt 最长 32,000 个字符。

`gpt-image-2` 不支持透明背景。本 Skill 也不使用 Responses、partial-image streaming、Batch、视频、音频或自动模型 fallback。

## 使用条件

- Python 3.10+
- 能够读取 `SKILL.md` 并执行本地 Python 命令的 Agent 客户端
- 网络和本地文件系统访问权限
- 同时实现 `/v1/images/generations` 与 `/v1/images/edits` 的 OpenAI-compatible 中转站
- 中转站 API key
- 以 `/v1` 结尾的完整 API 根地址，例如 `https://relay.example.com/v1`

模型固定为 `gpt-image-2`，不需要配置模型环境变量。

## 快速开始

1. 选择 CC Switch、Codex Skill Installer 或手工复制中的一种方式安装。
2. 从安装后的 Skill 目录安装 Python 依赖。
3. 配置全新的 `GPT_IMAGE_API_KEY` 和 `GPT_IMAGE_BASE_URL`。
4. 完全退出并重新启动 Agent 客户端。
5. 使用 `--help` 做无付费验证。
6. 新建任务，首次显式调用 `$gpt-image` 做付费验证。

不要同时让 CC Switch 和手工安装器管理同一份 Skill。CC Switch 维护自己的 source-of-truth 目录，从 CC Switch 卸载 Skill 时可能同时删除同步到 Agent 目录的副本。

## 安装

请选择一种安装方式。

### 方式 A：CC Switch 3.19.2+

CC Switch 会递归发现 `SKILL.md`，因此本仓库的 `skills/gpt-image/SKILL.md` 目录结构可以直接识别。

1. 在 CC Switch 中打开 **Skills**。
2. 点击右上角 **仓库管理**。
3. 点击 **添加仓库**。
4. 输入仓库地址：

   ```text
   https://github.com/sunfing/agent-skills
   ```

5. Branch 填写：

   ```text
   main
   ```

6. 添加仓库后返回 Skills 发现页面，点击 **刷新**。
7. 搜索 `gpt-image`，点击 **安装**，并为 Codex 或其他受支持客户端启用。

注意：Skills 发现页面的搜索框只过滤已经配置的仓库中发现的 Skill。把 GitHub URL 粘贴进搜索框不会添加仓库。

CC Switch 为 Codex 启用 Skill 后，应用侧目录通常是：

```text
~/.codex/skills/gpt-image
```

CC Switch 可能使用符号链接，也可能复制由其自身 Skill 存储目录管理的文件。接下来继续执行[安装依赖](#安装依赖)。

### 方式 B：Codex Skill Installer

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

目标 `gpt-image` 目录已经存在时，安装器会拒绝覆盖。应使用最初安装它的工具进行更新或卸载，不要混用管理方式。

### 方式 C：手工安装

1. 下载或 clone 本仓库。
2. 只将 `skills/gpt-image` 目录复制到 Agent 客户端规定的 Skills 目录。
3. 保持目标目录名为 `gpt-image`。

Codex 的默认目标目录通常是 Windows 的 `%USERPROFILE%\.codex\skills\gpt-image`，或 macOS/Linux 的 `~/.codex/skills/gpt-image`。

### 安装依赖

CC Switch 和手工安装只复制 Skill 文件，不会自动安装 Python package。

Windows PowerShell：

```powershell
python -m pip install -r "$env:USERPROFILE\.codex\skills\gpt-image\requirements.txt"
```

macOS/Linux：

```shell
python3 -m pip install -r ~/.codex/skills/gpt-image/requirements.txt
```

如果使用其他 Agent 客户端，请替换为该客户端实际安装 Skill 的路径。

## 配置

本 Skill 只从以下环境变量读取凭据：

| 变量 | 必须填写的值 |
| --- | --- |
| `GPT_IMAGE_API_KEY` | 中转站提供的有效 key |
| `GPT_IMAGE_BASE_URL` | 以 `/v1` 结尾的完整 HTTPS API 根地址 |

不要把凭据写入本仓库、`SKILL.md` 或 prompt。

### Windows PowerShell：持久用户变量

以下流程会隐藏 key，避免它以明文形式进入 PowerShell 命令历史：

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

只检查变量是否存在，不显示其内容：

```powershell
[bool][Environment]::GetEnvironmentVariable("GPT_IMAGE_API_KEY", "User")
[bool][Environment]::GetEnvironmentVariable("GPT_IMAGE_BASE_URL", "User")
```

两项都应返回 `True`。

### Bash：当前 shell

```shell
read -rsp "Enter GPT_IMAGE_API_KEY: " GPT_IMAGE_API_KEY
printf '\n'
read -rp "Enter GPT_IMAGE_BASE_URL ending in /v1: " GPT_IMAGE_BASE_URL
GPT_IMAGE_BASE_URL="${GPT_IMAGE_BASE_URL%/}"
export GPT_IMAGE_API_KEY GPT_IMAGE_BASE_URL
```

需要从这个 shell 启动 Agent，进程才能继承变量。如需持久化，请使用操作系统和 Agent 客户端推荐的安全环境变量或 secret 管理方式。

### Zsh：当前 shell

```shell
read -rs "GPT_IMAGE_API_KEY?Enter GPT_IMAGE_API_KEY: "
printf '\n'
read "GPT_IMAGE_BASE_URL?Enter GPT_IMAGE_BASE_URL ending in /v1: "
GPT_IMAGE_BASE_URL="${GPT_IMAGE_BASE_URL%/}"
export GPT_IMAGE_API_KEY GPT_IMAGE_BASE_URL
```

需要从这个 shell 启动 Agent，进程才能继承变量。如需持久化，请使用操作系统和 Agent 客户端推荐的安全环境变量或 secret 管理方式。

### 重启 Agent

设置持久环境变量后，应完全退出并重新启动 Agent 客户端。仅新建会话可能不会刷新父进程环境。重启后还应新建任务，让客户端重新加载 Skill metadata。

## 可选：在 Codex 中将 GPT Image 设为默认生图路由

本 Skill 已允许隐式调用：自然语言位图生成或编辑请求与 Skill description 匹配时，即使用户没有输入 `$gpt-image`，Codex 也可以自动选择它。如果 Codex 中还存在其他图像 Skill 或原生图像工具，并希望这些请求优先使用 `$gpt-image`，可在全局 Codex `AGENTS.md` 中增加以下规则。

- Windows：`%USERPROFILE%\.codex\AGENTS.md`
- macOS/Linux：`~/.codex/AGENTS.md`

```markdown
## 默认位图生成路由

- 用户自然语言要求生成或编辑位图时，默认使用 `$gpt-image`，无需用户显式写出 Skill 名称。
- 用户明确指定其他 Skill、provider、model，或要求 SVG、HTML/CSS、Canvas、Three.js、vector、代码绘图时，服从用户指定。
- 除非用户要求改写，否则原样传递用户 prompt。
```

这是可选的 Codex 全局偏好，不是隐式调用的必要条件；用户的明确选择仍然优先。如果 Codex home 目录中存在非空 `AGENTS.override.md`，Codex 会使用它而不是全局 `AGENTS.md`。修改全局规则后，应重启 Codex 并新建任务。具体机制参见 OpenAI 官方 [Skill 调用文档](https://learn.chatgpt.com/docs/build-skills)与 [`AGENTS.md` 文档](https://learn.chatgpt.com/docs/agent-configuration/agents-md)。

## 无付费验证

Windows PowerShell：

```powershell
python "$env:USERPROFILE\.codex\skills\gpt-image\scripts\image_gen.py" --help
```

macOS/Linux：

```shell
python3 ~/.codex/skills/gpt-image/scripts/image_gen.py --help
```

输出应包含 `generate` 和 `edit`。`--help` 不会访问中转站。

## 首次付费验证

在新的 Agent 任务中显式调用：

```text
$gpt-image 生成一只戴宇航员头盔的橘猫，像素插画风格，纯蓝色背景，无文字。
```

这会发送一次付费 Image API 请求。成功后，Agent 应嵌入实际生成图片并报告绝对输出路径。

## 使用方法

### 自然语言生成

```text
$gpt-image 生成一张灰色影棚背景上的白色陶瓷杯产品照片。
```

### 自然语言编辑

附加或明确指定现有图片，然后提出：

```text
$gpt-image 保持主体和构图不变，把背景替换为雨夜街道。
```

### CLI 文生图

```shell
python "/path/to/gpt-image/scripts/image_gen.py" generate \
  --prompt "A pixel-art orange cat wearing an astronaut helmet" \
  --size 1024x1024 \
  --quality medium \
  --output-format png \
  --out output.png
```

### CLI 图片编辑

```shell
python "/path/to/gpt-image/scripts/image_gen.py" edit \
  --prompt "Replace the background with a rainy night street" \
  --image input.png \
  --out edited.png
```

多参考图可重复传入 `--image`：

```shell
python "/path/to/gpt-image/scripts/image_gen.py" edit \
  --prompt "Use the subject from the first image and the color palette from the second" \
  --image subject.png \
  --image palette.png \
  --out combined.png
```

进行 inpainting 时增加 `--mask mask.png`。第一张输入图和 mask 必须是相同尺寸的 PNG，且 mask 必须包含 alpha channel。

未提供 `--out` 时，Skill 会保存到系统 Pictures 目录，文件名格式为 `yyyyMMdd-HHmmss-fff-<uuid>`。已有文件不会被覆盖。

## 常见问题

### CC Switch 显示 0 个匹配的 Skill

- 不要把仓库 URL 粘贴进 Skills 搜索框。
- 通过右上角 **仓库管理** 添加 URL。
- 确认仓库列表存在 `sunfing/agent-skills`，Branch 为 `main`。
- 返回 Skills 发现页面并点击 **刷新**。
- 搜索 `gpt-image`，不要搜索完整 URL。

### `error: GPT_IMAGE_API_KEY is not set.`

Agent 进程没有获得 `GPT_IMAGE_API_KEY`。确认变量已经设置，然后完全重启 Agent 应用。已经运行的桌面进程内部新建任务，仍可能继承旧环境。

### `error: GPT_IMAGE_BASE_URL must use HTTPS and be an absolute URL.`

中转站 URL 必须使用 `https://`；HTTP 明文传输会泄露 API key、prompt 和输入图片。

### `error: GPT_IMAGE_BASE_URL must be a complete API root ending in /v1.`

应填写完整 API 根地址，例如 `https://relay.example.com/v1`，不能填写网站首页，也不能填写完整的 `/images/generations` endpoint。

### Skill 已安装但没有触发

- 确认 `<skills-directory>/gpt-image/SKILL.md` 存在。
- 完全重启 Agent 并新建任务。
- 首次测试显式调用 `$gpt-image`。
- 如果客户端自带原生生图工具，其路由策略可能优先于 Skill 的隐式调用。

### 中转站拒绝生成或编辑请求

确认中转站在 `/v1/images/generations` 和 `/v1/images/edits` 上都支持 `gpt-image-2`。仅兼容 Chat Completions 或 Responses 不能证明它兼容 Image API。

### 安装管理方式冲突

如果由 CC Switch 管理，请从 CC Switch 更新或卸载。如果由 Codex Skill Installer 管理，不要在尚未确定管理归属时再次把它导入 CC Switch。

## 更新

- CC Switch：点击 **刷新**，出现更新提示后使用 Skill 的 **更新** 操作。
- Codex Skill Installer/手工复制：先备份本地自定义内容，再有意删除旧 `gpt-image` 目录并从 `main` 重新安装。API 凭据是环境变量，不应存放在 Skill 目录中。

## 开发验证

测试使用 mock，不会调用真实 API：

```shell
python -m unittest discover -s tests -v
```

API 行为依据 OpenAI 官方 [GPT Image 2](https://developers.openai.com/api/docs/models/gpt-image-2) 与 [Image generation](https://developers.openai.com/api/docs/guides/image-generation) 文档。CC Switch 安装流程依据其 [v3.19.2 Skills 中文手册](https://github.com/farion1231/cc-switch/blob/v3.19.2/docs/user-manual/zh/3-extensions/3.3-skills.md) 核对。
