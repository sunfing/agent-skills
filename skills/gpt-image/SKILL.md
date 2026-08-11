---
name: gpt-image
description: Generate or edit bitmap images with OpenAI GPT Image through a user-configured OpenAI-compatible Image API, using the bundled CLI and the fixed gpt-image-2 model. Use when the user invokes $gpt-image, asks to use GPT Image through a third-party relay, or naturally requests bitmap generation or editing while the agent's native image tool is unavailable. Do not use when the user names another Skill, provider, or model, or requests SVG, HTML/CSS, Canvas, Three.js, vector, or other code-drawn output.
---

# GPT Image

Use the bundled `scripts/image_gen.py` directly. Do not load or call a hosted image tool, another image Skill, or a client-internal image script.

## Requirements

- Require Python 3.10+ and the dependencies in `requirements.txt`.
- Read credentials only from `GPT_IMAGE_API_KEY` and `GPT_IMAGE_BASE_URL`.
- Require `GPT_IMAGE_BASE_URL` to be a complete OpenAI-compatible API root ending in `/v1`.
- Never print, persist, or pass the API key as a command-line argument.

## Workflow

1. Preserve the user's prompt exactly unless the user explicitly asks for prompt rewriting.
2. Choose `generate` when there is no input image. Choose `edit` when the user provides one or more reference images.
3. Resolve every user-provided input, mask, and output path to an absolute path.
4. Run exactly one live CLI command. Do not run a dry-run, API preflight, automatic retry, or provider/model fallback.
5. On success, embed each actual output image when the client supports image embedding and report every absolute output path. On failure, report the CLI error and stop.

Generate:

```shell
python "<skill-root>/scripts/image_gen.py" generate --prompt "<prompt>" [--out "<absolute-output-path>"]
```

Edit with one or more reference images:

```shell
python "<skill-root>/scripts/image_gen.py" edit --prompt "<prompt>" --image "<absolute-input-path>" [--image "<another-input-path>"] [--mask "<absolute-mask-path>"] [--out "<absolute-output-path>"]
```

Pass user-requested supported options when present:

- `--size auto|<WIDTH>x<HEIGHT>`
- `--quality auto|low|medium|high`
- `--output-format png|jpeg|webp`
- `--output-compression 0..100` for JPEG or WebP only
- `--background auto|opaque`
- `--moderation auto|low`
- `--n <1..10>`

When `--out` is omitted, let the CLI select the platform Pictures directory and a collision-resistant filename. Never add `--force`; the CLI refuses to overwrite existing files.

## Boundaries

- Use only the fixed API model `gpt-image-2` through `/v1/images/generations` or `/v1/images/edits`.
- Do not pass `input_fidelity`; `gpt-image-2` always processes image inputs at high fidelity.
- Do not request transparent backgrounds; `gpt-image-2` supports only `auto` and `opaque` backgrounds.
- Do not use Responses, streaming partial images, Batch, video, audio, or non-OpenAI image models.
- A mask is optional and applies to the first input image. Require both the mask and first input to be PNG with matching dimensions; the mask must contain an alpha channel.
- Accept at most 16 PNG, JPEG, or WebP input images, each smaller than 50 MB. Require a PNG mask smaller than 50 MB.
- Treat a successful CLI exit as completion. Do not inspect, rewrite, or post-process generated images.
