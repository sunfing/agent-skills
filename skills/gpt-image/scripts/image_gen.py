#!/usr/bin/env python3
"""Generate or edit images through an OpenAI-compatible Image API."""

from __future__ import annotations

import argparse
import base64
import ctypes
from contextlib import ExitStack
from datetime import datetime
import ipaddress
import json
import os
from pathlib import Path
import re
import socket
import sys
from typing import Any, Callable, Sequence
from urllib.parse import urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener
import uuid

try:
    from openai import OpenAI
except ImportError:  # pragma: no cover - exercised only without dependencies
    OpenAI = None  # type: ignore[assignment]


MODEL = "gpt-image-2"
API_KEY_ENV = "GPT_IMAGE_API_KEY"
BASE_URL_ENV = "GPT_IMAGE_BASE_URL"
REQUEST_TIMEOUT_SECONDS = 300.0
MAX_DOWNLOAD_BYTES = 100 * 1024 * 1024
MAX_INPUT_BYTES = 50_000_000
MAX_MASK_BYTES = 50_000_000
MAX_INPUT_IMAGES = 16
MAX_OUTPUT_IMAGES = 10
MAX_PROMPT_CHARACTERS = 32_000
MAX_ERROR_DETAIL_CHARACTERS = 4_000
PICTURES_FOLDER_ID = uuid.UUID("33e28130-4e1e-4676-835a-98395c3bc3bb")
INPUT_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}
FORMAT_EXTENSIONS = {
    "png": {".png"},
    "jpeg": {".jpg", ".jpeg"},
    "webp": {".webp"},
}


class CliError(RuntimeError):
    """A user-correctable CLI error."""


def _image_count(value: str) -> int:
    parsed = int(value)
    if not 1 <= parsed <= MAX_OUTPUT_IMAGES:
        raise argparse.ArgumentTypeError(
            f"must be between 1 and {MAX_OUTPUT_IMAGES}"
        )
    return parsed


def _compression(value: str) -> int:
    parsed = int(value)
    if not 0 <= parsed <= 100:
        raise argparse.ArgumentTypeError("must be between 0 and 100")
    return parsed


def _size(value: str) -> str:
    if value == "auto":
        return value

    match = re.fullmatch(r"(\d+)x(\d+)", value)
    if not match:
        raise argparse.ArgumentTypeError("must be 'auto' or WIDTHxHEIGHT")

    width, height = (int(part) for part in match.groups())
    long_edge, short_edge = max(width, height), min(width, height)
    pixels = width * height
    if width % 16 or height % 16:
        raise argparse.ArgumentTypeError("both dimensions must be multiples of 16")
    if long_edge > 3840:
        raise argparse.ArgumentTypeError("the longest edge must be at most 3840")
    if long_edge > short_edge * 3:
        raise argparse.ArgumentTypeError("the aspect ratio must not exceed 3:1")
    if not 655_360 <= pixels <= 8_294_400:
        raise argparse.ArgumentTypeError(
            "total pixels must be between 655360 and 8294400"
        )
    return value


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=f"Generate or edit images with the fixed {MODEL} model."
    )
    subparsers = parser.add_subparsers(dest="operation", required=True)

    def add_common_arguments(command: argparse.ArgumentParser) -> None:
        command.add_argument("--prompt", required=True)
        command.add_argument("--out")
        command.add_argument("--size", type=_size)
        command.add_argument(
            "--quality", choices=("auto", "low", "medium", "high")
        )
        command.add_argument(
            "--output-format", choices=("png", "jpeg", "webp")
        )
        command.add_argument("--output-compression", type=_compression)
        command.add_argument("--background", choices=("auto", "opaque"))
        command.add_argument("--moderation", choices=("auto", "low"))
        command.add_argument("--n", type=_image_count)

    generate = subparsers.add_parser("generate", help="Generate images from text")
    add_common_arguments(generate)

    edit = subparsers.add_parser("edit", help="Edit or reference existing images")
    add_common_arguments(edit)
    edit.add_argument("--image", action="append", required=True)
    edit.add_argument("--mask")
    return parser


def _validate_base_url(value: str) -> str:
    normalized = value.rstrip("/")
    parsed = urlsplit(normalized)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise CliError(f"{BASE_URL_ENV} must be an absolute HTTP(S) URL.")
    if parsed.username or parsed.password or parsed.query or parsed.fragment:
        raise CliError(
            f"{BASE_URL_ENV} must not contain credentials, a query, or a fragment."
        )
    if not parsed.path.rstrip("/").endswith("/v1"):
        raise CliError(f"{BASE_URL_ENV} must be a complete API root ending in /v1.")
    return normalized


def _load_config() -> tuple[str, str]:
    api_key = os.environ.get(API_KEY_ENV, "").strip()
    base_url = os.environ.get(BASE_URL_ENV, "").strip()
    if not api_key:
        raise CliError(f"{API_KEY_ENV} is not set.")
    if not base_url:
        raise CliError(f"{BASE_URL_ENV} is not set.")
    return api_key, _validate_base_url(base_url)


def _windows_pictures_dir() -> Path:
    class GUID(ctypes.Structure):
        _fields_ = [
            ("Data1", ctypes.c_uint32),
            ("Data2", ctypes.c_uint16),
            ("Data3", ctypes.c_uint16),
            ("Data4", ctypes.c_ubyte * 8),
        ]

    folder_id = GUID(
        PICTURES_FOLDER_ID.time_low,
        PICTURES_FOLDER_ID.time_mid,
        PICTURES_FOLDER_ID.time_hi_version,
        (ctypes.c_ubyte * 8)(*PICTURES_FOLDER_ID.bytes[8:]),
    )
    path_pointer = ctypes.c_wchar_p()
    shell32 = ctypes.windll.shell32  # type: ignore[attr-defined]
    ole32 = ctypes.windll.ole32  # type: ignore[attr-defined]
    shell32.SHGetKnownFolderPath.argtypes = [
        ctypes.POINTER(GUID),
        ctypes.c_uint32,
        ctypes.c_void_p,
        ctypes.POINTER(ctypes.c_wchar_p),
    ]
    shell32.SHGetKnownFolderPath.restype = ctypes.c_long
    ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    ole32.CoTaskMemFree.restype = None
    result = shell32.SHGetKnownFolderPath(
        ctypes.byref(folder_id), 0, None, ctypes.byref(path_pointer)
    )
    if result != 0 or not path_pointer.value:
        raise OSError(f"SHGetKnownFolderPath failed with HRESULT {result}")
    try:
        return Path(path_pointer.value)
    finally:
        ole32.CoTaskMemFree(ctypes.cast(path_pointer, ctypes.c_void_p))


def _linux_pictures_dir(home: Path) -> Path:
    config_home = Path(os.environ.get("XDG_CONFIG_HOME", home / ".config"))
    config_path = config_home / "user-dirs.dirs"
    try:
        lines = config_path.read_text(encoding="utf-8").splitlines()
    except OSError:
        return home / "Pictures"

    for line in lines:
        if not line.startswith("XDG_PICTURES_DIR="):
            continue
        value = line.split("=", 1)[1].strip().strip('"')
        value = value.replace("$HOME", str(home))
        return Path(os.path.expandvars(value)).expanduser()
    return home / "Pictures"


def _pictures_dir() -> Path:
    home = Path.home()
    if os.name == "nt":
        try:
            return _windows_pictures_dir()
        except OSError:
            return home / "Pictures"
    if sys.platform == "darwin":
        return home / "Pictures"
    return _linux_pictures_dir(home)


def _default_output_path(output_format: str) -> Path:
    now = datetime.now()
    timestamp = f"{now:%Y%m%d-%H%M%S}-{now.microsecond // 1000:03d}"
    filename = f"{timestamp}-{uuid.uuid4().hex}.{output_format}"
    return (_pictures_dir() / filename).resolve()


def _resolve_output_path(value: str | None, output_format: str) -> Path:
    if value is None:
        return _default_output_path(output_format)

    output = Path(value).expanduser()
    if not output.is_absolute():
        output = Path.cwd() / output
    output = output.resolve()
    if not output.suffix:
        output = output.with_suffix(f".{output_format}")
    if output.suffix.lower() not in FORMAT_EXTENSIONS[output_format]:
        allowed = ", ".join(sorted(FORMAT_EXTENSIONS[output_format]))
        raise CliError(
            f"Output extension {output.suffix!r} does not match {output_format}; "
            f"expected one of: {allowed}."
        )
    return output


def _candidate_paths(output: Path, count: int) -> list[Path]:
    if count == 1:
        return [output]
    return [
        output.with_name(f"{output.stem}-{index}{output.suffix}")
        for index in range(1, count + 1)
    ]


def _require_new_paths(paths: Sequence[Path]) -> None:
    existing = next((path for path in paths if path.exists()), None)
    if existing is not None:
        raise CliError(f"Refusing to overwrite existing output: {existing}")


def _resolve_input_path(value: str, label: str, *, mask: bool = False) -> Path:
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = Path.cwd() / path
    path = path.resolve()
    if not path.is_file():
        raise CliError(f"{label} is not an existing file: {path}")
    allowed_extensions = {".png"} if mask else INPUT_EXTENSIONS
    if path.suffix.lower() not in allowed_extensions:
        allowed = ", ".join(sorted(allowed_extensions))
        raise CliError(f"{label} must use one of these formats: {allowed}.")
    size_limit = MAX_MASK_BYTES if mask else MAX_INPUT_BYTES
    if path.stat().st_size >= size_limit:
        size_mb = size_limit // 1_000_000
        raise CliError(f"{label} must be smaller than {size_mb} MB: {path}")
    return path


def _validate_download_url(
    url: str, allowed_private_origin: tuple[str, int] | None
) -> None:
    parsed = urlsplit(url)
    if parsed.scheme not in {"http", "https"}:
        raise CliError("The API returned an image URL with an unsupported scheme.")
    if not parsed.hostname:
        raise CliError("The API returned an image URL without a hostname.")
    if parsed.username or parsed.password:
        raise CliError("The API returned an image URL containing credentials.")
    port = parsed.port or (443 if parsed.scheme == "https" else 80)
    if (parsed.hostname, port) == allowed_private_origin:
        return
    try:
        addresses = socket.getaddrinfo(
            parsed.hostname,
            port,
            type=socket.SOCK_STREAM,
        )
    except OSError as error:
        raise CliError("The API returned an image URL whose host cannot be resolved.") from error
    if not addresses:
        raise CliError("The API returned an image URL whose host cannot be resolved.")
    for address in addresses:
        ip = ipaddress.ip_address(address[4][0])
        if not ip.is_global:
            raise CliError("Refusing to download an image URL from a non-public host.")


def _download_url(
    url: str, allowed_private_origin: tuple[str, int] | None = None
) -> bytes:
    _validate_download_url(url, allowed_private_origin)

    class ValidatingRedirectHandler(HTTPRedirectHandler):
        def redirect_request(self, request, file_pointer, code, message, headers, new_url):
            _validate_download_url(new_url, allowed_private_origin)
            return super().redirect_request(
                request, file_pointer, code, message, headers, new_url
            )

    request = Request(url, headers={"User-Agent": "gpt-image-agent-skill/0.1"})
    opener = build_opener(ValidatingRedirectHandler())
    with opener.open(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
        content_length = response.headers.get("Content-Length")
        if content_length and int(content_length) > MAX_DOWNLOAD_BYTES:
            raise CliError("The returned image URL exceeds the 100 MiB download limit.")
        payload = bytearray()
        while chunk := response.read(1024 * 1024):
            payload.extend(chunk)
            if len(payload) > MAX_DOWNLOAD_BYTES:
                raise CliError("The returned image URL exceeds the 100 MiB download limit.")
        return bytes(payload)


def _item_value(item: Any, name: str) -> Any:
    if isinstance(item, dict):
        return item.get(name)
    return getattr(item, name, None)


def _response_bytes(
    response: Any,
    downloader: Callable[[str], bytes],
    expected_count: int | None,
) -> list[bytes]:
    items = getattr(response, "data", None)
    if not items:
        raise CliError("The API returned no image data.")
    actual_count = len(items)
    if actual_count > MAX_OUTPUT_IMAGES:
        raise CliError(
            f"The API returned {actual_count} images; at most "
            f"{MAX_OUTPUT_IMAGES} are accepted."
        )
    if expected_count is not None and actual_count != expected_count:
        raise CliError(
            f"The API returned {actual_count} images; expected {expected_count}."
        )

    images: list[bytes] = []
    for item in items:
        encoded = _item_value(item, "b64_json")
        if encoded:
            try:
                images.append(base64.b64decode(encoded, validate=True))
            except (ValueError, TypeError) as error:
                raise CliError("The API returned invalid base64 image data.") from error
            continue
        url = _item_value(item, "url")
        if url:
            images.append(downloader(url))
            continue
        raise CliError("The API response contained neither b64_json nor url image data.")
    return images


def _write_outputs(paths: Sequence[Path], images: Sequence[bytes]) -> list[Path]:
    if len(images) != len(paths):
        raise CliError("The number of output paths does not match the image data.")

    selected_paths = list(paths)
    _require_new_paths(selected_paths)
    created: list[Path] = []
    try:
        for path, image in zip(selected_paths, images, strict=True):
            with path.open("xb") as output_file:
                created.append(path)
                output_file.write(image)
    except Exception:
        for path in created:
            path.unlink(missing_ok=True)
        raise
    return created


def execute(
    args: argparse.Namespace,
    *,
    client_factory: Callable[..., Any] | None = None,
    downloader: Callable[[str], bytes] | None = None,
) -> list[Path]:
    api_key, base_url = _load_config()
    output_format = args.output_format or "png"
    expected_count = args.n
    if not args.prompt or len(args.prompt) > MAX_PROMPT_CHARACTERS:
        raise CliError(
            f"--prompt must contain between 1 and {MAX_PROMPT_CHARACTERS} characters."
        )
    if args.output_compression is not None and output_format == "png":
        raise CliError("--output-compression is supported only for jpeg or webp.")

    output = _resolve_output_path(args.out, output_format)
    preflight_paths = _candidate_paths(output, expected_count or 1)
    _require_new_paths(preflight_paths)

    image_paths: list[Path] = []
    mask_path: Path | None = None
    if args.operation == "edit":
        if len(args.image) > MAX_INPUT_IMAGES:
            raise CliError(f"At most {MAX_INPUT_IMAGES} input images are supported.")
        image_paths = [_resolve_input_path(value, "Input image") for value in args.image]
        if args.mask:
            mask_path = _resolve_input_path(args.mask, "Mask", mask=True)
            if image_paths[0].suffix.lower() != ".png":
                raise CliError("The first input image must be PNG when using a mask.")

    output.parent.mkdir(parents=True, exist_ok=True)

    factory = client_factory or OpenAI
    if factory is None:
        raise CliError(
            "The openai package is not installed. Install dependencies from requirements.txt."
        )
    client = factory(
        api_key=api_key,
        base_url=base_url,
        max_retries=0,
        timeout=REQUEST_TIMEOUT_SECONDS,
    )
    request: dict[str, Any] = {
        "model": MODEL,
        "prompt": args.prompt,
    }
    for name in (
        "size",
        "quality",
        "output_format",
        "output_compression",
        "background",
        "n",
    ):
        value = getattr(args, name)
        if value is not None:
            request[name] = value

    if args.operation == "generate":
        if args.moderation is not None:
            request["moderation"] = args.moderation
        response = client.images.generate(**request)
    else:
        if args.moderation is not None:
            request["extra_body"] = {"moderation": args.moderation}
        with ExitStack() as stack:
            request["image"] = [
                stack.enter_context(path.open("rb")) for path in image_paths
            ]
            if mask_path is not None:
                request["mask"] = stack.enter_context(mask_path.open("rb"))
            response = client.images.edit(**request)

    if downloader is None:
        parsed_base_url = urlsplit(base_url)
        assert parsed_base_url.hostname is not None
        private_origin = (
            parsed_base_url.hostname,
            parsed_base_url.port
            or (443 if parsed_base_url.scheme == "https" else 80),
        )
        downloader = lambda url: _download_url(url, private_origin)
    images = _response_bytes(response, downloader, expected_count)
    candidates = _candidate_paths(output, len(images))
    return _write_outputs(candidates, images)


def _sanitize_error(message: str) -> str:
    raw_key = os.environ.get(API_KEY_ENV, "")
    secrets = sorted({raw_key, raw_key.strip()} - {""}, key=len, reverse=True)
    for secret in secrets:
        message = message.replace(secret, "[REDACTED]")
    return message


def _serialize_error_detail(value: Any) -> str:
    try:
        if isinstance(value, bytes):
            text = value.decode("utf-8", errors="replace")
        elif isinstance(value, (dict, list)):
            text = json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)
        else:
            text = str(value)
    except Exception:
        return ""
    return text.strip()


def _safe_attribute(value: Any, name: str) -> Any:
    try:
        return getattr(value, name, None)
    except Exception:
        return None


def _format_error(error: Exception) -> str:
    message = _sanitize_error(_serialize_error_detail(error) or type(error).__name__)
    details: list[str] = [message]

    response = _safe_attribute(error, "response")
    body = _safe_attribute(error, "body")
    body_text = _serialize_error_detail(body) if body is not None else ""
    if not body_text:
        response_text = (
            _safe_attribute(response, "text") if response is not None else None
        )
        body_text = (
            _serialize_error_detail(response_text)
            if response_text is not None
            else ""
        )
    body_text = _sanitize_error(body_text)
    if body_text and body_text not in message:
        details.append(f"response_body={body_text[:MAX_ERROR_DETAIL_CHARACTERS]}")

    request_id = _safe_attribute(error, "request_id")
    if not request_id and response is not None:
        headers = _safe_attribute(response, "headers")
        try:
            request_id = headers.get("x-request-id") if headers is not None else None
        except Exception:
            request_id = None
    if request_id:
        request_id_text = _sanitize_error(_serialize_error_detail(request_id))
        if request_id_text and request_id_text not in message:
            details.append(f"request_id={request_id_text}")

    formatted = _sanitize_error(" | ".join(details))
    return formatted[:MAX_ERROR_DETAIL_CHARACTERS]


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        paths = execute(args)
    except Exception as error:
        print(f"error: {_format_error(error)}", file=sys.stderr)
        return 1

    for path in paths:
        print(f"OUTPUT={path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
