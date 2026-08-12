from __future__ import annotations

import argparse
import base64
from contextlib import ExitStack, nullcontext, redirect_stderr
import importlib.util
import io
import os
from pathlib import Path
import re
import struct
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import patch
import zlib

import httpx


SCRIPT_PATH = (
    Path(__file__).parents[1] / "skills" / "gpt-image" / "scripts" / "image_gen.py"
)
SPEC = importlib.util.spec_from_file_location("gpt_image_cli", SCRIPT_PATH)
assert SPEC and SPEC.loader
cli = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(cli)

API_KEY = "test-secret-key"
BASE_URL = "https://relay.example.com/v1/"
WEBP_BYTES = base64.b64decode(
    "UklGRiIAAABXRUJQVlA4IBYAAAAwAQCdASoBAAEADsD+JaQAA3AAAAAA"
)
JPEG_BYTES = b"\xff\xd8\xff\xd9"


def png_chunk(name, payload):
    checksum = zlib.crc32(name + payload) & 0xFFFFFFFF
    return struct.pack(">I", len(payload)) + name + payload + struct.pack(">I", checksum)


def png_scanlines(width, height, color_type=6, bit_depth=8, interlace=0):
    channels = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[color_type]
    passes = (
        ((0, 0, 1, 1),)
        if interlace == 0
        else (
            (0, 0, 8, 8),
            (4, 0, 8, 8),
            (0, 4, 4, 8),
            (2, 0, 4, 4),
            (0, 2, 2, 4),
            (1, 0, 2, 2),
            (0, 1, 1, 2),
        )
    )
    rows = bytearray()
    for x_start, y_start, x_step, y_step in passes:
        pass_width = (
            0 if width <= x_start else (width - x_start + x_step - 1) // x_step
        )
        pass_height = (
            0 if height <= y_start else (height - y_start + y_step - 1) // y_step
        )
        if not pass_width or not pass_height:
            continue
        row_bytes = (pass_width * channels * bit_depth + 7) // 8
        rows.extend((b"\x00" + b"\x00" * row_bytes) * pass_height)
    return bytes(rows)


def make_png(
    width=1,
    height=1,
    color_type=6,
    *,
    bit_depth=8,
    interlace=0,
    scanlines=None,
    idat_payloads=None,
    palette=None,
):
    ihdr = struct.pack(
        ">IIBBBBB", width, height, bit_depth, color_type, 0, 0, interlace
    )
    if scanlines is None:
        scanlines = png_scanlines(width, height, color_type, bit_depth, interlace)
    if idat_payloads is None:
        idat_payloads = [zlib.compress(scanlines)]
    return (
        b"\x89PNG\r\n\x1a\n"
        + png_chunk(b"IHDR", ihdr)
        + (png_chunk(b"PLTE", palette) if palette is not None else b"")
        + b"".join(png_chunk(b"IDAT", payload) for payload in idat_payloads)
        + png_chunk(b"IEND", b"")
    )


PNG_BYTES = make_png()


class FakeFactory:
    def __init__(self, response=None, error=None):
        self.response = response or SimpleNamespace(
            data=[SimpleNamespace(b64_json=base64.b64encode(PNG_BYTES).decode())]
        )
        self.error = error
        self.init_kwargs = None
        self.generate_kwargs = None
        self.edit_kwargs = None

    def __call__(self, **kwargs):
        self.init_kwargs = kwargs
        return SimpleNamespace(images=self)

    def generate(self, **kwargs):
        self.generate_kwargs = kwargs
        if self.error:
            raise self.error
        return self.response

    def edit(self, **kwargs):
        self.edit_kwargs = dict(kwargs)
        self.edit_kwargs["image"] = [Path(handle.name) for handle in kwargs["image"]]
        if "mask" in kwargs:
            self.edit_kwargs["mask"] = Path(kwargs["mask"].name)
        if self.error:
            raise self.error
        return self.response


class ImageGenTests(unittest.TestCase):
    def parse(self, *args):
        return cli.build_parser().parse_args(args)

    def env(self):
        return patch.dict(
            os.environ,
            {cli.API_KEY_ENV: API_KEY, cli.BASE_URL_ENV: BASE_URL},
            clear=True,
        )

    def patched_client(self, factory):
        stack = ExitStack()
        stack.enter_context(patch.object(cli, "OpenAI", factory))
        stack.enter_context(
            patch.object(
                cli,
                "DefaultHttpxClient",
                lambda **_kwargs: nullcontext(object()),
            )
        )
        return stack

    def test_generate_preserves_prompt_and_uses_one_fixed_model_call(self):
        prompt = "画一只猫，不要改写这个提示词。"
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env(), patch.object(
            cli, "_pictures_dir", return_value=Path(directory)
        ):
            paths = cli.execute(
                self.parse("generate", "--prompt", prompt),
                client_factory=factory,
            )

            self.assertEqual(len(paths), 1)
            self.assertEqual(paths[0].parent, Path(directory).resolve())
            self.assertEqual(paths[0].suffix, ".png")
            self.assertEqual(paths[0].read_bytes(), PNG_BYTES)
            self.assertEqual(
                factory.generate_kwargs,
                {"model": "gpt-image-2", "prompt": prompt},
            )
            self.assertEqual(factory.init_kwargs["base_url"], BASE_URL.rstrip("/"))
            self.assertEqual(factory.init_kwargs["max_retries"], 0)
            self.assertNotIn("http_client", factory.init_kwargs)
            self.assertIsNone(factory.edit_kwargs)

    def test_default_edit_request_contains_only_required_fields(self):
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            root = Path(directory)
            image = root / "input.png"
            image.write_bytes(PNG_BYTES)
            output = root / "edited.png"

            cli.execute(
                self.parse(
                    "edit",
                    "--prompt",
                    "preserve the subject",
                    "--image",
                    str(image),
                    "--out",
                    str(output),
                ),
                client_factory=factory,
            )

        self.assertEqual(factory.edit_kwargs["model"], "gpt-image-2")
        self.assertEqual(factory.edit_kwargs["prompt"], "preserve the subject")
        self.assertEqual(len(factory.edit_kwargs["image"]), 1)
        self.assertEqual(factory.edit_kwargs["image"][0].suffix, ".png")

    def test_explicit_generate_options_are_forwarded(self):
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            output = Path(directory) / "cat.webp"
            cli.execute(
                self.parse(
                    "generate",
                    "--prompt",
                    "draw a cat",
                    "--size",
                    "2048x1152",
                    "--quality",
                    "high",
                    "--output-format",
                    "webp",
                    "--output-compression",
                    "80",
                    "--background",
                    "opaque",
                    "--moderation",
                    "low",
                    "--n",
                    "1",
                    "--out",
                    str(output),
                ),
                client_factory=factory,
            )

        self.assertEqual(
            factory.generate_kwargs,
            {
                "model": "gpt-image-2",
                "prompt": "draw a cat",
                "size": "2048x1152",
                "quality": "high",
                "output_format": "webp",
                "output_compression": 80,
                "background": "opaque",
                "moderation": "low",
                "n": 1,
            },
        )

    def test_edit_passes_multiple_images_and_mask_without_input_fidelity(self):
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            root = Path(directory)
            first = root / "first.png"
            second = root / "second.png"
            mask = root / "mask.png"
            for path in (first, second, mask):
                path.write_bytes(PNG_BYTES)
            output = root / "edited.webp"

            cli.execute(
                self.parse(
                    "edit",
                    "--prompt",
                    "Keep both subjects",
                    "--image",
                    str(first),
                    "--image",
                    str(second),
                    "--mask",
                    str(mask),
                    "--output-format",
                    "webp",
                    "--output-compression",
                    "80",
                    "--moderation",
                    "low",
                    "--out",
                    str(output),
                ),
                client_factory=factory,
            )

            self.assertEqual(len(factory.edit_kwargs["image"]), 2)
            self.assertTrue(
                all(path.suffix == ".png" for path in factory.edit_kwargs["image"])
            )
            self.assertEqual(factory.edit_kwargs["mask"].suffix, ".png")
            self.assertEqual(factory.edit_kwargs["output_compression"], 80)
            self.assertEqual(factory.edit_kwargs["extra_body"], {"moderation": "low"})
            self.assertNotIn("moderation", factory.edit_kwargs)
            self.assertNotIn("input_fidelity", factory.edit_kwargs)
            self.assertIsNone(factory.generate_kwargs)

    def test_refuses_overwrite_before_creating_client(self):
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            output = Path(directory) / "existing.png"
            output.write_bytes(b"existing")
            with self.assertRaisesRegex(cli.CliError, "Refusing to overwrite"):
                cli.execute(
                    self.parse(
                        "generate", "--prompt", "test", "--out", str(output)
                    ),
                    client_factory=factory,
                )
            self.assertIsNone(factory.init_kwargs)
            self.assertEqual(output.read_bytes(), b"existing")

    def test_multiple_outputs_use_numbered_names(self):
        response = SimpleNamespace(
            data=[
                {"b64_json": base64.b64encode(b"one").decode()},
                {"b64_json": base64.b64encode(b"two").decode()},
            ]
        )
        factory = FakeFactory(response=response)
        with tempfile.TemporaryDirectory() as directory, self.env():
            output = Path(directory) / "result.png"
            paths = cli.execute(
                self.parse(
                    "generate",
                    "--prompt",
                    "two images",
                    "--n",
                    "2",
                    "--out",
                    str(output),
                ),
                client_factory=factory,
            )
            self.assertEqual([path.name for path in paths], ["result-1.png", "result-2.png"])
            self.assertEqual(paths[0].read_bytes(), b"one")
            self.assertEqual(paths[1].read_bytes(), b"two")
            self.assertEqual(factory.generate_kwargs["n"], 2)

    def test_unspecified_count_saves_all_returned_images(self):
        response = SimpleNamespace(
            data=[
                {"b64_json": base64.b64encode(b"one").decode()},
                {"b64_json": base64.b64encode(b"two").decode()},
            ]
        )
        factory = FakeFactory(response=response)
        with tempfile.TemporaryDirectory() as directory, self.env():
            pictures = Path(directory)
            with patch.object(cli, "_pictures_dir", return_value=pictures):
                paths = cli.execute(
                    self.parse("generate", "--prompt", "test"),
                    client_factory=factory,
                )

            resolved_pictures = pictures.resolve()
            self.assertEqual(
                [path.parent for path in paths],
                [resolved_pictures, resolved_pictures],
            )
            self.assertTrue(paths[0].stem.endswith("-1"))
            self.assertTrue(paths[1].stem.endswith("-2"))
            self.assertEqual(paths[0].stem[:-2], paths[1].stem[:-2])
            self.assertEqual(paths[0].read_bytes(), b"one")
            self.assertEqual(paths[1].read_bytes(), b"two")
            self.assertNotIn("n", factory.generate_kwargs)

    def test_url_response_is_rejected_without_downloading(self):
        factory = FakeFactory(response=SimpleNamespace(data=[{"url": "https://cdn.example/image"}]))

        with tempfile.TemporaryDirectory() as directory, self.env():
            output = Path(directory) / "result.png"
            with self.assertRaisesRegex(cli.CliError, "must contain b64_json"):
                cli.execute(
                    self.parse("generate", "--prompt", "test", "--out", str(output)),
                    client_factory=factory,
                )
            self.assertFalse(output.exists())

    def test_default_filename_uses_timestamp_milliseconds_and_uuid(self):
        with tempfile.TemporaryDirectory() as directory:
            with patch.object(cli, "_pictures_dir", return_value=Path(directory)):
                output = cli._default_output_path("png")
        self.assertRegex(
            output.name,
            re.compile(r"^\d{8}-\d{6}-\d{3}-[0-9a-f]{32}\.png$"),
        )

    def test_custom_size_constraints(self):
        self.assertEqual(cli._size("2048x1152"), "2048x1152")
        for invalid in ("1025x1024", "4096x1024", "3072x512", "640x640"):
            with self.subTest(invalid=invalid):
                with self.assertRaises(argparse.ArgumentTypeError):
                    cli._size(invalid)

    def test_image_count_is_limited_to_ten(self):
        self.assertEqual(cli._image_count("1"), 1)
        self.assertEqual(cli._image_count("10"), 10)
        for invalid in ("0", "11", "999999999"):
            with self.subTest(invalid=invalid):
                with self.assertRaises(argparse.ArgumentTypeError):
                    cli._image_count(invalid)

    def test_png_rejects_compression_before_api_call(self):
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            with self.assertRaisesRegex(cli.CliError, "jpeg or webp"):
                cli.execute(
                    self.parse(
                        "generate",
                        "--prompt",
                        "test",
                        "--output-compression",
                        "50",
                        "--out",
                        str(Path(directory) / "out.png"),
                    ),
                    client_factory=factory,
                )
        self.assertIsNone(factory.init_kwargs)

    def test_base_url_must_end_in_v1(self):
        with patch.dict(
            os.environ,
            {cli.API_KEY_ENV: API_KEY, cli.BASE_URL_ENV: "https://relay.example.com"},
            clear=True,
        ):
            with self.assertRaisesRegex(cli.CliError, "ending in /v1"):
                cli._load_config()

    def test_public_http_base_url_is_rejected_before_client_creation(self):
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, patch.dict(
            os.environ,
            {
                cli.API_KEY_ENV: API_KEY,
                cli.BASE_URL_ENV: "http://relay.example.com/v1",
            },
            clear=True,
        ):
            with self.assertRaisesRegex(cli.CliError, "must use HTTPS"):
                cli.execute(
                    self.parse(
                        "generate",
                        "--prompt",
                        "test",
                        "--out",
                        str(Path(directory) / "out.png"),
                    ),
                    client_factory=factory,
                )
        self.assertIsNone(factory.init_kwargs)

    def test_error_output_redacts_api_key(self):
        with tempfile.TemporaryDirectory() as directory, self.env():
            output = Path(directory) / "out.png"
            with self.patched_client(
                FakeFactory(error=RuntimeError(f"Authorization failed: {API_KEY}"))
            ):
                stderr = io.StringIO()
                with redirect_stderr(stderr):
                    code = cli.main(
                        ["generate", "--prompt", "test", "--out", str(output)]
                    )
        self.assertEqual(code, 1)
        self.assertNotIn(API_KEY, stderr.getvalue())
        self.assertIn("[REDACTED]", stderr.getvalue())

    def test_error_output_redacts_trimmed_api_key(self):
        spaced_key = f"  {API_KEY}  "
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "out.png"
            with patch.dict(
                os.environ,
                {cli.API_KEY_ENV: spaced_key, cli.BASE_URL_ENV: BASE_URL},
                clear=True,
            ), self.patched_client(
                FakeFactory(error=RuntimeError(f"Authorization failed: {API_KEY}"))
            ):
                stderr = io.StringIO()
                with redirect_stderr(stderr):
                    code = cli.main(
                        ["generate", "--prompt", "test", "--out", str(output)]
                    )
        self.assertEqual(code, 1)
        self.assertNotIn(API_KEY, stderr.getvalue())
        self.assertIn("[REDACTED]", stderr.getvalue())

    def test_error_output_includes_response_body_and_request_id(self):
        class RelayError(RuntimeError):
            body = {
                "error": {
                    "message": f"unsupported option for {API_KEY}",
                    "code": "invalid_request_error",
                }
            }
            request_id = "req_test_123"

        with tempfile.TemporaryDirectory() as directory, self.env():
            output = Path(directory) / "out.png"
            with self.patched_client(FakeFactory(error=RelayError("400 Bad Request"))):
                stderr = io.StringIO()
                with redirect_stderr(stderr):
                    code = cli.main(
                        ["generate", "--prompt", "test", "--out", str(output)]
                    )

        message = stderr.getvalue()
        self.assertEqual(code, 1)
        self.assertIn("400 Bad Request", message)
        self.assertIn("invalid_request_error", message)
        self.assertIn("req_test_123", message)
        self.assertNotIn(API_KEY, message)
        self.assertIn("[REDACTED]", message)

    def test_error_output_uses_response_fallback_details(self):
        response = SimpleNamespace(
            text='{"error":{"message":"relay rejected request"}}',
            headers={"x-request-id": "req_header_123"},
        )

        class RelayError(RuntimeError):
            pass

        error = RelayError("400 Bad Request")
        error.response = response
        with tempfile.TemporaryDirectory() as directory, self.env():
            output = Path(directory) / "out.png"
            with self.patched_client(FakeFactory(error=error)):
                stderr = io.StringIO()
                with redirect_stderr(stderr):
                    code = cli.main(
                        ["generate", "--prompt", "test", "--out", str(output)]
                    )

        message = stderr.getvalue()
        self.assertEqual(code, 1)
        self.assertIn("relay rejected request", message)
        self.assertIn("req_header_123", message)

    def test_error_output_redacts_key_before_truncating_details(self):
        boundary_key = "boundary-secret-1234567890"

        class RelayError(RuntimeError):
            body = "x" * (cli.MAX_ERROR_DETAIL_CHARACTERS - 8) + boundary_key

        with tempfile.TemporaryDirectory() as directory, patch.dict(
            os.environ,
            {cli.API_KEY_ENV: boundary_key, cli.BASE_URL_ENV: BASE_URL},
            clear=True,
        ):
            output = Path(directory) / "out.png"
            with self.patched_client(FakeFactory(error=RelayError("400"))):
                stderr = io.StringIO()
                with redirect_stderr(stderr):
                    code = cli.main(
                        ["generate", "--prompt", "test", "--out", str(output)]
                    )

        self.assertEqual(code, 1)
        self.assertNotIn("boundary", stderr.getvalue())

    def test_error_output_has_a_total_length_limit(self):
        long_message = "x" * 10_000

        class RelayError(RuntimeError):
            body = long_message

        formatted = cli._format_error(RelayError(long_message))

        self.assertEqual(len(formatted), cli.MAX_ERROR_DETAIL_CHARACTERS)

    def test_explicit_response_count_mismatch_is_checked_before_response_data(self):
        response = SimpleNamespace(
            data=[{"url": "https://cdn.example/one"}, {"url": "https://cdn.example/two"}]
        )
        with self.assertRaisesRegex(cli.CliError, "returned 2 images; expected 1"):
            cli._response_b64_items(response, 1)

    def test_unspecified_count_rejects_more_than_ten_images(self):
        response = SimpleNamespace(
            data=[{"url": f"https://cdn.example/{index}"} for index in range(11)]
        )
        with self.assertRaisesRegex(cli.CliError, "returned 11 images; at most 10"):
            cli._response_b64_items(response, None)

    def test_base64_is_decoded_in_chunks_without_a_total_output_limit(self):
        payload = b"a larger response"
        encoded = base64.b64encode(payload)
        with tempfile.TemporaryDirectory() as directory, patch.object(
            cli, "BASE64_DECODE_CHUNK_CHARACTERS", 4
        ):
            output = Path(directory) / "output.png"
            cli._write_outputs([output], [encoded])
            self.assertEqual(output.read_bytes(), payload)

    def test_invalid_second_base64_image_removes_all_partial_outputs(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            outputs = [root / "one.png", root / "two.png"]
            with self.assertRaisesRegex(cli.CliError, "invalid base64"):
                cli._write_outputs(
                    outputs,
                    [base64.b64encode(b"one").decode(), "not base64!"],
                )
            self.assertFalse(any(path.exists() for path in outputs))

    def test_nonfinal_base64_padding_is_rejected_and_output_is_removed(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(
            cli, "BASE64_DECODE_CHUNK_CHARACTERS", 4
        ):
            output = Path(directory) / "output.png"
            with self.assertRaisesRegex(cli.CliError, "invalid base64"):
                cli._write_outputs([output], ["TQ==TQ=="])
            self.assertFalse(output.exists())

    def test_failed_staging_write_leaves_no_final_output(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "broken.png"
            original_temporary_file = tempfile.TemporaryFile

            class BrokenWriter:
                def __init__(self, file_object):
                    self.file_object = file_object

                def __enter__(self):
                    return self

                def write(self, payload):
                    self.file_object.write(payload[:1])
                    raise OSError("simulated write failure")

                def __exit__(self, *args):
                    self.file_object.close()

            def broken_temporary_file(*args, **kwargs):
                return BrokenWriter(original_temporary_file(*args, **kwargs))

            with patch.object(cli.tempfile, "TemporaryFile", broken_temporary_file):
                with self.assertRaisesRegex(OSError, "simulated write failure"):
                    cli._write_outputs([output], [base64.b64encode(PNG_BYTES)])
            self.assertFalse(output.exists())
            self.assertEqual(list(Path(directory).glob("*.tmp")), [])

    def test_temporary_file_creation_failure_leaves_no_final_output(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "broken.png"
            with patch.object(
                cli.tempfile,
                "TemporaryFile",
                side_effect=OSError("simulated temporary file failure"),
            ):
                with self.assertRaisesRegex(
                    OSError, "simulated temporary file failure"
                ):
                    cli._write_outputs([output], [base64.b64encode(PNG_BYTES)])

            self.assertFalse(output.exists())
            self.assertEqual(list(Path(directory).glob("*.tmp")), [])

    def test_interrupted_staging_write_leaves_no_final_output(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "interrupted.png"
            original_temporary_file = tempfile.TemporaryFile

            class InterruptedWriter:
                def __init__(self, file_object):
                    self.file_object = file_object

                def __enter__(self):
                    return self

                def write(self, payload):
                    self.file_object.write(payload[:1])
                    raise KeyboardInterrupt

                def __exit__(self, *args):
                    self.file_object.close()

            def interrupted_temporary_file(*args, **kwargs):
                return InterruptedWriter(original_temporary_file(*args, **kwargs))

            with patch.object(
                cli.tempfile, "TemporaryFile", interrupted_temporary_file
            ):
                with self.assertRaises(KeyboardInterrupt):
                    cli._write_outputs(
                        [output], [base64.b64encode(PNG_BYTES).decode()]
                    )
            self.assertFalse(output.exists())
            self.assertEqual(list(Path(directory).glob("*.tmp")), [])

    def test_edit_rejects_more_than_sixteen_inputs_before_client_creation(self):
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            image = Path(directory) / "input.png"
            image.write_bytes(PNG_BYTES)
            arguments = ["edit", "--prompt", "test", "--out", str(Path(directory) / "out.png")]
            for _ in range(17):
                arguments.extend(("--image", str(image)))
            with self.assertRaisesRegex(cli.CliError, "At most 16"):
                cli.execute(self.parse(*arguments), client_factory=factory)
        self.assertIsNone(factory.init_kwargs)

    def test_mask_requires_png_first_input(self):
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            root = Path(directory)
            image = root / "input.jpg"
            mask = root / "mask.png"
            image.write_bytes(PNG_BYTES)
            mask.write_bytes(PNG_BYTES)
            with self.assertRaisesRegex(cli.CliError, "first input image must be PNG"):
                cli.execute(
                    self.parse(
                        "edit",
                        "--prompt",
                        "test",
                        "--image",
                        str(image),
                        "--mask",
                        str(mask),
                        "--out",
                        str(root / "out.png"),
                    ),
                    client_factory=factory,
                )
        self.assertIsNone(factory.init_kwargs)

    def test_mask_requires_matching_dimensions_before_client_creation(self):
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            root = Path(directory)
            image = root / "input.png"
            mask = root / "mask.png"
            image.write_bytes(make_png(2, 2, 2))
            mask.write_bytes(make_png(1, 1, 6))
            with self.assertRaisesRegex(cli.CliError, "matching dimensions"):
                cli.execute(
                    self.parse(
                        "edit",
                        "--prompt",
                        "test",
                        "--image",
                        str(image),
                        "--mask",
                        str(mask),
                        "--out",
                        str(root / "out.png"),
                    ),
                    client_factory=factory,
                )
        self.assertIsNone(factory.init_kwargs)

    def test_mask_requires_alpha_channel_before_client_creation(self):
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            root = Path(directory)
            image = root / "input.png"
            mask = root / "mask.png"
            image.write_bytes(make_png(1, 1, 2))
            mask.write_bytes(make_png(1, 1, 2))
            with self.assertRaisesRegex(cli.CliError, "alpha channel"):
                cli.execute(
                    self.parse(
                        "edit",
                        "--prompt",
                        "test",
                        "--image",
                        str(image),
                        "--mask",
                        str(mask),
                        "--out",
                        str(root / "out.png"),
                    ),
                    client_factory=factory,
                )
        self.assertIsNone(factory.init_kwargs)

    def test_truncated_png_is_rejected_before_client_creation(self):
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            root = Path(directory)
            image = root / "input.png"
            image.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 25)
            with self.assertRaisesRegex(cli.CliError, "valid PNG"):
                cli.execute(
                    self.parse(
                        "edit",
                        "--prompt",
                        "test",
                        "--image",
                        str(image),
                        "--out",
                        str(root / "out.png"),
                    ),
                    client_factory=factory,
                )
        self.assertIsNone(factory.init_kwargs)

    def test_extension_signature_mismatch_is_rejected_before_client(self):
        invalid_inputs = (
            ("input.jpg", b"not a JPEG"),
            ("input.webp", b"not a WebP"),
            ("input.png", b"not a PNG"),
        )
        for name, payload in invalid_inputs:
            with self.subTest(name=name), tempfile.TemporaryDirectory() as directory, self.env():
                factory = FakeFactory()
                root = Path(directory)
                image = root / name
                image.write_bytes(payload)
                with self.assertRaisesRegex(cli.CliError, "JPEG|WebP|PNG"):
                    cli.execute(
                        self.parse(
                            "edit",
                            "--prompt",
                            "test",
                            "--image",
                            str(image),
                            "--out",
                            str(root / "out.png"),
                        ),
                        client_factory=factory,
                    )
                self.assertIsNone(factory.init_kwargs)

    def test_matching_jpeg_webp_and_png_signatures_are_accepted(self):
        valid_inputs = (
            ("input.jpg", JPEG_BYTES),
            ("input.webp", WEBP_BYTES),
            ("input.png", make_png(color_type=3)),
        )
        for name, payload in valid_inputs:
            with self.subTest(name=name), tempfile.TemporaryDirectory() as directory, self.env():
                factory = FakeFactory()
                root = Path(directory)
                image = root / name
                image.write_bytes(payload)
                output = root / "out.png"
                cli.execute(
                    self.parse(
                        "edit",
                        "--prompt",
                        "test",
                        "--image",
                        str(image),
                        "--out",
                        str(output),
                    ),
                    client_factory=factory,
                )
                self.assertIsNotNone(factory.init_kwargs)

    def test_png_header_accepts_supported_color_types_and_interlace(self):
        grayscale = make_png(9, 2, 0, bit_depth=1)
        adam7 = make_png(9, 9, 6, interlace=1)

        for payload, expected in ((grayscale, (9, 2, 0)), (adam7, (9, 9, 6))):
            with self.subTest(expected=expected):
                stream = io.BytesIO(payload)
                self.assertEqual(
                    cli._png_header_info(
                        stream, "Input image", Path("input.png"), len(payload)
                    ),
                    expected,
                )

    def test_png_header_rejects_bad_ihdr_crc(self):
        payload = bytearray(PNG_BYTES)
        payload[29] ^= 0x01
        with self.assertRaisesRegex(cli.CliError, "valid PNG header"):
            cli._png_header_info(
                io.BytesIO(payload), "Input image", Path("input.png"), len(payload)
            )

    def test_input_file_is_opened_once_and_uploaded_from_snapshot(self):
        class ReadingFactory(FakeFactory):
            uploaded = None
            uploaded_handle = None

            def edit(self, **kwargs):
                self.uploaded_handle = kwargs["image"][0]
                self.uploaded = kwargs["image"][0].read()
                kwargs["image"][0].seek(0)
                return super().edit(**kwargs)

        factory = ReadingFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            root = Path(directory)
            image = root / "input.png"
            original_bytes = make_png(1, 1, 6)
            image.write_bytes(original_bytes)
            opened_handles = []
            original_open = open

            def track_open(file, *args, **kwargs):
                opened = original_open(file, *args, **kwargs)
                if Path(file) == image:
                    opened_handles.append(opened)
                return opened

            with patch("builtins.open", track_open):
                cli.execute(
                    self.parse(
                        "edit",
                        "--prompt",
                        "test",
                        "--image",
                        str(image),
                        "--out",
                        str(root / "out.png"),
                    ),
                    client_factory=factory,
                )

        self.assertEqual(factory.uploaded, original_bytes)
        self.assertEqual(len(opened_handles), 1)
        self.assertIsNot(opened_handles[0], factory.uploaded_handle)

    def test_edit_uploads_a_validated_snapshot_if_source_changes(self):
        class MutatingFactory(FakeFactory):
            uploaded = None

            def __init__(self, source):
                super().__init__()
                self.source = source

            def edit(self, **kwargs):
                self.source.write_bytes(b"not-an-image-secret")
                self.uploaded = kwargs["image"][0].read()
                kwargs["image"][0].seek(0)
                return super().edit(**kwargs)

        with tempfile.TemporaryDirectory() as directory, self.env():
            root = Path(directory)
            image = root / "input.png"
            image.write_bytes(PNG_BYTES)
            factory = MutatingFactory(image)
            cli.execute(
                self.parse(
                    "edit",
                    "--prompt",
                    "test",
                    "--image",
                    str(image),
                    "--out",
                    str(root / "out.png"),
                ),
                client_factory=factory,
            )

        self.assertEqual(factory.uploaded, PNG_BYTES)

    def test_publish_collision_never_removes_existing_file(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            output = root / "output.png"
            original_open = Path.open

            def colliding_open(path, mode="r", *args, **kwargs):
                if mode == "xb":
                    with original_open(path, "wb") as existing:
                        existing.write(b"other process")
                return original_open(path, mode, *args, **kwargs)

            with patch.object(cli.Path, "open", colliding_open):
                with self.assertRaisesRegex(cli.CliError, "Refusing to overwrite"):
                    cli._write_outputs([output], [base64.b64encode(PNG_BYTES)])

            self.assertEqual(output.read_bytes(), b"other process")
            self.assertEqual(list(root.glob("*.tmp")), [])

    def test_partial_publication_is_reported_and_published_file_is_preserved(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            outputs = [root / "one.png", root / "two.png"]
            original_open = Path.open
            calls = 0

            def fail_second_open(path, mode="r", *args, **kwargs):
                nonlocal calls
                if mode == "xb":
                    calls += 1
                    if calls == 2:
                        raise PermissionError("simulated publish failure")
                return original_open(path, mode, *args, **kwargs)

            with patch.object(cli.Path, "open", fail_second_open):
                with self.assertRaisesRegex(
                    cli.CliError, "Could not publish.*already published.*one.png"
                ):
                    cli._write_outputs(
                        outputs,
                        [base64.b64encode(b"one"), base64.b64encode(b"two")],
                    )

            self.assertEqual(outputs[0].read_bytes(), b"one")
            self.assertFalse(outputs[1].exists())
            self.assertEqual(list(root.glob("*.tmp")), [])

    def test_output_is_published_from_anonymous_staging(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "output.png"
            cli._write_outputs([output], [base64.b64encode(PNG_BYTES)])

            self.assertEqual(output.read_bytes(), PNG_BYTES)
            self.assertEqual(list(Path(directory).glob("*.tmp")), [])

    def test_interrupted_publication_reports_partial_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            output = root / "output.png"
            original_open = Path.open

            class InterruptedWriter:
                def __init__(self, file_object):
                    self.file_object = file_object

                def __enter__(self):
                    return self

                def write(self, payload):
                    self.file_object.write(payload[:1])
                    self.file_object.flush()
                    raise KeyboardInterrupt

                def __exit__(self, *args):
                    self.file_object.close()

            def interrupted_open(path, mode="r", *args, **kwargs):
                file_object = original_open(path, mode, *args, **kwargs)
                if mode == "xb":
                    return InterruptedWriter(file_object)
                return file_object

            with patch.object(cli.Path, "open", interrupted_open):
                with self.assertRaisesRegex(
                    cli.CliError, "interrupted.*partial file may remain.*output.png"
                ):
                    cli._write_outputs([output], [base64.b64encode(PNG_BYTES)])

            self.assertEqual(output.read_bytes(), PNG_BYTES[:1])
            self.assertEqual(list(root.glob("*.tmp")), [])

    def test_prompt_length_is_validated_before_client_creation(self):
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            output = Path(directory) / "out.png"
            with self.assertRaisesRegex(cli.CliError, "between 1 and 32000"):
                cli.execute(
                    self.parse(
                        "generate",
                        "--prompt",
                        "x" * 32_001,
                        "--out",
                        str(output),
                    ),
                    client_factory=factory,
                )
        self.assertIsNone(factory.init_kwargs)

    def test_platform_pictures_directory_is_absolute(self):
        self.assertTrue(cli._pictures_dir().is_absolute())

    def test_linux_pictures_directory_honors_xdg_config(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = root / "config"
            config.mkdir()
            pictures = root / "Custom Pictures"
            (config / "user-dirs.dirs").write_text(
                f'XDG_PICTURES_DIR="{pictures}"\n', encoding="utf-8"
            )
            with patch.dict(os.environ, {"XDG_CONFIG_HOME": str(config)}, clear=True):
                self.assertEqual(cli._linux_pictures_dir(root), pictures)

    def test_real_sdk_builds_one_edit_request_with_supported_extra_body(self):
        requests = []

        def handler(request):
            requests.append(request)
            return httpx.Response(
                200,
                json={
                    "data": [
                        {"b64_json": base64.b64encode(PNG_BYTES).decode()}
                    ]
                },
            )

        def factory(**kwargs):
            return cli.OpenAI(
                **kwargs,
                http_client=httpx.Client(transport=httpx.MockTransport(handler)),
            )

        with tempfile.TemporaryDirectory() as directory, self.env():
            root = Path(directory)
            image = root / "input.png"
            mask = root / "mask.png"
            image.write_bytes(PNG_BYTES)
            mask.write_bytes(PNG_BYTES)
            output = root / "output.png"
            cli.execute(
                self.parse(
                    "edit",
                    "--prompt",
                    "edit exactly once",
                    "--image",
                    str(image),
                    "--mask",
                    str(mask),
                    "--moderation",
                    "low",
                    "--out",
                    str(output),
                ),
                client_factory=factory,
            )

        self.assertEqual(len(requests), 1)
        self.assertEqual(requests[0].url.path, "/v1/images/edits")
        body = requests[0].content
        self.assertIn(b'name="model"', body)
        self.assertIn(b"gpt-image-2", body)
        self.assertIn(b'name="moderation"', body)
        self.assertIn(b"low", body)

    def test_real_sdk_does_not_retry_server_error(self):
        request_count = 0

        def handler(_request):
            nonlocal request_count
            request_count += 1
            return httpx.Response(
                500,
                json={"error": {"message": "server error", "type": "server_error"}},
            )

        def factory(**kwargs):
            return cli.OpenAI(
                **kwargs,
                http_client=httpx.Client(transport=httpx.MockTransport(handler)),
            )

        with tempfile.TemporaryDirectory() as directory, self.env():
            output = Path(directory) / "output.png"
            with self.assertRaises(Exception):
                cli.execute(
                    self.parse(
                        "generate", "--prompt", "test", "--out", str(output)
                    ),
                    client_factory=factory,
                )
        self.assertEqual(request_count, 1)

    def test_production_http_client_does_not_follow_redirects(self):
        requests = []
        clients = []

        def handler(request):
            requests.append(request)
            return httpx.Response(
                307,
                headers={"location": "http://127.0.0.1/private"},
            )

        def http_client_factory(**kwargs):
            client = httpx.Client(
                transport=httpx.MockTransport(handler),
                follow_redirects=kwargs["follow_redirects"],
            )
            clients.append(client)
            return client

        with tempfile.TemporaryDirectory() as directory, self.env(), patch.object(
            cli, "DefaultHttpxClient", http_client_factory
        ):
            output = Path(directory) / "output.png"
            with self.assertRaises(Exception) as caught:
                cli.execute(
                    self.parse(
                        "generate", "--prompt", "do not redirect", "--out", str(output)
                    )
                )

        self.assertEqual(getattr(caught.exception, "status_code", None), 307)
        self.assertEqual(len(requests), 1)
        self.assertEqual(requests[0].url.host, "relay.example.com")
        self.assertEqual(
            requests[0].extensions["timeout"]["read"], cli.REQUEST_TIMEOUT_SECONDS
        )
        self.assertEqual(len(clients), 1)
        self.assertTrue(clients[0].is_closed)


if __name__ == "__main__":
    unittest.main()
