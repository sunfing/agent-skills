from __future__ import annotations

import argparse
import base64
from contextlib import redirect_stderr
import importlib.util
import io
import os
from pathlib import Path
import re
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import patch

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
PNG_BYTES = b"\x89PNG\r\n\x1a\nmock"


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

    def test_generate_preserves_prompt_and_uses_one_fixed_model_call(self):
        prompt = "画一只猫，不要改写这个提示词。"
        factory = FakeFactory()
        with tempfile.TemporaryDirectory() as directory, self.env():
            output = Path(directory) / "cat.png"
            paths = cli.execute(
                self.parse("generate", "--prompt", prompt, "--out", str(output)),
                client_factory=factory,
            )

            self.assertEqual(paths, [output.resolve()])
            self.assertEqual(output.read_bytes(), PNG_BYTES)
            self.assertEqual(factory.generate_kwargs["prompt"], prompt)
            self.assertEqual(factory.generate_kwargs["model"], "gpt-image-2")
            self.assertEqual(factory.init_kwargs["base_url"], BASE_URL.rstrip("/"))
            self.assertEqual(factory.init_kwargs["max_retries"], 0)
            self.assertIsNone(factory.edit_kwargs)

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

            self.assertEqual(
                factory.edit_kwargs["image"], [first.resolve(), second.resolve()]
            )
            self.assertEqual(factory.edit_kwargs["mask"], mask.resolve())
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

    def test_url_response_uses_downloader(self):
        factory = FakeFactory(response=SimpleNamespace(data=[{"url": "https://cdn.example/image"}]))
        calls = []

        def download(url):
            calls.append(url)
            return b"downloaded"

        with tempfile.TemporaryDirectory() as directory, self.env():
            output = Path(directory) / "result.png"
            cli.execute(
                self.parse("generate", "--prompt", "test", "--out", str(output)),
                client_factory=factory,
                downloader=download,
            )
            self.assertEqual(calls, ["https://cdn.example/image"])
            self.assertEqual(output.read_bytes(), b"downloaded")

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

    def test_error_output_redacts_api_key(self):
        with tempfile.TemporaryDirectory() as directory, self.env():
            output = Path(directory) / "out.png"
            with patch.object(
                cli,
                "OpenAI",
                FakeFactory(error=RuntimeError(f"Authorization failed: {API_KEY}")),
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
            ), patch.object(
                cli,
                "OpenAI",
                FakeFactory(error=RuntimeError(f"Authorization failed: {API_KEY}")),
            ):
                stderr = io.StringIO()
                with redirect_stderr(stderr):
                    code = cli.main(
                        ["generate", "--prompt", "test", "--out", str(output)]
                    )
        self.assertEqual(code, 1)
        self.assertNotIn(API_KEY, stderr.getvalue())
        self.assertIn("[REDACTED]", stderr.getvalue())

    def test_response_count_mismatch_does_not_download_urls(self):
        response = SimpleNamespace(
            data=[{"url": "https://cdn.example/one"}, {"url": "https://cdn.example/two"}]
        )
        downloads = []
        with self.assertRaisesRegex(cli.CliError, "returned 2 images; expected 1"):
            cli._response_bytes(response, downloads.append, 1)
        self.assertEqual(downloads, [])

    def test_failed_write_removes_partial_output(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "broken.png"
            original_open = Path.open

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

            def broken_open(path, *args, **kwargs):
                return BrokenWriter(original_open(path, *args, **kwargs))

            with patch.object(Path, "open", broken_open):
                with self.assertRaisesRegex(OSError, "simulated write failure"):
                    cli._write_outputs([output], [PNG_BYTES])
            self.assertFalse(output.exists())

    def test_private_download_url_is_rejected_unless_it_matches_relay_origin(self):
        with self.assertRaisesRegex(cli.CliError, "non-public host"):
            cli._validate_download_url("http://127.0.0.1/image.png", None)
        cli._validate_download_url(
            "http://127.0.0.1:15721/image.png", ("127.0.0.1", 15721)
        )
        with self.assertRaisesRegex(cli.CliError, "non-public host"):
            cli._validate_download_url(
                "http://127.0.0.1:8080/image.png", ("127.0.0.1", 15721)
            )

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


if __name__ == "__main__":
    unittest.main()
