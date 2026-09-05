from __future__ import annotations

from contextlib import redirect_stderr
from datetime import datetime, timezone
import importlib.util
import io
import json
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import patch


SCRIPT_PATH = (
    Path(__file__).parents[1]
    / "skills"
    / "codex-chat-titles"
    / "scripts"
    / "chat_titles.py"
)
SPEC = importlib.util.spec_from_file_location("codex_chat_titles_cli", SCRIPT_PATH)
assert SPEC and SPEC.loader
cli = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(cli)


def timestamp(value: str) -> int:
    return int(datetime.fromisoformat(value).replace(tzinfo=timezone.utc).timestamp())


class FakeClient:
    def __init__(self, responses=None):
        self.responses = responses or {}
        self.calls = []
        self.names = {}

    def request(self, method, params):
        self.calls.append((method, params))
        response = self.responses.get(method)
        if callable(response):
            return response(params)
        if isinstance(response, list):
            return response.pop(0)
        if response is not None:
            return response
        if method == "thread/name/set":
            self.names[params["threadId"]] = params["name"]
            return {}
        if method == "thread/read":
            return {"thread": {"name": self.names.get(params["threadId"])}}
        raise AssertionError(f"Unexpected request: {method} {params}")


class TitleValidationTests(unittest.TestCase):
    def test_created_at_uses_asia_shanghai_date(self):
        created_at = timestamp("2026-09-02T16:30:00")
        cli.validate_title("0903｜优化｜批次文字显示", created_at)

    def test_rejects_updated_style_wrong_date(self):
        created_at = timestamp("2026-09-02T16:30:00")
        with self.assertRaisesRegex(ValueError, "does not match createdAt"):
            cli.validate_title("0902｜优化｜批次文字显示", created_at)

    def test_rejects_unknown_type_and_ascii_separator(self):
        created_at = timestamp("2026-09-03T02:00:00")
        for title in ("0903｜测试｜批次文字显示", "0903|优化|批次文字显示"):
            with self.subTest(title=title), self.assertRaisesRegex(
                ValueError, "Invalid title format"
            ):
                cli.validate_title(title, created_at)

    def test_accepts_spaces_between_chinese_and_ascii(self):
        created_at = timestamp("2026-09-03T02:00:00")
        for title in (
            "0903｜修复｜ChatGPT 启动失败",
            "0903｜功能｜配置 WebDAV 同步",
            "0903｜设计｜P3 安全方案",
            "0903｜研究｜SS2022 UDP 默认行为",
        ):
            with self.subTest(title=title):
                cli.validate_title(title, created_at)

    def test_rejects_missing_spaces_between_chinese_and_ascii(self):
        created_at = timestamp("2026-09-03T02:00:00")
        for title in (
            "0903｜修复｜ChatGPT启动失败",
            "0903｜功能｜配置WebDAV同步",
            "0903｜设计｜P3安全方案",
            "0903｜研究｜生成3个版本",
            "0903｜修复｜ChatGPT  启动失败",
            "0903｜功能｜配置\tWebDAV 同步",
        ):
            with self.subTest(title=title), self.assertRaisesRegex(
                ValueError, "separate Chinese text"
            ):
                cli.validate_title(title, created_at)


class AppServerDataTests(unittest.TestCase):
    def test_paged_requests_follow_next_cursor(self):
        client = FakeClient(
            {
                "thread/list": [
                    {"data": [{"id": "one"}], "nextCursor": "next"},
                    {"data": [{"id": "two"}], "nextCursor": None},
                ]
            }
        )
        records = cli._paged(client, "thread/list", {"archived": False})
        self.assertEqual(["one", "two"], [item["id"] for item in records])
        self.assertNotIn("cursor", client.calls[0][1])
        self.assertEqual("next", client.calls[1][1]["cursor"])

    def test_export_keeps_project_and_created_metadata(self):
        created_at = timestamp("2026-09-02T16:30:00")
        thread = {
            "id": "thread-1",
            "name": "优化批次文字显示",
            "preview": "优化列表中的批次文字",
            "createdAt": created_at,
            "projectId": "project-1",
            "cwd": "C:/work/example",
        }
        with patch.object(
            cli, "list_projects", return_value=[{"id": "project-1", "name": "Example"}]
        ), patch.object(cli, "list_threads", return_value=[thread]):
            payload = cli.build_export(FakeClient(), include_archived=False)
        item = payload["items"][0]
        self.assertEqual("0903", item["mmdd"])
        self.assertEqual("Example", item["projectName"])
        self.assertEqual("project", item["projectSource"])
        self.assertEqual("优化批次文字显示", item["displayName"])
        self.assertFalse(item["archived"])

    def test_export_uses_workspace_name_when_project_is_unassigned(self):
        created_at = timestamp("2026-09-03T02:00:00")
        thread = {
            "id": "thread-1",
            "name": "旧标题",
            "preview": "",
            "createdAt": created_at,
            "projectId": None,
            "cwd": "D:\\work\\MyProject",
        }
        with patch.object(cli, "list_projects", return_value=[]), patch.object(
            cli, "list_threads", return_value=[thread]
        ):
            item = cli.build_export(FakeClient(), include_archived=False)["items"][0]
        self.assertEqual("MyProject", item["projectName"])
        self.assertEqual("cwd", item["projectSource"])

    def test_context_keeps_prose_and_strips_code(self):
        client = FakeClient(
            {
                "thread/read": {
                    "thread": {
                        "name": "Untitled",
                        "preview": "请修复崩溃",
                        "turns": [
                            {
                                "items": [
                                    {
                                        "type": "userMessage",
                                        "content": [
                                            {
                                                "type": "text",
                                                "text": "请分析崩溃。```python\nraise Error\n```",
                                            }
                                        ],
                                    },
                                    {
                                        "type": "agentMessage",
                                        "text": "已经定位初始化错误。`raise Error`",
                                    },
                                ]
                            }
                        ],
                    }
                }
            }
        )
        item = cli.build_context(client, ["thread-1"])["items"][0]
        self.assertEqual("请分析崩溃。", item["userMessages"][0])
        self.assertEqual("已经定位初始化错误。", item["assistantWrapUp"][0])


class MappingTests(unittest.TestCase):
    def setUp(self):
        self.created_at = timestamp("2026-09-03T02:00:00")
        self.payload = {
            "schemaVersion": 1,
            "timezone": "Asia/Shanghai",
            "items": [
                {
                    "id": "thread-1",
                    "oldName": "旧标题",
                    "createdAt": self.created_at,
                    "newName": "0903｜修复｜启动时崩溃",
                }
            ],
        }

    def test_validate_mapping_checks_live_old_name_and_created_at(self):
        fresh = {
            "thread-1": {
                "id": "thread-1",
                "name": "旧标题",
                "createdAt": self.created_at,
            }
        }
        with patch.object(cli, "_fresh_threads", return_value=fresh):
            items = cli.validate_mapping(FakeClient(), self.payload)
        self.assertEqual("0903｜修复｜启动时崩溃", items[0]["newName"])

    def test_validate_mapping_rejects_title_changed_after_preview(self):
        fresh = {
            "thread-1": {
                "id": "thread-1",
                "name": "另一个标题",
                "createdAt": self.created_at,
            }
        }
        with patch.object(cli, "_fresh_threads", return_value=fresh), self.assertRaisesRegex(
            ValueError, "Title changed since preview"
        ):
            cli.validate_mapping(FakeClient(), self.payload)

    def test_apply_uses_only_name_set_and_verifies_read_back(self):
        client = FakeClient()
        result = cli.apply_mapping(client, self.payload["items"])
        self.assertEqual(
            {"changed": 1, "skipped": 0, "failed": 0, "failures": []}, result
        )
        self.assertEqual(
            ["thread/name/set", "thread/read"], [method for method, _ in client.calls]
        )

    def test_apply_rejects_changed_mapping_hash_before_connecting(self):
        with tempfile.TemporaryDirectory() as directory:
            mapping = Path(directory) / "map.json"
            mapping.write_text(json.dumps(self.payload, ensure_ascii=False), encoding="utf-8")
            args = SimpleNamespace(
                map=mapping,
                expected_sha256="0" * 64,
                confirm=True,
                timeout=1,
            )
            with self.assertRaisesRegex(ValueError, "SHA-256"), patch.object(
                cli, "AppServerClient"
            ) as client_class:
                cli.command_apply(args)
            client_class.assert_not_called()

    def test_apply_requires_explicit_confirm(self):
        args = SimpleNamespace(
            map=Path("unused.json"),
            expected_sha256="0" * 64,
            confirm=False,
            timeout=1,
        )
        with self.assertRaisesRegex(ValueError, "requires --confirm"):
            cli.command_apply(args)


if __name__ == "__main__":
    unittest.main()
