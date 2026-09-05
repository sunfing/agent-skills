#!/usr/bin/env python3
"""Preview and apply Codex thread names through the local app-server."""

from __future__ import annotations

import argparse
from collections import deque
from datetime import datetime, timedelta, timezone
import hashlib
import hmac
import json
import ntpath
import os
from pathlib import Path
import queue
import re
import shutil
import subprocess
import sys
import threading
from typing import Any
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError


SCHEMA_VERSION = 1
TIMEZONE = "Asia/Shanghai"
TYPES = ("功能", "设计", "修复", "优化", "发布", "探索", "文档", "研究")
TITLE_RE = re.compile(
    rf"^(?P<date>\d{{4}})｜(?P<type>{'|'.join(TYPES)})｜(?P<topic>[^\r\n｜]+)$"
)
FENCE_RE = re.compile(r"```[\s\S]*?```|~~~[\s\S]*?~~~")
INLINE_CODE_RE = re.compile(r"`[^`\n]{1,300}`")
MEMORY_CITATION_RE = re.compile(r"<oai-mem-citation>[\s\S]*?</oai-mem-citation>")

try:
    SHANGHAI_TZ = ZoneInfo(TIMEZONE)
except ZoneInfoNotFoundError:
    # China has used UTC+08:00 without clock changes throughout the Codex era.
    SHANGHAI_TZ = timezone(timedelta(hours=8), TIMEZONE)


class AppServerError(RuntimeError):
    pass


def _resolve_codex_command() -> list[str]:
    configured = os.environ.get("CODEX_CLI_PATH")
    candidate = configured or shutil.which("codex") or shutil.which("codex.cmd")
    if not candidate:
        raise AppServerError(
            "Codex CLI not found. Put codex on PATH or set CODEX_CLI_PATH."
        )
    path = Path(candidate).expanduser()
    if configured and not path.is_file():
        raise AppServerError(f"CODEX_CLI_PATH does not exist: {path}")
    if os.name == "nt" and path.suffix.lower() == ".ps1":
        pwsh = shutil.which("pwsh")
        if not pwsh:
            raise AppServerError("pwsh is required to launch a Codex .ps1 command.")
        return [pwsh, "-NoProfile", "-File", str(path)]
    return [str(path)]


class AppServerClient:
    def __init__(self, timeout: float = 30.0):
        self.timeout = timeout
        self._next_id = 1
        self._messages: queue.Queue[dict[str, Any] | None] = queue.Queue()
        self._stderr: deque[str] = deque(maxlen=20)
        command = _resolve_codex_command() + ["app-server", "--listen", "stdio://"]
        flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        self.process = subprocess.Popen(
            command,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            creationflags=flags,
        )
        threading.Thread(target=self._read_stdout, daemon=True).start()
        threading.Thread(target=self._read_stderr, daemon=True).start()
        self.request(
            "initialize",
            {
                "clientInfo": {
                    "name": "codex_chat_titles",
                    "title": "Codex Chat Titles",
                    "version": "1.0.0",
                },
                "capabilities": {"experimentalApi": True},
            },
        )
        self.notify("initialized", {})

    def _read_stdout(self) -> None:
        assert self.process.stdout is not None
        for line in self.process.stdout:
            try:
                message = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(message, dict):
                self._messages.put(message)
        self._messages.put(None)

    def _read_stderr(self) -> None:
        assert self.process.stderr is not None
        for line in self.process.stderr:
            if line.strip():
                self._stderr.append(line.rstrip())

    def _send(self, message: dict[str, Any]) -> None:
        if self.process.poll() is not None:
            raise AppServerError(self._exit_message())
        assert self.process.stdin is not None
        self.process.stdin.write(json.dumps(message, ensure_ascii=False) + "\n")
        self.process.stdin.flush()

    def _exit_message(self) -> str:
        detail = "\n".join(self._stderr)
        suffix = f"\n{detail}" if detail else ""
        return f"Codex app-server exited with code {self.process.poll()}.{suffix}"

    def notify(self, method: str, params: dict[str, Any]) -> None:
        self._send({"method": method, "params": params})

    def request(self, method: str, params: dict[str, Any]) -> dict[str, Any]:
        request_id = self._next_id
        self._next_id += 1
        self._send({"method": method, "id": request_id, "params": params})
        while True:
            try:
                message = self._messages.get(timeout=self.timeout)
            except queue.Empty as exc:
                raise AppServerError(f"Timed out waiting for {method}.") from exc
            if message is None:
                raise AppServerError(self._exit_message())
            if message.get("id") == request_id:
                if "error" in message:
                    error = message["error"]
                    raise AppServerError(
                        f"{method} failed: {json.dumps(error, ensure_ascii=False)}"
                    )
                result = message.get("result")
                return result if isinstance(result, dict) else {}
            if "id" in message and "method" in message:
                self._send(
                    {
                        "id": message["id"],
                        "error": {
                            "code": -32601,
                            "message": "Client method not supported",
                        },
                    }
                )

    def close(self) -> None:
        if self.process.poll() is not None:
            return
        if self.process.stdin:
            self.process.stdin.close()
        try:
            self.process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            self.process.terminate()
            try:
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.process.kill()
                self.process.wait(timeout=5)

    def __enter__(self) -> "AppServerClient":
        return self

    def __exit__(self, exc_type, exc, traceback) -> None:
        self.close()


def _paged(client: AppServerClient, method: str, params: dict[str, Any]) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    cursor: str | None = None
    while True:
        page_params = {**params, "limit": 100}
        if cursor:
            page_params["cursor"] = cursor
        result = client.request(method, page_params)
        records.extend(item for item in result.get("data", []) if isinstance(item, dict))
        cursor = result.get("nextCursor")
        if not cursor:
            return records


def list_projects(client: AppServerClient) -> list[dict[str, Any]]:
    return _paged(client, "project/list", {})


def list_threads(
    client: AppServerClient, *, include_archived: bool
) -> list[dict[str, Any]]:
    threads = _paged(
        client,
        "thread/list",
        {"archived": False, "sortKey": "created_at", "sortDirection": "asc"},
    )
    if include_archived:
        threads.extend(
            _paged(
                client,
                "thread/list",
                {"archived": True, "sortKey": "created_at", "sortDirection": "asc"},
            )
        )
    return threads


def _created_iso(timestamp: int) -> str:
    return datetime.fromtimestamp(timestamp, timezone.utc).isoformat().replace("+00:00", "Z")


def _mmdd(timestamp: int) -> str:
    return datetime.fromtimestamp(timestamp, timezone.utc).astimezone(SHANGHAI_TZ).strftime(
        "%m%d"
    )


def _workspace_name(cwd: Any) -> str | None:
    if not isinstance(cwd, str) or not cwd.strip():
        return None
    return ntpath.basename(cwd.rstrip("/\\")) or None


def build_export(
    client: AppServerClient, *, include_archived: bool
) -> dict[str, Any]:
    projects = {item["id"]: item for item in list_projects(client)}
    active = list_threads(client, include_archived=False)
    archived_ids: set[str] = set()
    threads = active
    if include_archived:
        all_threads = list_threads(client, include_archived=True)
        active_ids = {item.get("id") for item in active}
        archived_ids = {item.get("id") for item in all_threads} - active_ids
        threads = all_threads
    items = []
    for thread in threads:
        created_at = thread.get("createdAt")
        thread_id = thread.get("id")
        if not isinstance(created_at, int) or not isinstance(thread_id, str):
            continue
        project = projects.get(thread.get("projectId"), {})
        explicit_name = thread.get("name")
        display_name = explicit_name or thread.get("preview") or ""
        cwd = thread.get("cwd")
        project_name = project.get("name") or _workspace_name(cwd)
        items.append(
            {
                "id": thread_id,
                "oldName": explicit_name if isinstance(explicit_name, str) else None,
                "displayName": display_name,
                "createdAt": created_at,
                "createdAtIso": _created_iso(created_at),
                "mmdd": _mmdd(created_at),
                "projectId": thread.get("projectId"),
                "projectName": project_name,
                "projectSource": "project" if project.get("name") else "cwd",
                "cwd": cwd,
                "archived": thread_id in archived_ids,
                "preview": thread.get("preview") or "",
            }
        )
    return {
        "schemaVersion": SCHEMA_VERSION,
        "generatedAt": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "timezone": TIMEZONE,
        "includeArchived": include_archived,
        "items": items,
    }


def _compact_prose(text: str, limit: int) -> str:
    text = MEMORY_CITATION_RE.sub(" ", FENCE_RE.sub(" ", text or ""))
    text = INLINE_CODE_RE.sub(" ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:limit]


def build_context(client: AppServerClient, thread_ids: list[str]) -> dict[str, Any]:
    items = []
    for thread_id in thread_ids:
        thread = client.request(
            "thread/read", {"threadId": thread_id, "includeTurns": True}
        ).get("thread", {})
        user_parts: list[str] = []
        assistant_parts: list[str] = []
        for turn in thread.get("turns", []):
            for item in turn.get("items", []):
                if item.get("type") == "userMessage":
                    for content in item.get("content", []):
                        if content.get("type") == "text" and content.get("text"):
                            user_parts.append(content["text"])
                elif item.get("type") == "agentMessage" and item.get("text"):
                    assistant_parts.append(item["text"])
        items.append(
            {
                "id": thread_id,
                "name": thread.get("name"),
                "preview": thread.get("preview") or "",
                "userMessages": [
                    _compact_prose(text, 600) for text in user_parts[:3]
                ],
                "assistantWrapUp": [
                    _compact_prose(text, 700) for text in assistant_parts[-2:]
                ],
            }
        )
    return {"schemaVersion": SCHEMA_VERSION, "items": items}


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    if not path.parent.is_dir():
        raise ValueError(f"Output directory does not exist: {path.parent}")
    with path.open("x", encoding="utf-8", newline="\n") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def _load_mapping(path: Path) -> tuple[dict[str, Any], str]:
    raw = path.read_bytes()
    digest = hashlib.sha256(raw).hexdigest()
    payload = json.loads(raw.decode("utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("Mapping root must be an object.")
    return payload, digest


def validate_title(title: str, created_at: int) -> None:
    match = TITLE_RE.fullmatch(title)
    if not match:
        raise ValueError(f"Invalid title format: {title}")
    if match.group("date") != _mmdd(created_at):
        raise ValueError(
            f"Title date {match.group('date')} does not match createdAt {_mmdd(created_at)}: {title}"
        )
    topic = match.group("topic").strip()
    if topic != match.group("topic") or len(topic) < 2 or len(topic) > 24:
        raise ValueError(f"Topic must contain 2-24 characters without edge spaces: {title}")


def _validate_mapping_shape(payload: dict[str, Any]) -> list[dict[str, Any]]:
    if payload.get("schemaVersion") != SCHEMA_VERSION:
        raise ValueError(f"schemaVersion must be {SCHEMA_VERSION}.")
    if payload.get("timezone") != TIMEZONE:
        raise ValueError(f"timezone must be {TIMEZONE}.")
    items = payload.get("items")
    if not isinstance(items, list):
        raise ValueError("items must be an array.")
    seen: set[str] = set()
    for index, item in enumerate(items):
        if not isinstance(item, dict):
            raise ValueError(f"items[{index}] must be an object.")
        if set(item) != {"id", "oldName", "createdAt", "newName"}:
            raise ValueError(
                f"items[{index}] must contain only id, oldName, createdAt, and newName."
            )
        thread_id = item["id"]
        if not isinstance(thread_id, str) or not thread_id:
            raise ValueError(f"items[{index}].id must be a non-empty string.")
        if thread_id in seen:
            raise ValueError(f"Duplicate thread id: {thread_id}")
        seen.add(thread_id)
        if item["oldName"] is not None and not isinstance(item["oldName"], str):
            raise ValueError(f"items[{index}].oldName must be a string or null.")
        if not isinstance(item["createdAt"], int):
            raise ValueError(f"items[{index}].createdAt must be an integer.")
        if not isinstance(item["newName"], str):
            raise ValueError(f"items[{index}].newName must be a string.")
        validate_title(item["newName"], item["createdAt"])
    return items


def _fresh_threads(client: AppServerClient) -> dict[str, dict[str, Any]]:
    return {
        item["id"]: item
        for item in list_threads(client, include_archived=True)
        if isinstance(item.get("id"), str)
    }


def validate_mapping(
    client: AppServerClient, payload: dict[str, Any]
) -> list[dict[str, Any]]:
    items = _validate_mapping_shape(payload)
    fresh = _fresh_threads(client)
    errors = []
    for item in items:
        current = fresh.get(item["id"])
        if current is None:
            errors.append(f"Thread not found: {item['id']}")
            continue
        if current.get("name") != item["oldName"]:
            errors.append(f"Title changed since preview: {item['id']}")
        if current.get("createdAt") != item["createdAt"]:
            errors.append(f"createdAt changed since preview: {item['id']}")
    if errors:
        raise ValueError("\n".join(errors))
    return items


def apply_mapping(
    client: AppServerClient, items: list[dict[str, Any]]
) -> dict[str, Any]:
    changed = 0
    skipped = 0
    failures = []
    for item in items:
        if item["oldName"] == item["newName"]:
            skipped += 1
            continue
        try:
            client.request(
                "thread/name/set",
                {"threadId": item["id"], "name": item["newName"]},
            )
            current = client.request(
                "thread/read", {"threadId": item["id"], "includeTurns": False}
            ).get("thread", {})
            if current.get("name") != item["newName"]:
                raise AppServerError("read-back title did not match")
            changed += 1
        except (AppServerError, KeyError, TypeError, ValueError) as exc:
            failures.append(
                {"id": item["id"], "newName": item["newName"], "error": str(exc)}
            )
    return {"changed": changed, "skipped": skipped, "failed": len(failures), "failures": failures}


def command_export(args: argparse.Namespace) -> None:
    with AppServerClient(args.timeout) as client:
        payload = build_export(client, include_archived=args.include_archived)
    _write_json(args.out, payload)
    print(json.dumps({"out": str(args.out), "threads": len(payload["items"])}, ensure_ascii=False))


def command_context(args: argparse.Namespace) -> None:
    with AppServerClient(args.timeout) as client:
        payload = build_context(client, args.id)
    _write_json(args.out, payload)
    print(json.dumps({"out": str(args.out), "threads": len(payload["items"])}, ensure_ascii=False))


def command_validate(args: argparse.Namespace) -> None:
    payload, digest = _load_mapping(args.map)
    with AppServerClient(args.timeout) as client:
        items = validate_mapping(client, payload)
    print(json.dumps({"valid": len(items), "sha256": digest}, ensure_ascii=False))


def command_apply(args: argparse.Namespace) -> None:
    if not args.confirm:
        raise ValueError("Apply requires --confirm after the user approves the preview table.")
    payload, digest = _load_mapping(args.map)
    if not hmac.compare_digest(digest, args.expected_sha256.lower()):
        raise ValueError("Mapping SHA-256 does not match the validated preview.")
    with AppServerClient(args.timeout) as client:
        items = validate_mapping(client, payload)
        result = apply_mapping(client, items)
    print(json.dumps(result, ensure_ascii=False))
    if result["failed"]:
        raise SystemExit(1)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Preview and rename Codex conversations through the local app-server."
    )
    parser.add_argument("--timeout", type=float, default=30.0)
    subparsers = parser.add_subparsers(dest="command", required=True)

    export = subparsers.add_parser("export")
    export.add_argument("--out", type=Path, required=True)
    export.add_argument("--include-archived", action="store_true")
    export.set_defaults(func=command_export)

    context = subparsers.add_parser("context")
    context.add_argument("--id", action="append", required=True)
    context.add_argument("--out", type=Path, required=True)
    context.set_defaults(func=command_context)

    validate = subparsers.add_parser("validate")
    validate.add_argument("--map", type=Path, required=True)
    validate.set_defaults(func=command_validate)

    apply = subparsers.add_parser("apply")
    apply.add_argument("--map", type=Path, required=True)
    apply.add_argument("--expected-sha256", required=True)
    apply.add_argument("--confirm", action="store_true")
    apply.set_defaults(func=command_apply)
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    try:
        args.func(args)
    except (AppServerError, OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
