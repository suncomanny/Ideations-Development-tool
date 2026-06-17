from __future__ import annotations

import ast
import datetime as datetime_module
import json
import platform
import queue
import subprocess
import threading
import time
from decimal import Decimal
from typing import Any


REDSHIFT_MCP_URL = "https://mcpservers.sunco.com/v2/servers/redshift/mcp/"


class McpRemoteClient:
    def __init__(self, server_url: str = REDSHIFT_MCP_URL, timeout_seconds: int = 120) -> None:
        self.server_url = server_url
        self.timeout_seconds = timeout_seconds
        self.process: subprocess.Popen[str] | None = None
        self.messages: "queue.Queue[dict[str, Any]]" = queue.Queue()
        self.stderr_lines: "queue.Queue[str]" = queue.Queue()
        self.next_id = 1

    def __enter__(self) -> "McpRemoteClient":
        npx = "npx.cmd" if platform.system().lower().startswith("win") else "npx"
        self.process = subprocess.Popen(
            [npx, "-y", "mcp-remote@latest", self.server_url],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )
        threading.Thread(target=self._read_stdout, daemon=True).start()
        threading.Thread(target=self._read_stderr, daemon=True).start()
        self._initialize()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self.process and self.process.poll() is None:
            self.process.terminate()
            try:
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.process.kill()

    def _read_stdout(self) -> None:
        assert self.process and self.process.stdout
        for line in self.process.stdout:
            line = line.strip()
            if not line:
                continue
            try:
                self.messages.put(json.loads(line))
            except json.JSONDecodeError:
                self.stderr_lines.put(f"Non-JSON stdout: {line[:500]}")

    def _read_stderr(self) -> None:
        assert self.process and self.process.stderr
        for line in self.process.stderr:
            line = line.strip()
            if line:
                self.stderr_lines.put(line)

    def _send(self, message: dict[str, Any]) -> None:
        if not self.process or not self.process.stdin:
            raise RuntimeError("MCP process is not running.")
        self.process.stdin.write(json.dumps(message) + "\n")
        self.process.stdin.flush()

    def _request(self, method: str, params: dict[str, Any] | None = None, timeout_seconds: int | None = None) -> dict[str, Any]:
        request_id = self.next_id
        self.next_id += 1
        self._send({"jsonrpc": "2.0", "id": request_id, "method": method, "params": params or {}})
        deadline = time.time() + (timeout_seconds or self.timeout_seconds)
        while time.time() < deadline:
            try:
                message = self.messages.get(timeout=0.25)
            except queue.Empty:
                if self.process and self.process.poll() is not None:
                    raise RuntimeError(f"MCP process exited with code {self.process.returncode}.")
                continue
            if message.get("id") == request_id:
                if "error" in message:
                    raise RuntimeError(json.dumps(message["error"], indent=2))
                return message.get("result") or {}
        recent_errors = []
        while not self.stderr_lines.empty() and len(recent_errors) < 12:
            recent_errors.append(self.stderr_lines.get_nowait())
        raise TimeoutError(f"MCP request timed out for {method}. Recent logs: {' | '.join(recent_errors)}")

    def _notify(self, method: str, params: dict[str, Any] | None = None) -> None:
        self._send({"jsonrpc": "2.0", "method": method, "params": params or {}})

    def _initialize(self) -> None:
        self._request(
            "initialize",
            {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "sunco-product-demand-ideation", "version": "0.1"},
            },
            timeout_seconds=60,
        )
        self._notify("notifications/initialized")

    def execute_sql(self, sql: str, timeout_seconds: int | None = None) -> list[dict[str, Any]]:
        result = self._request(
            "tools/call",
            {"name": "pg_execute_sql", "arguments": {"sql": sql}},
            timeout_seconds=timeout_seconds or self.timeout_seconds,
        )
        content = result.get("content") or result.get("result") or []
        if not content:
            return []
        first = content[0] if isinstance(content, list) else content
        text = first.get("text") if isinstance(first, dict) else str(first)
        if not text:
            return []
        if text.strip().lower().startswith("error:"):
            raise RuntimeError(text)
        try:
            parsed = ast.literal_eval(text)
        except (ValueError, SyntaxError):
            parsed = eval(  # noqa: S307 - MCP returns Python reprs for Decimal/date values.
                text,
                {"__builtins__": {}},
                {"Decimal": Decimal, "datetime": datetime_module},
            )
        if not isinstance(parsed, list):
            raise RuntimeError(f"Unexpected SQL result shape: {type(parsed).__name__}")
        return [row for row in parsed if isinstance(row, dict)]
