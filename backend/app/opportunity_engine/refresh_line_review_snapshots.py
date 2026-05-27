from __future__ import annotations

import argparse
import ast
import json
import platform
import queue
import subprocess
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .categories import Category, load_categories
from .line_review import build_line_review_sql
from .paths import ProjectPaths


POSTGRES_MCP_URL = "https://mcpservers.sunco.com/v2/servers/postgres/mcp/"
SNAPSHOT_FOLDER = Path("postgres_exports") / "line_reviews"


class McpRemoteClient:
    def __init__(self, server_url: str = POSTGRES_MCP_URL, timeout_seconds: int = 120) -> None:
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
                "clientInfo": {"name": "sunco-line-review-refresh", "version": "0.1"},
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
        parsed = ast.literal_eval(text)
        if not isinstance(parsed, list):
            raise RuntimeError(f"Unexpected SQL result shape: {type(parsed).__name__}")
        return [row for row in parsed if isinstance(row, dict)]


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def snapshot_path(paths: ProjectPaths, category: Category, generated_at: str) -> Path:
    stamp = generated_at.replace("-", "").replace(":", "").split("+", 1)[0].replace("T", "_")
    folder = paths.source_data / SNAPSHOT_FOLDER
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f"{category.slug}_line_review_{stamp}_postgres_mcp.json"


def write_snapshot(paths: ProjectPaths, category: Category, sql: str, rows: list[dict[str, Any]]) -> Path:
    generated_at = utc_now()
    payload = {
        "source_system": "postgres_mcp",
        "category_slug": category.slug,
        "category_run_name": category.run_name,
        "generated_at": generated_at,
        "row_count": len(rows),
        "source_reference": "Postgres MCP line-review refresh",
        "sql": sql,
        "rows": rows,
    }
    target = snapshot_path(paths, category, generated_at)
    target.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    return target


def select_categories(paths: ProjectPaths, names: list[str] | None) -> list[Category]:
    categories = load_categories(paths)
    if not names:
        return categories
    wanted = {name.strip().lower().replace(" ", "_").replace("/", "_") for name in names if name.strip()}
    selected = [
        category
        for category in categories
        if category.slug in wanted or category.run_name.lower() in wanted or category.name.lower() in wanted
    ]
    missing = sorted(wanted - {category.slug for category in selected})
    if missing:
        print(f"Warning: no active category matched: {', '.join(missing)}")
    return selected


def refresh_line_review_snapshots(
    paths: ProjectPaths,
    categories: list[Category],
    timeout_seconds: int = 180,
    dry_run: bool = False,
) -> dict[str, Any]:
    paths.ensure()
    results: list[dict[str, Any]] = []
    with McpRemoteClient(timeout_seconds=timeout_seconds) as client:
        for index, category in enumerate(categories, start=1):
            sql = build_line_review_sql(category, paths)
            print(f"[{index}/{len(categories)}] Refreshing {category.run_name}...")
            if dry_run:
                path = paths.cache / "ideation_data" / category.slug / "sql" / "line_review_postgres.sql"
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(sql, encoding="utf-8")
                results.append({"category": category.slug, "row_count": None, "snapshot": None, "status": "dry_run"})
                continue
            try:
                rows = client.execute_sql(sql, timeout_seconds=timeout_seconds)
                target = write_snapshot(paths, category, sql, rows)
                print(f"  wrote {len(rows)} row(s): {target}")
                results.append({"category": category.slug, "row_count": len(rows), "snapshot": str(target), "status": "ok"})
            except Exception as exc:
                print(f"  failed: {exc}")
                results.append({"category": category.slug, "row_count": 0, "snapshot": None, "status": "failed", "error": str(exc)})
    return {
        "generated_at": utc_now(),
        "categories": len(categories),
        "results": results,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Refresh approved Postgres line-review snapshots for Step 1.")
    parser.add_argument("--root", default=str(Path.cwd()), help="Ideation Development project root.")
    parser.add_argument("--category", action="append", help="Optional category slug/name. Repeat to refresh multiple categories.")
    parser.add_argument("--timeout-seconds", type=int, default=180, help="Per-category MCP query timeout.")
    parser.add_argument("--dry-run", action="store_true", help="Write SQL files only; do not query Postgres.")
    args = parser.parse_args()

    paths = ProjectPaths.from_root(args.root)
    categories = select_categories(paths, args.category)
    result = refresh_line_review_snapshots(paths, categories, args.timeout_seconds, args.dry_run)
    manifest = paths.source_data / SNAPSHOT_FOLDER / "line_review_refresh_manifest.json"
    manifest.parent.mkdir(parents=True, exist_ok=True)
    manifest.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(f"\nRefresh manifest:\n{manifest}")


if __name__ == "__main__":
    main()
