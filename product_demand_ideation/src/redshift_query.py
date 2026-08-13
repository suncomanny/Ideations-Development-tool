from __future__ import annotations

import os
import re
import sys
from pathlib import Path
from typing import Any


def is_redshift_mcp_source(source: str) -> bool:
    return source.strip().lower() == "redshift mcp"


def _ensure_backend_app_on_path() -> None:
    current = Path(__file__).resolve()
    for parent in current.parents:
        candidate = parent / "backend" / "app"
        if candidate.exists():
            if str(candidate) not in sys.path:
                sys.path.insert(0, str(candidate))
            return


def sanitize_redshift_error(exc: Exception) -> str:
    try:
        _ensure_backend_app_on_path()
        from opportunity_engine.db_mcp import sanitize_mcp_error

        text = sanitize_mcp_error(exc)
    except Exception:
        text = str(exc)
    return re.sub(
        r"(pwd|password|token|authorization|secret)=([^;)'\\s]+)",
        r"\1=***",
        text,
        flags=re.IGNORECASE,
    )


def _redshift_mcp_client(timeout_seconds: int, client_name: str):
    _ensure_backend_app_on_path()
    from opportunity_engine.db_mcp import REDSHIFT_MCP_URL, McpRemoteClient

    return McpRemoteClient(
        server_url=os.environ.get("REDSHIFT_MCP_URL") or REDSHIFT_MCP_URL,
        timeout_seconds=timeout_seconds,
        client_name=client_name,
    )


class RedshiftQueryClient:
    """Execute Redshift SQL through MCP only.

    The old local ODBC fallback is intentionally disabled for production Step 1
    because workstation DSN issues can silently produce empty or stale reports.
    """

    def __init__(
        self,
        timeout_seconds: int = 240,
        client_name: str = "sunco-ideation-redshift",
    ) -> None:
        self.timeout_seconds = timeout_seconds
        self.client_name = client_name
        self.connection_source = ""
        self._mcp_client: Any | None = None
        self._mcp_start_error: str | None = None

    def __enter__(self) -> "RedshiftQueryClient":
        try:
            self._mcp_client = _redshift_mcp_client(self.timeout_seconds, self.client_name)
            self._mcp_client.__enter__()
            self.connection_source = "Redshift MCP"
            return self
        except Exception as exc:
            self._mcp_client = None
            self._mcp_start_error = sanitize_redshift_error(exc)
            raise RuntimeError(
                "Redshift MCP is required for this workflow. Local Redshift ODBC fallback is disabled; "
                f"refresh MCP access and rerun. Detail: {self._mcp_start_error}"
            ) from exc

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._mcp_client is not None:
            self._mcp_client.__exit__(exc_type, exc, tb)

    @property
    def using_mcp(self) -> bool:
        return self._mcp_client is not None

    @property
    def mcp_start_error(self) -> str | None:
        return self._mcp_start_error

    def execute_sql(self, sql: str, timeout_seconds: int | None = None) -> list[dict[str, Any]]:
        timeout = timeout_seconds or self.timeout_seconds
        if self._mcp_client is not None:
            return self._mcp_client.execute_sql(sql, timeout_seconds=timeout)
        raise RuntimeError("Redshift MCP client is not connected.")


def execute_redshift_sql(
    sql: str,
    timeout_seconds: int = 240,
    client_name: str = "sunco-ideation-redshift",
) -> tuple[list[dict[str, Any]], str]:
    with RedshiftQueryClient(timeout_seconds=timeout_seconds, client_name=client_name) as client:
        return client.execute_sql(sql, timeout_seconds=timeout_seconds), client.connection_source
